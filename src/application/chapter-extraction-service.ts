import {
  buildStatistics,
  deriveTypeOrder,
  type ChapterExtraction,
  type ExtractionResult,
  type Question,
} from '../domain/question';
import { resolvePageTitle } from '../extractors/page-context';
import { delay } from '../utils/async';
import { stableHash } from '../utils/hash';
import { ChapterLocator, type ChapterDescriptor } from './chapter-locator';
import { incompleteChoiceCount, hasQuestionDom, type ExtractionService } from './extraction-service';
import { QuestionTabGuard, TaskTabLocator } from './task-tab-locator';

// 单章等待上限：章节切换后要依次等卡片页重建 → 任务卡切换 → 答题页加载
const CHAPTER_SETTLE_TIMEOUT_MS = 20_000;

// 页面已停在目标章节、没有任务卡可切时，等待「选项中补上来」的额外上限
const PARTIAL_FILL_TIMEOUT_MS = 5_000;

// 收尾还原任务卡前的短暂等待：章节页面正在重建，任务卡栏属于新页面
const RESTORE_TAB_DELAY_MS = 400;

// 「本小节没有测验任务点」的确认时长：任务卡栏在切章瞬间会整体消失，
// 要求「栏存在且没有题目卡」持续这么久才算数，避免把重建过程中的中间态当成纯视频小节
const NO_QUESTION_TAB_GRACE_MS = 1_000;

export interface ChapterProgress {
  readonly completed: number;
  readonly total: number;
  readonly chapter: ChapterDescriptor;
  /** skipped：该小节只有视频/文档，没有测验任务点，没有题目可提 */
  readonly state: 'loading' | 'success' | 'failed' | 'skipped';
  readonly message?: string;
}

export type ChapterProgressListener = (progress: ChapterProgress) => void;

/**
 * 「本小节没有题目」探测器。
 *
 * 有些小节只放了视频/文档，任务卡栏里根本没有「章节测验」——题目永远不会出现，
 * 采样循环若照常等满超时，批量提取就会在这里卡住一下（用户感知为「停顿」）。
 * 任务卡栏在切章瞬间会整体消失（`#mainid` 先放 loading 图），
 * 因此要求「栏已存在、但没有题目卡」持续 `graceMs` 才算确认。
 *
 * 判据不能只看卡名：真实页面里存在「题目已经在 iframe 里渲染、任务卡栏里却没有
 * 测验卡」的情况（章节测验从别的入口进入、卡名被课程自定义改名等）。那种页面照常
 * 提取即可，因此再加一道保险 —— **页面里已经有题目 DOM 时绝不放弃**。
 */
export function createNoQuestionTabProbe(
  tabs: TaskTabLocator,
  graceMs = NO_QUESTION_TAB_GRACE_MS,
  questionsPresent: () => boolean = () => hasQuestionDom(),
): { check: (result?: ExtractionResult | null) => boolean; aborted: () => boolean } {
  let missingSince: number | null = null;
  let aborted = false;
  const check = (result?: ExtractionResult | null): boolean => {
    if (aborted) return true;
    // 栏还没建出来 / 题目卡又出现了 / 本轮已经读到题目 / 页面里已有题目结构：
    // 都不是「纯视频小节」，重新计时
    if (
      !tabs.present() ||
      tabs.questionTabIndex() !== null ||
      (result !== undefined && result !== null && result.questions.length > 0) ||
      questionsPresent()
    ) {
      missingSince = null;
      return false;
    }
    if (missingSince === null) missingSince = Date.now();
    aborted = Date.now() - missingSince >= graceMs;
    return aborted;
  };
  return { check, aborted: () => aborted };
}

/** Sequentially visits selected chapters and produces a single aggregate snapshot. */
export class ChapterExtractionService {
  constructor(
    private readonly extractionService: ExtractionService,
    private readonly locator = new ChapterLocator(),
    private readonly taskTabs = new TaskTabLocator(),
  ) {}

  listChapters(): ChapterDescriptor[] {
    return this.locator.list();
  }

  /** 当前激活章节的标识；供「切换章节自动提取」检测章节是否已被切走 */
  activeChapterKey(): string | null {
    return this.locator.activeKey();
  }

  async extractSelected(
    selectedIndexes: readonly number[],
    onProgress: ChapterProgressListener = () => undefined,
  ): Promise<ExtractionResult> {
    const chapters = this.locator.list();
    const selected = [...new Set(selectedIndexes)]
      .map((index) => chapters[index])
      .filter((chapter): chapter is ChapterDescriptor => Boolean(chapter));
    if (selected.length === 0) throw new Error('未选择可提取的章节');

    const originalIndex = this.locator.activeIndex();
    // 任务卡栏只在知识卡片页存在；批量开始前探测一次，避免逐章空等
    const hasTaskTabs = this.taskTabs.present();
    const originalTabIndex = hasTaskTabs ? this.taskTabs.activeIndex() : null;
    const chapterResults: ChapterExtraction[] = [];
    // 只有视频/文档的小节没有题目可提，单独计数，好在全部为空时给出准确提示
    let skippedCount = 0;
    let previousFingerprint = this.extractionService.fingerprint(this.extractionService.extract());

    try {
      for (let index = 0; index < selected.length; index += 1) {
        const chapter = selected[index] as ChapterDescriptor;
        onProgress({
          completed: index,
          total: selected.length,
          chapter,
          state: 'loading',
        });

        const alreadyActive = this.locator.activeIndex() === chapter.index;
        if (!alreadyActive && !this.locator.activate(chapter.index)) {
          onProgress({
            completed: index,
            total: selected.length,
            chapter,
            state: 'failed',
            message: '章节入口已变化或不可点击',
          });
          continue;
        }

        const outcome = await this.settleChapterResult({
          alreadyActive,
          hasTaskTabs,
          previousFingerprint,
        });

        if (!outcome.result) {
          if (outcome.skipped) skippedCount += 1;
          onProgress({
            completed: index + 1,
            total: selected.length,
            chapter,
            state: outcome.skipped ? 'skipped' : 'failed',
            message: outcome.skipped
              ? '该小节没有测验任务点（仅视频/文档），已跳过'
              : '未在限定时间内识别到题目',
          });
          continue;
        }

        const result = outcome.result;
        const questions = result.questions.map((question, questionIndex) => {
          return this.attachChapter(question, chapter, questionIndex);
        });
        const chapterExtraction: ChapterExtraction = {
          id: chapter.id,
          // 两级信息原样带进结果：导出要靠「小节编号 + 所属章」拼标题
          title: chapter.title,
          label: chapter.label,
          group: chapter.group,
          questions,
          typeOrder: deriveTypeOrder(questions),
          statistics: buildStatistics(questions),
          sourceUrl: result.sourceUrl,
          extractor: result.extractor,
        };
        chapterResults.push(chapterExtraction);
        previousFingerprint = this.extractionService.fingerprint(result);

        onProgress({
          completed: index + 1,
          total: selected.length,
          chapter,
          state: 'success',
        });
      }
    } finally {
      // 收尾：先还原章节，再还原任务卡（任务卡栏依赖章节页面重建后的结构）
      if (originalIndex !== null) {
        this.locator.activate(originalIndex);
        if (hasTaskTabs && originalTabIndex !== null) {
          await delay(RESTORE_TAB_DELAY_MS);
          this.taskTabs.activate(originalTabIndex);
        }
      }
    }

    if (chapterResults.length === 0) {
      throw new Error(
        skippedCount === selected.length
          ? '所选小节都没有测验任务点（仅视频/文档），没有题目可提取'
          : '所选章节均未提取到题目',
      );
    }

    const questions = chapterResults.flatMap((chapter) => chapter.questions);
    return {
      title: this.aggregateTitle(chapterResults),
      questions,
      typeOrder: deriveTypeOrder(questions),
      statistics: buildStatistics(questions),
      sourceUrl: window.location.href,
      extractor: [...new Set(chapterResults.map((chapter) => chapter.extractor))].join('+'),
      extractedAt: new Date().toISOString(),
      chapters: chapterResults,
    };
  }

  /**
   * 采样等待本章题目就绪，题目一出现立即返回。
   *
   * 章节切换会重建卡片页，新页面默认回到「视频」卡，只切一次的任务卡会随旧页面失效，
   * 因此不用「切一次卡 + 固定长等待」，而是在采样循环里每轮补切一次（限次 + 冷却）。
   * 章节刚切换时旧章节的题目 DOM 可能还在，必须等指纹变化才算本章内容，避免串页。
   * 只有视频/文档的小节永远等不到题目，由「无题目卡」探测器提前放弃。
   */
  private async settleChapterResult(options: {
    readonly alreadyActive: boolean;
    readonly hasTaskTabs: boolean;
    readonly previousFingerprint: string;
  }): Promise<{ readonly result: ExtractionResult | null; readonly skipped: boolean }> {
    // 章节未切换且页面没有任务卡栏：结构不会自己变出题目，只取一次即可。
    // 例外是「选择题只渲染了题干、选项还没补上」，这种中间态要短暂等待补齐
    if (options.alreadyActive && !options.hasTaskTabs) {
      const immediate = this.extractionService.extract();
      if (!immediate || incompleteChoiceCount(immediate) === 0) {
        return { result: immediate, skipped: false };
      }
      return {
        result:
          (await this.extractionService.settleQuestions({
            timeoutMs: PARTIAL_FILL_TIMEOUT_MS,
          })) ?? immediate,
        skipped: false,
      };
    }

    const guard = new QuestionTabGuard(this.taskTabs);
    const probe = createNoQuestionTabProbe(this.taskTabs);
    const result = await this.extractionService.settleQuestions({
      timeoutMs: CHAPTER_SETTLE_TIMEOUT_MS,
      previousFingerprint: options.alreadyActive ? undefined : options.previousFingerprint,
      repair: options.hasTaskTabs ? () => void guard.ensure() : undefined,
      // 页面没有任务卡栏时不启用：那种页面本来就不该按「没有测验卡」判断
      shouldAbort: options.hasTaskTabs ? probe.check : undefined,
    });
    return { result, skipped: result === null && probe.aborted() };
  }

  private attachChapter(
    question: Question,
    chapter: ChapterDescriptor,
    questionIndex: number,
  ): Question {
    return {
      ...question,
      id: stableHash(`${chapter.id}|${question.id}|${questionIndex}`),
      chapterId: chapter.id,
      chapterTitle: chapter.title,
    };
  }

  /**
   * 结果标题：单章节直接用章节名（文件名即章节名）；
   * 多章节用「页面标题 - N 个章节」，页面标题已内含 #prev_title 优先级。
   */
  private aggregateTitle(chapters: readonly ChapterExtraction[]): string {
    const scopeTitle = resolvePageTitle(document).trim();
    if (chapters.length === 1) {
      return chapters[0]?.title.trim() || scopeTitle || '学习通课程';
    }
    return [scopeTitle || '学习通课程', `${chapters.length} 个章节`].join(' - ');
  }
}
