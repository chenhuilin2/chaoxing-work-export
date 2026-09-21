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
import { incompleteChoiceCount, type ExtractionService } from './extraction-service';
import { QuestionTabGuard, TaskTabLocator } from './task-tab-locator';

// 单章等待上限：章节切换后要依次等卡片页重建 → 任务卡切换 → 答题页加载
const CHAPTER_SETTLE_TIMEOUT_MS = 20_000;

// 页面已停在目标章节、没有任务卡可切时，等待「选项中补上来」的额外上限
const PARTIAL_FILL_TIMEOUT_MS = 5_000;

// 收尾还原任务卡前的短暂等待：章节页面正在重建，任务卡栏属于新页面
const RESTORE_TAB_DELAY_MS = 400;

export interface ChapterProgress {
  readonly completed: number;
  readonly total: number;
  readonly chapter: ChapterDescriptor;
  readonly state: 'loading' | 'success' | 'failed';
  readonly message?: string;
}

export type ChapterProgressListener = (progress: ChapterProgress) => void;

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

        const result = await this.settleChapterResult({
          alreadyActive,
          hasTaskTabs,
          previousFingerprint,
        });

        if (!result) {
          onProgress({
            completed: index + 1,
            total: selected.length,
            chapter,
            state: 'failed',
            message: '未在限定时间内识别到题目',
          });
          continue;
        }

        const questions = result.questions.map((question, questionIndex) => {
          return this.attachChapter(question, chapter, questionIndex);
        });
        const chapterExtraction: ChapterExtraction = {
          id: chapter.id,
          title: chapter.title,
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

    if (chapterResults.length === 0) throw new Error('所选章节均未提取到题目');

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
   */
  private async settleChapterResult(options: {
    readonly alreadyActive: boolean;
    readonly hasTaskTabs: boolean;
    readonly previousFingerprint: string;
  }): Promise<ExtractionResult | null> {
    // 章节未切换且页面没有任务卡栏：结构不会自己变出题目，只取一次即可。
    // 例外是「选择题只渲染了题干、选项还没补上」，这种中间态要短暂等待补齐
    if (options.alreadyActive && !options.hasTaskTabs) {
      const immediate = this.extractionService.extract();
      if (!immediate || incompleteChoiceCount(immediate) === 0) return immediate;
      return (
        (await this.extractionService.settleQuestions({ timeoutMs: PARTIAL_FILL_TIMEOUT_MS })) ??
        immediate
      );
    }

    const guard = new QuestionTabGuard(this.taskTabs);
    return this.extractionService.settleQuestions({
      timeoutMs: CHAPTER_SETTLE_TIMEOUT_MS,
      previousFingerprint: options.alreadyActive ? undefined : options.previousFingerprint,
      repair: options.hasTaskTabs ? () => void guard.ensure() : undefined,
    });
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
