import {
  buildStatistics,
  deriveTypeOrder,
  type ExtractionResult,
  type Question,
  type QuestionType,
} from '../domain/question';
import { CompositeExtractor } from '../extractors/composite-extractor';
import { resolvePrevTitle } from '../extractors/page-context';
import { FrameBridge, type ReceivedFrameResult } from '../infrastructure/frame-bridge';
import { delay } from '../utils/async';
import { stableHash } from '../utils/hash';
import { collectAccessibleDocuments } from '../utils/dom';

export interface QuestionSettleOptions {
  /** 等待上限；超时返回 null，不返回可能是上一章残留的内容 */
  readonly timeoutMs?: number;
  readonly intervalMs?: number;
  /** 传入时要求结果指纹发生变化，用于「章节刚切换」的场景 */
  readonly previousFingerprint?: string;
  /** 每轮采样前的修正动作，例如补切被新页面重置的任务卡 */
  readonly repair?: () => void;
}

// 采样间隔：只决定「题目出现」被检测到的延迟，越小越快（每轮含一次跨层提取）
const QUESTION_SETTLE_INTERVAL_MS = 120;

const DEFAULT_QUESTION_TIMEOUT_MS = 12_000;

// 结果需连续多帧完全一致才判定「渲染完毕」：答题页会分几批渲染题目，
// 只比对两帧会在两批之间误判为结束（表现为只拿到前面几道题）
const SETTLE_STABLE_MS = 600;

// 选择题确认渲染不出选项时的兜底：稳定这么久后按现状返回，避免空等超时
const PARTIAL_SETTLE_MS = 3_000;

// 判断「文档里是否已经有题目结构」用：只要命中就说明该文档可能还有选项在路上
const QUESTION_CONTAINER_SELECTOR = '.TiMu, .questionLi, .answerBg, .mark_item, .mark_name';

/** 需要选项才算「渲染完整」的题型 */
const CHOICE_TYPES: ReadonlySet<QuestionType> = new Set<QuestionType>([
  'single-choice',
  'multiple-choice',
]);

/**
 * 统计「渲染明显不完整」的选择题数量：选择题的题干与选项是同一批 HTML，
 * 只出现题干而没有选项，几乎只可能是页面还没渲染完。
 * 答题页中途被采样到时若直接采用，导出结果就会只剩题干。
 */
export function incompleteChoiceCount(result: ExtractionResult | null): number {
  if (!result) return 0;
  return result.questions.filter(
    (question) => CHOICE_TYPES.has(question.type) && question.options.length === 0,
  ).length;
}

/** 结果丰富度：题数优先，其次选项总数，用于在加载过程中保留最完整的一帧 */
function richnessOf(result: ExtractionResult): number {
  const options = result.questions.reduce((total, question) => total + question.options.length, 0);
  return result.questions.length * 1_000 + options;
}

interface ExtractionCandidate {
  readonly result: ExtractionResult;
  readonly sourcePriority: number;
  readonly timestamp: number;
}

/**
 * Coordinates all extraction sources visible to the current top-level page.
 *
 * A Chaoxing assignment may render directly in the document, in a same-origin
 * iframe, or in a cross-origin iframe. The service deliberately keeps these
 * transport concerns out of individual DOM extractors.
 */
export class ExtractionService {
  constructor(
    private readonly extractor = new CompositeExtractor(),
    private readonly frameBridge = new FrameBridge(),
  ) {}

  extract(): ExtractionResult | null {
    return this.applyScopeTitle(
      this.selectBest([
        ...this.collectDocumentCandidates(document),
        ...this.collectFrameCandidate(this.frameBridge.latest()),
      ]),
    );
  }

  /** Extracts only documents that can be accessed synchronously from this window. */
  extractAccessibleDocuments(): ExtractionResult | null {
    return this.applyScopeTitle(this.selectBest(this.collectDocumentCandidates(document)));
  }

  /** Extracts only the current document, without recursively reading iframes. */
  extractCurrentDocument(root: Document = document): ExtractionResult | null {
    return this.extractor.extract(root);
  }

  fingerprint(result: ExtractionResult | null): string {
    if (!result) return 'empty';
    const identity = result.questions
      .map((question) => `${question.id}:${question.type}`)
      .join('|');
    // 标题不参与指纹：外层壳的章节名（.prev_title）会比题目 DOM 更早更新，
    // 若纳入指纹，chapter 提取会把上一章残留的题目误判成本章的新内容
    return stableHash(`${result.sourceUrl}|${identity}`);
  }

  /**
   * 用外层壳（学生学习页面）的章节名覆盖标题。
   *
   * 章节页的题目在嵌套 iframe 里，嵌套文档只能解析出自身标题（「章节测验 待完成」），
   * 而调用方需要的是章节名，因此在此统一覆盖；无 #prev_title / .prev_title 时行为不变。
   */
  private applyScopeTitle(result: ExtractionResult | null): ExtractionResult | null {
    if (!result) return null;
    const scopeTitle = resolvePrevTitle(document);
    if (!scopeTitle || scopeTitle === result.title) return result;
    return { ...result, title: scopeTitle };
  }

  /**
   * 采样等待题目「渲染完整」，而不是等到第一眼看见题目就收工。
   *
   * 页面就绪的时刻并不确定：任务卡切换、答题页 AJAX 渲染都可能晚于我们的点击。
   * 「切一次卡 + 固定长等待」要么空等、要么在旧页面上点了空，因此改为按采样间隔反复确认，
   * 期间通过 repair 修正被新页面重置的状态（例如任务卡回到默认的「视频」卡）。
   *
   * 但只判断「有没有题目」并不够：答题页会分几批渲染，中途被采样到时可能只渲染出题干、
   * 选项还没补上。因此这里做了两件事：
   * 1. 结果要连续多帧完全一致才认为渲染结束（`SETTLE_STABLE_MS`）；
   * 2. 选择题缺选项视为「仍在渲染」，除非页面稳定很久仍是如此（`PARTIAL_SETTLE_MS` 兜底，
   *    避免个别确实没有选项的页面白等满整段超时）；若承载题目的文档还在 loading，
   *    说明 HTML 还没解析完，此时不允许按兜底提前返回，继续等（`deadline` 仍然兜底）；
   * 3. 全程记录「最丰富的一帧」，超时或兜底返回它，绝不返回比中间帧更差的结果。
   */
  async settleQuestions(options: QuestionSettleOptions = {}): Promise<ExtractionResult | null> {
    const previousFingerprint = options.previousFingerprint;
    const intervalMs = options.intervalMs ?? QUESTION_SETTLE_INTERVAL_MS;
    const deadline = Date.now() + (options.timeoutMs ?? DEFAULT_QUESTION_TIMEOUT_MS);

    let best: ExtractionResult | null = null;
    let bestRichness = Number.NEGATIVE_INFINITY;
    let seenFingerprint = '';
    let stableMs = 0;

    for (;;) {
      options.repair?.();
      const result = this.extract();

      if (result && result.questions.length > 0) {
        const fingerprint = this.fingerprint(result);
        // 指纹未变化说明页面还停在上一个内容上（章节尚未切过去），这一类帧不参与判定
        const changed = previousFingerprint === undefined || fingerprint !== previousFingerprint;
        if (changed) {
          const richness = richnessOf(result);
          if (richness > bestRichness) {
            best = result;
            bestRichness = richness;
          }

          stableMs = fingerprint === seenFingerprint ? stableMs + intervalMs : 0;
          seenFingerprint = fingerprint;

          const incomplete = incompleteChoiceCount(result) > 0;
          // 答题页是一整个大 HTML，浏览器按到达顺序解析：题干（.Zy_TItle，结构里在前）
          // 可能已经入 DOM，而同一题的选项（ul.Zy_ulTop，结构里在后）还没解析到。
          // 这种「题干有了、选项还在路上」的中间态不满足「稳定 3 秒就收工」的前提，
          // 只要承载题目的文档还在 loading 就继续等（最终仍由 deadline 兜底）。
          const waitingForOptions = incomplete && this.hasLoadingQuestionDocument();
          if (
            stableMs >= SETTLE_STABLE_MS &&
            (!incomplete || (stableMs >= PARTIAL_SETTLE_MS && !waitingForOptions))
          ) {
            return incomplete ? (best ?? result) : result;
          }
        }
      }

      if (Date.now() >= deadline) return best;
      await delay(intervalMs);
    }
  }

  /**
   * 是否还有「承载题目的文档」处于加载中。
   *
   * 只有「选择题缺选项」的结果会用到它：分批到达的 HTML 会让题干先出现、选项后到，
   * 这时页面既没有变化也没有选项，但选项并非不会来，所以不能按「稳定 3 秒」收工。
   * 结果本身是完整的（选择题都有选项）时完全不参与判定，正常页面行为不变。
   */
  private hasLoadingQuestionDocument(): boolean {
    for (const { document: current } of collectAccessibleDocuments(document)) {
      if (current.readyState !== 'loading') continue;
      try {
        if (current.querySelector(QUESTION_CONTAINER_SELECTOR)) return true;
      } catch {
        // 不可访问的文档直接跳过
      }
    }
    return false;
  }

  private collectDocumentCandidates(root: Document): ExtractionCandidate[] {
    const candidates: ExtractionCandidate[] = [];

    // 跨层收集交给公共遍历工具，和任务卡/章节定位共用同一套 iframe 规则
    for (const { document: current, depth } of collectAccessibleDocuments(root)) {
      try {
        const result = this.extractor.extract(current);
        if (result) {
          candidates.push({
            result,
            sourcePriority: Math.max(1, 100 - depth * 10),
            timestamp: Date.parse(result.extractedAt) || Date.now(),
          });
        }
      } catch {
        // One malformed document must not prevent other frames from being inspected.
      }
    }

    return candidates;
  }

  private collectFrameCandidate(envelope: ReceivedFrameResult | null): ExtractionCandidate[] {
    if (!envelope) return [];
    return [
      {
        result: envelope.result,
        sourcePriority: 95,
        timestamp: envelope.receivedAt,
      },
    ];
  }

  private selectBest(candidates: readonly ExtractionCandidate[]): ExtractionResult | null {
    if (candidates.length === 0) return null;

    const sorted = [...candidates].sort((left, right) => {
      const questionDifference = right.result.questions.length - left.result.questions.length;
      if (questionDifference !== 0) return questionDifference;
      // 题数相同时优先选选项更完整的候选：同一份题目可能有「只渲染了题干」的中间态，
      // 若按来源优先级挑选，半渲染的候选会把带选项的完整候选顶掉
      const richnessDifference = richnessOf(right.result) - richnessOf(left.result);
      if (richnessDifference !== 0) return richnessDifference;
      const priorityDifference = right.sourcePriority - left.sourcePriority;
      if (priorityDifference !== 0) return priorityDifference;
      return right.timestamp - left.timestamp;
    });

    const best = sorted[0];
    if (!best) return null;

    // A page can expose the same question in nested accessible documents. Merge
    // candidates with the winning title only when their question identities differ.
    const questions = new Map<string, Question>();
    for (const candidate of sorted) {
      if (candidate.result.title !== best.result.title && best.result.questions.length > 0) continue;
      for (const question of candidate.result.questions) questions.set(question.id, question);
    }
    const merged = [...questions.values()];

    return {
      ...best.result,
      questions: merged,
      typeOrder: deriveTypeOrder(merged),
      statistics: buildStatistics(merged),
    };
  }
}
