import {
  buildStatistics,
  deriveTypeOrder,
  type ExtractionResult,
  type Question,
} from '../domain/question';
import { CompositeExtractor } from '../extractors/composite-extractor';
import { resolvePrevTitle } from '../extractors/page-context';
import { FrameBridge, type ReceivedFrameResult } from '../infrastructure/frame-bridge';
import { stableHash } from '../utils/hash';
import { poll } from '../utils/async';
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
   * 采样等待题目就绪，题目一出现立即返回。
   *
   * 页面就绪的时刻并不确定：任务卡切换、答题页 AJAX 渲染都可能晚于我们的点击。
   * 「切一次卡 + 固定长等待」要么空等、要么在旧页面上点了空，因此改为按采样间隔
   * 反复确认——检测延迟收敛到一次采样间隔，期间还能通过 repair 修正被新页面
   * 重置的状态（例如任务卡回到默认的「视频」卡）。
   */
  async settleQuestions(options: QuestionSettleOptions = {}): Promise<ExtractionResult | null> {
    const previousFingerprint = options.previousFingerprint;
    return poll(
      () => {
        options.repair?.();
        return this.extract();
      },
      (result) => {
        if (!result || result.questions.length === 0) return false;
        return (
          previousFingerprint === undefined || this.fingerprint(result) !== previousFingerprint
        );
      },
      {
        timeoutMs: options.timeoutMs ?? DEFAULT_QUESTION_TIMEOUT_MS,
        intervalMs: options.intervalMs ?? QUESTION_SETTLE_INTERVAL_MS,
      },
    );
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
