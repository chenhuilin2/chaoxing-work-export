import {
  buildStatistics,
  deriveTypeOrder,
  type ExtractionResult,
  type Question,
} from '../domain/question';
import { CompositeExtractor } from '../extractors/composite-extractor';
import { FrameBridge, type ReceivedFrameResult } from '../infrastructure/frame-bridge';
import { stableHash } from '../utils/hash';
import { poll } from '../utils/async';

export interface ExtractionWaitOptions {
  readonly afterTimestamp?: number;
  readonly previousFingerprint?: string;
  readonly timeoutMs?: number;
  readonly intervalMs?: number;
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
    return this.selectBest([
      ...this.collectDocumentCandidates(document),
      ...this.collectFrameCandidate(this.frameBridge.latest()),
    ]);
  }

  /** Extracts only documents that can be accessed synchronously from this window. */
  extractAccessibleDocuments(): ExtractionResult | null {
    return this.selectBest(this.collectDocumentCandidates(document));
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
    return stableHash(`${result.sourceUrl}|${result.title}|${identity}`);
  }

  async waitForChangedResult(options: ExtractionWaitOptions = {}): Promise<ExtractionResult | null> {
    const afterTimestamp = options.afterTimestamp ?? 0;
    const previousFingerprint = options.previousFingerprint ?? '';
    const timeoutMs = options.timeoutMs ?? 12_000;
    const intervalMs = options.intervalMs ?? 300;

    return poll(
      () => {
        const freshFrame = this.frameBridge.latest(afterTimestamp);
        return freshFrame?.result ?? this.extractAccessibleDocuments();
      },
      (result) => {
        if (!result || result.questions.length === 0) return false;
        const hasFreshFrame = this.frameBridge.latest(afterTimestamp) !== null;
        return hasFreshFrame || this.fingerprint(result) !== previousFingerprint;
      },
      { timeoutMs, intervalMs },
    );
  }

  private collectDocumentCandidates(root: Document): ExtractionCandidate[] {
    const candidates: ExtractionCandidate[] = [];
    const visited = new Set<Document>();

    const visit = (current: Document, depth: number): void => {
      if (visited.has(current) || depth > 5) return;
      visited.add(current);

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

      current.querySelectorAll<HTMLIFrameElement>('iframe').forEach((frame) => {
        try {
          if (frame.contentDocument) visit(frame.contentDocument, depth + 1);
        } catch {
          // Cross-origin frames are handled through FrameBridge/postMessage.
        }
      });
    };

    visit(root, 0);
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
