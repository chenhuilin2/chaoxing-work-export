import type { ExtractionResult } from '../domain/question';
import { CompositeExtractor } from '../extractors/composite-extractor';
import { debounce } from '../utils/async';
import { stableHash } from '../utils/hash';
import { Logger } from './logger';

const MESSAGE_TYPE = 'chaoxing-work-export:frame-result';
const REQUEST_TYPE = 'chaoxing-work-export:frame-request';
const PROTOCOL_VERSION = 1;

interface FrameRequest {
  readonly type: typeof REQUEST_TYPE;
  readonly protocolVersion: typeof PROTOCOL_VERSION;
}

interface FrameMessage {
  readonly type: typeof MESSAGE_TYPE;
  readonly protocolVersion: typeof PROTOCOL_VERSION;
  readonly frameUrl: string;
  readonly sentAt: number;
  readonly result: ExtractionResult;
}

export interface ReceivedFrameResult {
  readonly receivedAt: number;
  readonly frameUrl: string;
  readonly result: ExtractionResult;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isFrameRequest(value: unknown): value is FrameRequest {
  return (
    isRecord(value) &&
    value.type === REQUEST_TYPE &&
    value.protocolVersion === PROTOCOL_VERSION
  );
}

function isFrameMessage(value: unknown): value is FrameMessage {
  if (!isRecord(value)) return false;
  const result = value.result;
  return (
    value.type === MESSAGE_TYPE &&
    value.protocolVersion === PROTOCOL_VERSION &&
    typeof value.frameUrl === 'string' &&
    typeof value.sentAt === 'number' &&
    isRecord(result) &&
    Array.isArray(result.questions) &&
    result.questions.length <= 10_000 &&
    typeof result.title === 'string' &&
    typeof result.sourceUrl === 'string' &&
    typeof result.extractedAt === 'string'
  );
}

function isTrustedOrigin(origin: string): boolean {
  if (origin === window.location.origin) return true;
  try {
    const hostname = new URL(origin).hostname.toLowerCase();
    return hostname === 'chaoxing.com' || hostname.endsWith('.chaoxing.com');
  } catch {
    return false;
  }
}

export class FrameBridge {
  private readonly results = new Map<string, ReceivedFrameResult>();
  private readonly onMessage = (event: MessageEvent<unknown>): void => {
    if (!isTrustedOrigin(event.origin) || !isFrameMessage(event.data)) return;
    this.results.set(event.data.frameUrl, {
      receivedAt: Date.now(),
      frameUrl: event.data.frameUrl,
      result: event.data.result,
    });
  };

  start(): void {
    window.addEventListener('message', this.onMessage);
  }

  stop(): void {
    window.removeEventListener('message', this.onMessage);
  }

  latest(after = 0): ReceivedFrameResult | null {
    return (
      [...this.results.values()]
        .filter((entry) => entry.receivedAt >= after)
        .sort((left, right) => right.receivedAt - left.receivedAt)[0] ?? null
    );
  }

  requestRefresh(root: Document = document): void {
    const request: FrameRequest = { type: REQUEST_TYPE, protocolVersion: PROTOCOL_VERSION };
    const visited = new Set<Document>();
    const visit = (current: Document, depth: number): void => {
      if (visited.has(current) || depth > 5) return;
      visited.add(current);
      current.querySelectorAll<HTMLIFrameElement>('iframe').forEach((frame) => {
        try {
          frame.contentWindow?.postMessage(request, '*');
          if (frame.contentDocument) visit(frame.contentDocument, depth + 1);
        } catch {
          // Posting to a cross-origin WindowProxy is allowed; reading its document is not.
        }
      });
    };
    visit(root, 0);
  }

  clear(): void {
    this.results.clear();
  }
}

export class FrameAgent {
  private observer: MutationObserver | null = null;
  private lastFingerprint = '';
  private readonly logger = new Logger();
  private readonly sendDebounced: () => void;
  private readonly onRequest = (event: MessageEvent<unknown>): void => {
    if (event.source !== window.parent || !isTrustedOrigin(event.origin) || !isFrameRequest(event.data)) return;
    this.sendResult(true);
  };

  constructor(private readonly extractor = new CompositeExtractor()) {
    this.sendDebounced = debounce(() => this.sendResult(), 350);
  }

  start(): void {
    window.addEventListener('message', this.onRequest);
    this.sendResult();
    if (!document.body) {
      document.addEventListener('DOMContentLoaded', () => this.observe(), { once: true });
      return;
    }
    this.observe();
  }

  stop(): void {
    window.removeEventListener('message', this.onRequest);
    this.observer?.disconnect();
    this.observer = null;
  }

  private observe(): void {
    if (!document.body) return;
    this.observer?.disconnect();
    this.observer = new MutationObserver(() => this.sendDebounced());
    this.observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    window.setTimeout(() => this.sendResult(), 1_000);
    window.setTimeout(() => this.sendResult(), 3_000);
  }

  private sendResult(force = false): void {
    const result = this.extractor.extract(document);
    if (!result) return;
    const fingerprint = stableHash(
      JSON.stringify(
        result.questions.map((question) => ({
          id: question.id,
          correctAnswer: question.correctAnswer,
          userAnswer: question.userAnswer,
          analysis: question.analysis,
          isWrong: question.isWrong,
        })),
      ),
    );
    if (!force && fingerprint === this.lastFingerprint) return;
    this.lastFingerprint = fingerprint;
    const message: FrameMessage = {
      type: MESSAGE_TYPE,
      protocolVersion: PROTOCOL_VERSION,
      frameUrl: window.location.href,
      sentAt: Date.now(),
      result,
    };
    try {
      window.parent.postMessage(message, this.parentOrigin());
    } catch (error) {
      this.logger.debug('Unable to post iframe result', error);
    }
  }

  private parentOrigin(): string {
    try {
      const origin = new URL(document.referrer).origin;
      return isTrustedOrigin(origin) ? origin : '*';
    } catch {
      return '*';
    }
  }
}
