import {
  buildStatistics,
  deriveTypeOrder,
  type ExtractionResult,
  type Question,
} from '../domain/question';
import { uniqueBy } from '../utils/array';
import type { ExtractionCandidate, QuestionExtractor } from './contracts';
import { GenericExerciseExtractor } from './generic-exercise-extractor';
import { MarkItemExtractor } from './mark-item-extractor';
import { resolvePageTitle, resolvePageUrl } from './page-context';
import { QuestionLiExtractor } from './question-li-extractor';
import { TiMuExtractor } from './timu-extractor';

export class CompositeExtractor {
  private readonly extractors: readonly QuestionExtractor[];

  constructor(
    extractors: readonly QuestionExtractor[] = [
      new TiMuExtractor(),
      new MarkItemExtractor(),
      new QuestionLiExtractor(),
      new GenericExerciseExtractor(),
    ],
  ) {
    this.extractors = extractors;
  }

  extract(root: Document): ExtractionResult | null {
    const title = resolvePageTitle(root);
    const pageUrl = resolvePageUrl(root);
    const candidates: ExtractionCandidate[] = [];

    for (const extractor of this.extractors) {
      if (!extractor.supports(root)) continue;
      const questions = uniqueBy(extractor.extract({ root, title, pageUrl }), (question) => {
        return question.id;
      });
      if (questions.length > 0) {
        candidates.push({ extractor: extractor.id, confidence: extractor.confidence, questions });
      }
    }

    const best = candidates.sort((left, right) => {
      const confidenceDifference = right.confidence - left.confidence;
      return confidenceDifference || right.questions.length - left.questions.length;
    })[0];
    if (!best) return null;

    return this.buildResult(title, pageUrl, best.extractor, best.questions);
  }

  private buildResult(
    title: string,
    sourceUrl: string,
    extractor: string,
    questions: readonly Question[],
  ): ExtractionResult {
    return {
      title,
      questions,
      typeOrder: deriveTypeOrder(questions),
      statistics: buildStatistics(questions),
      sourceUrl,
      extractor,
      extractedAt: new Date().toISOString(),
    };
  }
}
