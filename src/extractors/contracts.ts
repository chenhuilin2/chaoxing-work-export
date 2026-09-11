import type { Question } from '../domain/question';

export interface ExtractorContext {
  readonly root: Document;
  readonly pageUrl: string;
  readonly title: string;
}

export interface ExtractionCandidate {
  readonly extractor: string;
  readonly confidence: number;
  readonly questions: readonly Question[];
}

export interface QuestionExtractor {
  readonly id: string;
  readonly confidence: number;
  supports(root: Document): boolean;
  extract(context: ExtractorContext): readonly Question[];
}
