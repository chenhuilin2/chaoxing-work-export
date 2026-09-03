import type {
  Question,
  QuestionOption,
  QuestionSource,
  QuestionType,
  RichContent,
} from '../domain/question';
import { stableHash } from '../utils/hash';
import { compareAnswers } from './answer-comparison';
import { isRichContentEmpty, richContentToText } from './rich-content';

export interface QuestionDraft {
  readonly number?: number;
  readonly type: QuestionType;
  readonly typeMeta?: string;
  readonly stem: RichContent;
  readonly options?: readonly QuestionOption[];
  readonly correctAnswer?: RichContent;
  readonly userAnswer?: RichContent;
  readonly analysis?: RichContent;
  readonly explicitWrong?: boolean;
  readonly source: QuestionSource;
  readonly chapterId?: string;
  readonly chapterTitle?: string;
}

export function createQuestion(draft: QuestionDraft): Question | null {
  if (isRichContentEmpty(draft.stem)) return null;
  const options = draft.options ?? [];
  const correctAnswer = draft.correctAnswer ?? [];
  const userAnswer = draft.userAnswer ?? [];
  const analysis = draft.analysis ?? [];
  const fingerprint = [
    draft.type,
    draft.number ?? '',
    richContentToText(draft.stem, (url) => url),
    ...options.map((option) => `${option.key}:${richContentToText(option.content, (url) => url)}`),
  ].join('|');

  return {
    id: `${draft.source.extractor}-${stableHash(fingerprint)}`,
    number: draft.number,
    type: draft.type,
    typeMeta: draft.typeMeta,
    stem: draft.stem,
    options,
    correctAnswer,
    userAnswer,
    analysis,
    isWrong: Boolean(draft.explicitWrong) || compareAnswers(draft.type, userAnswer, correctAnswer),
    source: draft.source,
    chapterId: draft.chapterId,
    chapterTitle: draft.chapterTitle,
  };
}
