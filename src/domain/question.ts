export const QUESTION_TYPES = [
  'single-choice',
  'multiple-choice',
  'fill-blank',
  'true-false',
  'short-answer',
] as const;

export type QuestionType = (typeof QUESTION_TYPES)[number];

export const QUESTION_TYPE_LABELS: Readonly<Record<QuestionType, string>> = {
  'single-choice': '单选',
  'multiple-choice': '多选',
  'fill-blank': '填空',
  'true-false': '判断',
  'short-answer': '简答',
};

export const QUESTION_TYPE_LONG_LABELS: Readonly<Record<QuestionType, string>> = {
  'single-choice': '单项选择题',
  'multiple-choice': '多项选择题',
  'fill-blank': '填空题',
  'true-false': '判断题',
  'short-answer': '简答题',
};

export interface TextPart {
  readonly type: 'text';
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly subScript?: boolean;
  readonly superScript?: boolean;
}

export interface ImagePart {
  readonly type: 'image';
  readonly url: string;
  readonly alt: string;
}

export interface BreakPart {
  readonly type: 'break';
}

export type RichPart = TextPart | ImagePart | BreakPart;
export type RichContent = readonly RichPart[];

export interface QuestionOption {
  readonly key: string;
  readonly content: RichContent;
}

export interface QuestionSource {
  readonly extractor: string;
  readonly pageUrl: string;
  readonly selector?: string;
}

export interface Question {
  readonly id: string;
  readonly number?: number;
  readonly type: QuestionType;
  readonly typeMeta?: string;
  readonly stem: RichContent;
  readonly options: readonly QuestionOption[];
  readonly correctAnswer: RichContent;
  readonly userAnswer: RichContent;
  readonly analysis: RichContent;
  readonly isWrong: boolean;
  readonly source: QuestionSource;
  readonly chapterId?: string;
  readonly chapterTitle?: string;
}

export type QuestionGroups = Readonly<Record<QuestionType, readonly Question[]>>;

export interface ExtractionStatistics {
  readonly total: number;
  readonly wrong: number;
  readonly withCorrectAnswer: number;
  readonly withUserAnswer: number;
  readonly withAnalysis: number;
  readonly byType: Readonly<Record<QuestionType, number>>;
}

export interface ChapterExtraction {
  readonly id: string;
  readonly title: string;
  readonly questions: readonly Question[];
  readonly typeOrder: readonly QuestionType[];
  readonly statistics: ExtractionStatistics;
  readonly sourceUrl: string;
  readonly extractor: string;
}

export interface ExtractionResult {
  readonly title: string;
  readonly questions: readonly Question[];
  readonly typeOrder: readonly QuestionType[];
  readonly statistics: ExtractionStatistics;
  readonly sourceUrl: string;
  readonly extractor: string;
  readonly extractedAt: string;
  readonly chapters?: readonly ChapterExtraction[];
}

export function emptyTypeCounts(): Record<QuestionType, number> {
  return {
    'single-choice': 0,
    'multiple-choice': 0,
    'fill-blank': 0,
    'true-false': 0,
    'short-answer': 0,
  };
}

export function buildStatistics(questions: readonly Question[]): ExtractionStatistics {
  const byType = emptyTypeCounts();
  let wrong = 0;
  let withCorrectAnswer = 0;
  let withUserAnswer = 0;
  let withAnalysis = 0;

  for (const question of questions) {
    byType[question.type] += 1;
    if (question.isWrong) wrong += 1;
    if (question.correctAnswer.length > 0) withCorrectAnswer += 1;
    if (question.userAnswer.length > 0) withUserAnswer += 1;
    if (question.analysis.length > 0) withAnalysis += 1;
  }

  return {
    total: questions.length,
    wrong,
    withCorrectAnswer,
    withUserAnswer,
    withAnalysis,
    byType,
  };
}

export function deriveTypeOrder(questions: readonly Question[]): QuestionType[] {
  const seen = new Set<QuestionType>();
  const order: QuestionType[] = [];
  for (const question of questions) {
    if (!seen.has(question.type)) {
      seen.add(question.type);
      order.push(question.type);
    }
  }
  return order;
}

export function groupQuestions(
  questions: readonly Question[],
): Record<QuestionType, Question[]> {
  const groups: Record<QuestionType, Question[]> = {
    'single-choice': [],
    'multiple-choice': [],
    'fill-blank': [],
    'true-false': [],
    'short-answer': [],
  };
  for (const question of questions) groups[question.type].push(question);
  return groups;
}
