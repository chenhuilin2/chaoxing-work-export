import {
  buildStatistics,
  deriveTypeOrder,
  groupQuestions,
  type ChapterExtraction,
  type ExtractionResult,
  type Question,
} from '../domain/question';
import { shuffleCopy } from '../utils/array';

function shuffleByType(
  questions: readonly Question[],
  typeOrder: ExtractionResult['typeOrder'],
): Question[] {
  const groups = groupQuestions(questions);
  return typeOrder.flatMap((type) => shuffleCopy(groups[type]));
}

function transformChapter(chapter: ChapterExtraction, shuffle: boolean): ChapterExtraction {
  if (!shuffle) return chapter;
  const questions = shuffleByType(chapter.questions, chapter.typeOrder);
  return {
    ...chapter,
    questions,
    typeOrder: deriveTypeOrder(questions),
    statistics: buildStatistics(questions),
  };
}

export function transformResult(result: ExtractionResult, shuffle: boolean): ExtractionResult {
  if (!shuffle) return result;

  if (result.chapters && result.chapters.length > 0) {
    const chapters = result.chapters.map((chapter) => transformChapter(chapter, true));
    const questions = chapters.flatMap((chapter) => chapter.questions);
    return {
      ...result,
      chapters,
      questions,
      typeOrder: deriveTypeOrder(questions),
      statistics: buildStatistics(questions),
    };
  }

  const questions = shuffleByType(result.questions, result.typeOrder);
  return {
    ...result,
    questions,
    typeOrder: deriveTypeOrder(questions),
    statistics: buildStatistics(questions),
  };
}
