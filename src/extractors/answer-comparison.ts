import type { QuestionType, RichContent } from '../domain/question';
import { normalizeAnswer } from '../utils/text';
import { richContentToText } from './rich-content';

export function compareAnswers(
  type: QuestionType,
  userAnswer: RichContent,
  correctAnswer: RichContent,
): boolean {
  if (type === 'short-answer') return false;
  const user = normalizeAnswer(richContentToText(userAnswer));
  const correct = normalizeAnswer(richContentToText(correctAnswer));
  if (!user || !correct) return false;

  if (type === 'multiple-choice') {
    return [...user].sort().join('') !== [...correct].sort().join('');
  }
  return user !== correct;
}

export function hasExplicitWrongMarker(container: Element): boolean {
  return Boolean(
    container.querySelector(
      [
        '.colorRed',
        '.wrong',
        '.answer-wrong',
        '.is-wrong',
        '[data-correct="false"]',
        '[data-result="wrong"]',
      ].join(','),
    ),
  );
}
