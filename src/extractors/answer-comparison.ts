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
        // 章节测验「已完成 / 已批阅」视图不输出「正确答案」文本，只在每题下用图标表示批阅结果：
        // span.marking_dui 批阅正确、span.marking_cuo 批阅错误、span.marking_bandui 部分正确。
        // 这类页面没有可比较的答案文本，只能靠该标记判定错题，否则错题汇总会整页漏掉。
        // 部分正确同样需要复习，因此一并计入错题。
        '.marking_cuo',
        '.marking_bandui',
      ].join(','),
    ),
  );
}
