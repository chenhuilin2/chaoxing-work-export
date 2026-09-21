import type { RichContent } from '../domain/question';
import { allMatches, firstMatch } from '../utils/dom';
import {
  extractRichContent,
  isRichContentEmpty,
  joinRichContents,
  stripAnswerLabel,
  textContent,
} from './rich-content';
import { answersFromCheckedInputs } from './option-parser';

const CORRECT_SELECTORS = [
  '.mark_answer .mark_key .colorGreen .stuAnswerContent',
  '.mark_answer .mark_key .colorGreen',
  '.newAnswerBx .correctAnswerBx .answerCon',
  '.newAnswerBx .correctAnswer .answerCon',
  '.correctAnswerBx .answerCon',
  '.correctAnswerBx .correctAnswer',
  '.correctAnswerContent',
  '.correct-answer',
  '.rightAnswer',
  '[data-role="correct-answer"]',
] as const;

const USER_SELECTORS = [
  '.mark_answer .mark_key .colorDeep .stuAnswerContent',
  '.mark_answer .mark_key .colorDeep',
  '.newAnswerBx .myAnswerBx .answerCon',
  // 章节测验「已批阅」视图一次作答会渲染多组答案，此时用 myAllAnswerBx 包裹
  '.newAnswerBx .myAllAnswerBx .myAnswerBx',
  '.myAnswerBx .answerCon',
  '.myAnswer .answerCon',
  '.my-answer',
  '[data-role="user-answer"]',
] as const;

const ANALYSIS_SELECTORS = [
  '.newAnswerBx .answerKeyBx .answerCon',
  '.answerKeyBx .answerCon',
  '.newAnswerBx .analysisBx .answerCon',
  '.analysisBx .answerCon',
  '.mark_answer .mark_analysis',
  '.mark_answer .analysis',
  '.answerAnalysis',
  '.analysisContent',
  '.answer-analysis',
  '[data-role="analysis"]',
] as const;

function extractFirstContent(container: ParentNode, selectors: readonly string[]): RichContent {
  const element = firstMatch<Element>(container, selectors);
  return stripAnswerLabel(extractRichContent(element));
}

function extractFillAnswer(container: ParentNode, selector: string): RichContent {
  const elements = allMatches<Element>(container, [selector]);
  const contents = elements.map((element) => stripAnswerLabel(extractRichContent(element)));
  return joinRichContents(contents);
}

export function extractCorrectAnswer(container: Element): RichContent {
  const fill = extractFillAnswer(container, '.mark_answer .mark_fill.colorGreen dd');
  if (!isRichContentEmpty(fill)) return fill;
  return extractFirstContent(container, CORRECT_SELECTORS);
}

export function extractUserAnswer(container: Element): RichContent {
  const fill = extractFillAnswer(
    container,
    '.mark_answer .mark_fill .colorDeep, .mark_answer .mark_fill dd.colorDeep',
  );
  if (!isRichContentEmpty(fill)) return fill;

  const direct = extractFirstContent(container, USER_SELECTORS);
  if (!isRichContentEmpty(direct)) return direct;

  const checked = answersFromCheckedInputs(container);
  if (!isRichContentEmpty(checked)) return checked;

  const textInputs = Array.from(
    container.querySelectorAll<HTMLInputElement>(
      'input[type="text"], input:not([type]), .blankInput, [data-role="blank-input"]',
    ),
  )
    .map((input) => input.value.trim())
    .filter(Boolean);
  if (textInputs.length > 0) return textContent(textInputs.join('；'));

  const textarea = container.querySelector<HTMLTextAreaElement>('textarea');
  if (textarea?.value.trim()) return textContent(textarea.value);
  const editable = container.querySelector<HTMLElement>('[contenteditable="true"]');
  if (editable?.innerText.trim()) return textContent(editable.innerText);
  return [];
}

export function extractAnalysis(container: Element): RichContent {
  return extractFirstContent(container, ANALYSIS_SELECTORS);
}

/**
 * 从「教师批阅结果」推断正确答案。
 *
 * 章节测验「已完成 / 已批阅」视图只在每题下用图标给出批阅结果（`.marking_dui` 正确 /
 * `.marking_cuo` 错误 / `.marking_bandui` 部分正确），**不输出「正确答案」文本**。
 * 批阅正确说明该题我的答案与标准答案一致，此时用我的答案回填，否则导出的「答案汇总」
 * 会整页变成「（未找到答案）」，期末复习没有意义。
 *
 * 只在拿不到正确答案文本时介入（调用方负责判断），页面本身有正确答案时行为不变；
 * 批阅错误 / 部分正确时不回填，避免把错的答案当成标准答案。
 */
export function inferCorrectAnswerFromGrading(
  container: Element,
  userAnswer: RichContent,
): RichContent {
  if (isRichContentEmpty(userAnswer)) return [];
  if (!container.querySelector('.marking_dui')) return [];
  return userAnswer;
}
