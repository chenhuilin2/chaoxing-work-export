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
  '.correctAnswerBx .answerCon',
  '.correctAnswerContent',
  '.correct-answer',
  '.rightAnswer',
  '[data-role="correct-answer"]',
] as const;

const USER_SELECTORS = [
  '.mark_answer .mark_key .colorDeep .stuAnswerContent',
  '.mark_answer .mark_key .colorDeep',
  '.newAnswerBx .myAnswerBx .answerCon',
  '.myAnswerBx .answerCon',
  '.myAnswer .answerCon',
  '.my-answer',
  '[data-role="user-answer"]',
] as const;

const ANALYSIS_SELECTORS = [
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
