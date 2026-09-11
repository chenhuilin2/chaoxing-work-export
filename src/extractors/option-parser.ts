import type { QuestionOption, RichContent } from '../domain/question';
import { normalizeInlineWhitespace } from '../utils/text';
import {
  extractRichContent,
  isRichContentEmpty,
  stripOptionPrefix,
} from './rich-content';

export function parseLegacyOption(element: Element, index: number): QuestionOption | null {
  const parsed = stripOptionPrefix(extractRichContent(element));
  const key = parsed.key || String.fromCharCode(65 + index);
  if (isRichContentEmpty(parsed.content)) return null;
  return { key, content: parsed.content };
}

export function parseAnswerBackgroundOption(
  element: Element,
  index: number,
): QuestionOption | null {
  const keyElement = element.querySelector('.num_option, .option-letter, .option-index');
  const contentElement = element.querySelector('.answer_p, .option-content, .option-text') ?? element;
  const rawKey = normalizeInlineWhitespace(
    keyElement?.getAttribute('data') ??
      keyElement?.getAttribute('data-option') ??
      keyElement?.textContent ??
      String.fromCharCode(65 + index),
  );
  const key = rawKey.match(/[A-Z]/iu)?.[0]?.toUpperCase() ?? '';
  const parsed = stripOptionPrefix(extractRichContent(contentElement));
  if (isRichContentEmpty(parsed.content)) return null;
  return { key: key || parsed.key || String.fromCharCode(65 + index), content: parsed.content };
}

export function parseOptions(
  container: ParentNode,
  selectors: readonly string[],
): QuestionOption[] {
  for (const selector of selectors) {
    const elements = Array.from(container.querySelectorAll(selector));
    if (elements.length === 0) continue;
    const options = elements
      .map((element, index) => {
        return element.matches('.answerBg, [class*="answerBg"]')
          ? parseAnswerBackgroundOption(element, index)
          : parseLegacyOption(element, index);
      })
      .filter((option): option is QuestionOption => option !== null);
    if (options.length > 0) return options.sort((left, right) => left.key.localeCompare(right.key));
  }
  return [];
}

export function optionKeyFromContainer(container: Element, index: number): string {
  const direct = [
    container.getAttribute('data'),
    container.getAttribute('data-option'),
    container.getAttribute('data-key'),
    container.querySelector('.num_option, .option-letter, .mark_letter')?.textContent,
  ].find((value) => Boolean(value?.trim()));
  if (direct) {
    const normalized = normalizeInlineWhitespace(direct).match(/[A-Z]/iu)?.[0];
    if (normalized) return normalized.toUpperCase();
  }
  return String.fromCharCode(65 + index);
}

export function answersFromCheckedInputs(container: Element): RichContent {
  const selected = Array.from(
    container.querySelectorAll<HTMLInputElement>('input[type="radio"]:checked, input[type="checkbox"]:checked'),
  );
  if (selected.length === 0) return [];
  const keys = selected.map((input, index) => {
    const optionContainer = input.closest('.answerBg, li, .option, label') ?? input;
    return optionKeyFromContainer(optionContainer, index);
  });
  return [{ type: 'text', text: keys.join('') }];
}
