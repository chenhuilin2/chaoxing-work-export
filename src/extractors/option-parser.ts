import type { QuestionOption, RichContent } from '../domain/question';
import { normalizeInlineWhitespace } from '../utils/text';
import { extractRichContent, isRichContentEmpty, stripOptionPrefix } from './rich-content';

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
  const contentElement =
    element.querySelector('.answer_p, .option-content, .option-text') ?? element;
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

/**
 * 选项容器兜底选择器：在各提取器自带的候选之后追加。
 *
 * 学习通 2026 版答题页的每个选项是 `<li class="font-cxsecret before-after" role="radio|checkbox"
 * onclick="addChoice(this)">`，选项文字在 `a.after` 里。类名随版本变化，但 `role` 与
 * `onclick` 这两个语义标记极稳定，因此单独放在最后兜底：正常页面走自带候选、行为不变，
 * 只有在类名全部失效（表现为「只提取到题干、提取不到选项」）时才由这里接手。
 */
export const FALLBACK_OPTION_SELECTORS = [
  '.Zy_ulTop li',
  '.qtDetail li',
  'li[onclick*="addChoice"]',
  'li[role="radio"]',
  'li[role="checkbox"]',
  // 已批阅视图的选项是 <li class="clearfix" role="option"><i class="fl">A、</i><a class="fl">…</a></li>，
  // 该类名同样随版本变化，只有 role="option" 稳定；它也是页面里最常见的通用 role，
  // 因此放在最末，仅当前面所有候选都落空时才接手。
  'li[role="option"]',
] as const;

export function parseOptions(
  container: ParentNode,
  selectors: readonly string[],
): QuestionOption[] {
  for (const selector of [...selectors, ...FALLBACK_OPTION_SELECTORS]) {
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
    container.querySelectorAll<HTMLInputElement>(
      'input[type="radio"]:checked, input[type="checkbox"]:checked',
    ),
  );
  if (selected.length === 0) return [];
  const keys = selected.map((input, index) => {
    const optionContainer = input.closest('.answerBg, li, .option, label') ?? input;
    return optionKeyFromContainer(optionContainer, index);
  });
  return [{ type: 'text', text: keys.join('') }];
}
