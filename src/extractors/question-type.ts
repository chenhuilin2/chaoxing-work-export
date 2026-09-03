import type { QuestionType } from '../domain/question';
import { normalizeInlineWhitespace } from '../utils/text';

const LEGACY_NUMERIC_TYPES: Readonly<Record<string, QuestionType>> = {
  '0': 'single-choice',
  '1': 'multiple-choice',
  '2': 'fill-blank',
  '3': 'true-false',
  '4': 'short-answer',
};

const TYPE_PATTERNS: readonly [RegExp, QuestionType][] = [
  [/(?:多选(?:题)?|多项选择(?:题)?|不定项选择(?:题)?|multiple\s*choice)/iu, 'multiple-choice'],
  [/(?:单选(?:题)?|单项选择(?:题)?|选择题|single\s*choice)/iu, 'single-choice'],
  [/(?:填空(?:题)?|完形填空|fill(?:ing)?\s*(?:in\s*)?blank)/iu, 'fill-blank'],
  [/(?:判断(?:题)?|是非(?:题)?|true\s*or\s*false)/iu, 'true-false'],
  [
    /(?:简答(?:题)?|问答(?:题)?|论述(?:题)?|计算(?:题)?|名词解释|材料分析题|主观题|essay|short\s*answer)/iu,
    'short-answer',
  ],
];

export function detectQuestionType(
  ...values: readonly (string | null | undefined)[]
): QuestionType | null {
  const text = normalizeInlineWhitespace(values.filter(Boolean).join(' '));
  if (!text) return null;
  const numericType = LEGACY_NUMERIC_TYPES[text];
  if (numericType) return numericType;
  for (const [pattern, type] of TYPE_PATTERNS) {
    if (pattern.test(text)) return type;
  }
  return null;
}

export function inferQuestionType(container: Element): QuestionType | null {
  const explicit = detectQuestionType(
    container.getAttribute('typeName'),
    container.getAttribute('typename'),
    container.getAttribute('data-type-name'),
    container.getAttribute('data-question-type'),
    container.getAttribute('data-type'),
    container.querySelector('.type_tit, .newTestType, .question-type, [class*="typeName"]')
      ?.textContent,
  );
  if (explicit) return explicit;

  if (container.querySelector('input[type="checkbox"]')) return 'multiple-choice';
  if (container.querySelector('input[type="radio"]')) return 'single-choice';
  if (container.querySelector('textarea, [contenteditable="true"]')) return 'short-answer';

  const optionTexts = Array.from(
    container.querySelectorAll('.answerBg, .option, .Zy_ulTop li, .mark_letter li'),
  ).map((option) => normalizeInlineWhitespace(option.textContent ?? ''));
  if (
    optionTexts.length === 2 &&
    optionTexts.every((value) => /(?:正确|错误|对|错|true|false|√|×)/iu.test(value))
  ) {
    return 'true-false';
  }
  return null;
}
