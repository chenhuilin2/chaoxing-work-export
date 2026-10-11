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
  if (!hasGradingCorrectMarker(container)) return [];
  return userAnswer;
}

/**
 * 判定该题是否被批阅为「全对」。
 *
 * 学习通各版本的答对图标类名不统一（章节测验是 `.marking_dui`，试卷批阅页还有
 * `.marking_dui_true`、`.icon_dui` 等变体），逐个列举容易漏。这里的策略是：
 * 先认明确的类名，再用「名字里含 dui / correct / right 且不含 cuo / wrong」兜底，
 * **同时排除部分正确与错误标记**，宁可漏回填也不要把错答案当成标准答案。
 *
 * 注意：`.marking_cuo`（错误）与 `.marking_bandui`（部分正确）优先级最高，
 * 只要出现就判定为「非全对」，避免一个容器里同时渲染多次作答记录时误判。
 */
export function hasGradingCorrectMarker(container: Element): boolean {
  if (container.querySelector('.marking_cuo, .marking_bandui')) return false;

  const direct = container.querySelector(
    [
      '.marking_dui',
      '.marking_dui_true',
      '.icon_dui',
      '.icon-dui',
      '[data-result="right"]',
      '[data-result="correct"]',
      '[data-correct="true"]',
      // 2026 版试卷批阅页把批阅结果同时写在 aria-label 上
      // （`<span class="marking_dui" aria-label="答案正确">`），类名被改时这层语义标记通常还在
      '[aria-label*="答案正确"]',
      '[aria-label*="回答正确"]',
      '[aria-label*="完全正确"]',
    ].join(','),
  );
  if (direct) return true;

  // 兜底：class 名里带「dui / correct / right」的答对图标。
  //
  // 必须按**单个 class 词**匹配并要求词边界，不能拿整串 class 做子串匹配 —— 学习通布局里
  // `marginRight40`（见试卷批阅页的 `.mark_key > .colorDeep.marginRight40`）满地都是，
  // 子串匹配会把「未批阅」的题目误判成答对，进而把「我的答案」当正确答案回填。
  return Array.from(container.querySelectorAll<HTMLElement>('[class]')).some((element) => {
    const className = element.className;
    if (typeof className !== 'string') return false;
    if (/(cuo|wrong|bandui|error|fail)/i.test(className)) return false;
    return className
      .split(/\s+/)
      .some((token) => /(^|[-_])(dui|correct|right|true)([-_]|$)/i.test(token));
  });
}

/**
 * 取该题的正确答案：优先读页面上的「正确答案」文本，读不到时按批阅结果推断。
 *
 * 所有提取器都应当用这个函数而不是直接调 `extractCorrectAnswer` —— 批阅结果页
 * （试卷批阅 `.mark_item .questionLi`、章节测验 `.questionLi`、通用容器等）本身
 * 不含正确答案文本，漏接回填会让整份「答案汇总」变成「（未找到答案）」。
 */
export function resolveCorrectAnswer(container: Element): RichContent {
  const direct = extractCorrectAnswer(container);
  if (!isRichContentEmpty(direct)) return direct;
  return inferCorrectAnswerFromGrading(container, extractUserAnswer(container));
}
