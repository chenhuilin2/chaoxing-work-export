import type { Question, QuestionOption, QuestionType } from '../domain/question';
import { parseLeadingNumber, textOf } from '../utils/dom';
import { hasExplicitWrongMarker } from './answer-comparison';
import {
  extractAnalysis,
  extractCorrectAnswer,
  extractUserAnswer,
  inferCorrectAnswerFromGrading,
} from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseLegacyOption, parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { resolveQuestionStem } from './question-stem';
import { detectQuestionType, inferQuestionType } from './question-type';
import { isRichContentEmpty } from './rich-content';

/**
 * 选项容器候选（当前版本 → 历史版本），命中即用；解析与诊断报告共用同一份，
 * 新增候选只改这里，不要在两处各写一套。
 */
export const TIMU_OPTION_SELECTORS = [
  '.Zy_ulTop.qtDetail > li',
  '.Zy_ulTop > li',
  '.answerBg',
  '.option-list > li',
] as const;

/** 「单题块」候选：选项被挪出题目容器时，通常仍留在这几个块里 */
const QUESTION_BLOCK_SELECTOR = '.singleQuesId, .questionLi, .mark_item';

/** 页面给每道题标的 qid（形如 405842140）：跨容器找选项时用它精确定位，不会串题 */
function resolveQuestionId(container: Element): string {
  const block = container.closest(QUESTION_BLOCK_SELECTOR);
  const candidates = [
    container.getAttribute('qid'),
    container.getAttribute('data-qid'),
    block?.getAttribute('data'),
    block?.getAttribute('id'),
    container.getAttribute('id'),
    container.querySelector('input[id^="answer"]')?.getAttribute('id'),
    container.querySelector('li[qid]')?.getAttribute('qid'),
  ];
  for (const raw of candidates) {
    // 要求至少 4 位数字，避免把 <i class="fl">1</i> 这类题号当成 qid
    const matched = (raw ?? '').match(/\d{4,}/u);
    if (matched) return matched[0];
  }
  return '';
}

/** 按 qid 取该题的选项节点（只认页面自己标的 qid） */
function parseOptionsByQuestionId(scope: ParentNode, qid: string): QuestionOption[] {
  for (const selector of [`li[qid="${qid}"]`, `li[data-qid="${qid}"]`]) {
    let elements: Element[];
    try {
      elements = Array.from(scope.querySelectorAll(selector));
    } catch {
      continue;
    }
    if (elements.length === 0) continue;
    const options = elements
      .map((element, index) => parseLegacyOption(element, index))
      .filter((option): option is QuestionOption => option !== null);
    if (options.length > 0) return options.sort((left, right) => left.key.localeCompare(right.key));
  }
  return [];
}

/**
 * 选项解析：先取题目容器内的候选，容器内一个都取不到时向外层逐级兜底。
 *
 * 已批阅视图（章节测验「已完成」）的部分版本把选项放在题目容器之外（同级表单 /
 * 外层单题块里），此时容器内所有候选都会落空，表现为「题目都提取到了、选项一个都没有」
 * ——即使选项明明显示在页面上。兜底只在容器内确实取不到时触发，命中即用，正常页面不受影响。
 */
function resolveOptions(container: Element): QuestionOption[] {
  const direct = parseOptions(container, TIMU_OPTION_SELECTORS);
  if (direct.length > 0) return direct;

  const block = container.closest(QUESTION_BLOCK_SELECTOR);
  for (const scope of [block, container.parentElement]) {
    if (!scope || scope === container) continue;
    // 只有这一层确实只装了一题时才用类名候选，避免把邻题的选项串进来
    if (scope.querySelectorAll('.TiMu').length > 1) continue;
    const options = parseOptions(scope, TIMU_OPTION_SELECTORS);
    if (options.length > 0) return options;
  }

  const qid = resolveQuestionId(container);
  if (qid) {
    for (const scope of [block, container.parentElement, container.ownerDocument]) {
      if (!scope) continue;
      const options = parseOptionsByQuestionId(scope, qid);
      if (options.length > 0) return options;
    }
  }
  return [];
}

/**
 * 取题目分组标题（`h3.newTestType`）声明的题型。
 *
 * 答题页把分组标题放在 `.aiArea` 的后代里，而已批阅视图（章节测验「已完成」）把它放在
 * `.aiArea` 之前的兄弟节点上，只查后代会永远落空。落空的后果是：题干里没有
 * 「【单选题】」这类标记时（例如填空题）整题被丢弃，因此这里补上兄弟节点查找。
 */
function resolveGroupType(area: Element): QuestionType | null {
  const explicit = detectQuestionType(
    textOf(area.querySelector('.newTestType')),
    area.getAttribute('data-question-type'),
  );
  if (explicit) return explicit;

  // 向上找最近的一个分组标题：标题与它后面的若干 .aiArea 是同级兄弟
  for (
    let sibling = area.previousElementSibling;
    sibling;
    sibling = sibling.previousElementSibling
  ) {
    if (sibling.matches('.newTestType')) return detectQuestionType(textOf(sibling));
  }
  return null;
}

export class TiMuExtractor implements QuestionExtractor {
  readonly id = 'timu';
  readonly confidence = 110;

  supports(root: Document): boolean {
    return root.querySelector('.TiMu.newTiMu, #ZyBottom .TiMu') !== null;
  }

  extract(context: ExtractorContext): readonly Question[] {
    const questions: Question[] = [];
    const areas = Array.from(context.root.querySelectorAll<HTMLElement>('#ZyBottom .aiArea'));

    if (areas.length === 0) {
      context.root.querySelectorAll<HTMLElement>('.TiMu.newTiMu').forEach((container) => {
        const question = this.extractQuestion(container, null, context);
        if (question) questions.push(question);
      });
      return questions;
    }

    let currentType: QuestionType | null = null;
    for (const area of areas) {
      currentType = resolveGroupType(area) ?? currentType;
      area.querySelectorAll<HTMLElement>('.TiMu.newTiMu, .TiMu').forEach((container) => {
        const question = this.extractQuestion(container, currentType, context);
        if (question) questions.push(question);
      });
    }
    return questions;
  }

  private extractQuestion(
    container: HTMLElement,
    inheritedType: QuestionType | null,
    context: ExtractorContext,
  ): Question | null {
    const title = container.querySelector('.Zy_TItle, .question-title') ?? container;
    const type =
      detectQuestionType(
        textOf(title.querySelector('.newZy_TItle')),
        container.getAttribute('typeName'),
        container.getAttribute('data-question-type'),
      ) ??
      inheritedType ??
      inferQuestionType(container);
    if (!type) return null;

    // 题干经分层定位，避免学习通改版换掉题干类名后整题被丢弃
    const stem = resolveQuestionStem(container, title);

    // 「已批阅」视图没有正确答案文本，只有批阅结果图标，拿不到答案时按批阅结果推断
    const userAnswer = extractUserAnswer(container);
    const correctAnswer = extractCorrectAnswer(container);

    return createQuestion({
      number: parseLeadingNumber(textOf(title.querySelector('i.fl')) || stem.rawText),
      type,
      typeMeta: textOf(title.querySelector('.newZy_TItle')) || undefined,
      stem: stem.content,
      options: resolveOptions(container),
      correctAnswer: isRichContentEmpty(correctAnswer)
        ? inferCorrectAnswerFromGrading(container, userAnswer)
        : correctAnswer,
      userAnswer,
      analysis: extractAnalysis(container),
      explicitWrong: hasExplicitWrongMarker(container),
      source: {
        extractor: this.id,
        pageUrl: context.pageUrl,
        selector: '.TiMu.newTiMu',
      },
    });
  }
}
