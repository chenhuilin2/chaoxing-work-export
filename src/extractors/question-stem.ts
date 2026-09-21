import type { RichContent } from '../domain/question';
import { firstMatch, textOf } from '../utils/dom';
import { extractRichContent, isRichContentEmpty, stripQuestionPrefix } from './rich-content';

/**
 * 题干容器候选选择器：按「当前版本 → 历史版本」排列，命中即用。
 *
 * 学习通改版导致题干丢失时，优先在这里追加新类名，无需改动任何提取器的流程代码。
 * 列表内容为四个提取器原有选择器的并集，保证既有页面行为不变。
 */
export const STEM_CONTAINER_SELECTORS = [
  '[data-role="stem"]',
  '.question-stem',
  '.questionStem',
  '.question-content',
  '.qtContent',
  '.mark_name',
  '.fontLabel', // 2026 版答题页：.Zy_TItle > div.fontLabel
  '.subject-title',
  '.subject',
] as const;

/**
 * 题型标记选择器：题干与「【单选题】」同处一个元素，是超星最稳定的结构特征。
 * 用于「已知类名全部失效」时反推题干容器。
 */
export const STEM_TYPE_MARKER_SELECTORS = [
  '.newZy_TItle',
  '[class*="TestType"]',
  '[class*="ztType"]',
] as const;

/**
 * 非题干节点：选项列表、作答控件、答案区、题型标记。
 * 仅在「剔除式兜底」路径使用（容器本身混有选项时才需要裁剪）。
 */
const STEM_EXCLUDE_SELECTOR = [
  'script',
  'style',
  'ul',
  'ol',
  'select',
  'textarea',
  '.num_option',
  '.mark_answer',
  '.ans-job-icon',
  '.newZy_TItle',
  '[class*="TestType"]',
  '[class*="ztType"]',
].join(',');

/** 题号候选元素：形如 <i class="fl">1</i> */
const INDEX_CANDIDATE_SELECTOR = 'i,em,.fl,[class*="index"],[class*="number"]';

/** 纯题号文本：不含「.」「、」等分隔符，避免误伤题干中自带序号的句子 */
const BARE_INDEX = /^\s*\d{1,3}\s*$/u;

export interface QuestionStem {
  /** 已剥离题号与题型前缀的题干内容，直接用于 createQuestion */
  readonly content: RichContent;
  /** 剥离前的原始文本，供提取题号使用（与旧实现的 textOf(stemElement) 语义一致） */
  readonly rawText: string;
}

const EMPTY_STEM: QuestionStem = { content: [], rawText: '' };

/** 元素是否含有可作题干的内容（文本或图片） */
function hasStemContent(element: Element): boolean {
  if (element.querySelector('img') !== null) return true;
  return Boolean((element.textContent ?? '').replace(/\u00a0/g, ' ').trim());
}

/** 复制节点并剔除题号、题型标记、选项列表等非题干节点 */
function pruneStemElement(element: Element): Element {
  const clone = element.cloneNode(true) as Element;
  clone.querySelectorAll(STEM_EXCLUDE_SELECTOR).forEach((node) => node.remove());
  clone.querySelectorAll(INDEX_CANDIDATE_SELECTOR).forEach((node) => {
    if (BARE_INDEX.test(node.textContent ?? '')) node.remove();
  });
  return clone;
}

/** 读取容器内容并剥离题号/题型前缀；prune 为真时先裁剪非题干节点 */
function readStem(element: Element, prune: boolean): QuestionStem {
  const source = prune ? pruneStemElement(element) : element;
  return {
    content: stripQuestionPrefix(extractRichContent(source)),
    rawText: textOf(source),
  };
}

/**
 * 分层定位题干，任一层取到非空内容即返回：
 *
 * 1. 已知题干类名（当前版本 → 历史版本）
 * 2. 题型标记的父元素（题干容器改名也不影响这条结构特征）
 * 3. 层层兜底：剔除题号与选项后，取标题容器 / 整题容器的剩余内容
 *
 * 题干为空会让 createQuestion 判该题无效并整题丢弃，因此这里必须尽力给出结果。
 */
export function resolveQuestionStem(container: Element, title?: Element | null): QuestionStem {
  const scope = title ?? container;

  // 第一层：已知题干容器。不裁剪，以保留懒加载图片已解析出的 currentSrc
  const known =
    firstMatch<Element>(scope, STEM_CONTAINER_SELECTORS) ??
    firstMatch<Element>(container, STEM_CONTAINER_SELECTORS);
  if (known && hasStemContent(known)) {
    const stem = readStem(known, false);
    if (!isRichContentEmpty(stem.content)) return stem;
  }

  // 第二层：题型标记的父元素 —— 题干与「【单选题】」同级
  const marker =
    firstMatch<Element>(scope, STEM_TYPE_MARKER_SELECTORS) ??
    firstMatch<Element>(container, STEM_TYPE_MARKER_SELECTORS);
  const markerParent = marker?.parentElement ?? null;
  if (markerParent && markerParent !== scope) {
    const stem = readStem(markerParent, true);
    if (!isRichContentEmpty(stem.content)) return stem;
  }

  // 第三层：标题容器本身，剔除题号、选项列表与答案区后的剩余内容
  const holders = scope === container ? [scope] : [scope, container];
  for (const holder of holders) {
    if (!hasStemContent(holder)) continue;
    const stem = readStem(holder, true);
    if (!isRichContentEmpty(stem.content)) return stem;
  }
  return EMPTY_STEM;
}