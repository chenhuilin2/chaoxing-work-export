import { normalizeInlineWhitespace } from '../utils/text';

const TITLE_SELECTORS = [
  '.mark_title',
  '.newTestTitle',
  '.TestTitle_name',
  '.testTitle',
  '.courseName',
  '.chapterText',
  '.chapter-name',
  'h1',
] as const;

// 学生学习页面：章节/测验标题所在的容器，优先作为提取标题来源。
// 真实 DOM 用类名（<div class="prev_title" title="毛泽东思想的主要内容">），
// 这里同时兼容 id 形式，两者取先命中者。
const PREV_TITLE_SELECTORS = ['#prev_title', '.prev_title'] as const;

/**
 * 学习通学生学习页面（章节页外层壳）上的章节标题。
 *
 * 该节点位于外层壳文档，而题目在嵌套的知识卡片 / 答题 iframe 里，嵌套文档只能解析出
 * 自己的标题（如「章节测验 待完成」），因此需要由外层壳提供章节名。无此节点时返回空串。
 */
export function resolvePrevTitle(root: Document): string {
  for (const selector of PREV_TITLE_SELECTORS) {
    const element = root.querySelector<HTMLElement>(selector);
    if (!element) continue;
    const fromAttribute = normalizeInlineWhitespace(element.getAttribute('title') ?? '');
    if (fromAttribute) return fromAttribute;
    const fromText = normalizeInlineWhitespace(element.textContent ?? '');
    if (fromText) return fromText;
  }
  return '';
}

export function resolvePageTitle(root: Document): string {
  // 学生学习页面：优先使用 #prev_title / .prev_title 的 title 属性，回退到其正文
  const prevTitle = resolvePrevTitle(root);
  if (prevTitle) return prevTitle;
  for (const selector of TITLE_SELECTORS) {
    const value = normalizeInlineWhitespace(root.querySelector(selector)?.textContent ?? '');
    if (value) return value;
  }
  return normalizeInlineWhitespace(root.title)
    .replace(/[-_|]\s*(?:超星学习通|学习通).*$/u, '')
    .trim() || '学习通题目';
}

export function resolvePageUrl(root: Document): string {
  try {
    return root.location.href;
  } catch {
    return window.location.href;
  }
}
