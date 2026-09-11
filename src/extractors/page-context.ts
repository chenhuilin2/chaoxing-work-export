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

export function resolvePageTitle(root: Document): string {
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
