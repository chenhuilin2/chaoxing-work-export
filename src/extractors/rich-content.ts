import type { ImagePart, RichContent, RichPart, TextPart } from '../domain/question';
import { normalizeWhitespace } from '../utils/text';

const BLOCK_TAGS = new Set([
  'P',
  'DIV',
  'LI',
  'DD',
  'DT',
  'TR',
  'TABLE',
  'SECTION',
  'ARTICLE',
  'BLOCKQUOTE',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
]);

function sameTextStyle(left: TextPart, right: TextPart): boolean {
  return (
    Boolean(left.bold) === Boolean(right.bold) &&
    Boolean(left.italic) === Boolean(right.italic) &&
    Boolean(left.subScript) === Boolean(right.subScript) &&
    Boolean(left.superScript) === Boolean(right.superScript)
  );
}

export function normalizeRichContent(parts: readonly RichPart[]): RichPart[] {
  const output: RichPart[] = [];
  for (const part of parts) {
    if (part.type === 'text') {
      const text = part.text.replace(/\u00a0/g, ' ').replace(/[\t\r\f ]+/g, ' ');
      if (!text) continue;
      const previous = output[output.length - 1];
      if (previous?.type === 'text' && sameTextStyle(previous, part)) {
        output[output.length - 1] = { ...previous, text: previous.text + text };
      } else {
        output.push({ ...part, text });
      }
      continue;
    }
    if (part.type === 'image') {
      if (part.url) output.push(part);
      continue;
    }
    if (output[output.length - 1]?.type !== 'break') output.push(part);
  }

  // 修剪首尾的换行与纯空白文本：尾部空白会挡住相邻换行，导致 Word 导出时题目与选项之间出现多余空行
  const isTrimmable = (part: RichPart | undefined): boolean => {
    if (!part) return false;
    return part.type === 'break' || (part.type === 'text' && !part.text.trim());
  };
  while (isTrimmable(output[0])) output.shift();
  while (isTrimmable(output[output.length - 1])) output.pop();
  return output;
}

function resolveUrl(value: string, baseUrl: string): string {
  if (!value || value === 'about:blank') return '';
  try {
    return new URL(value, baseUrl).href;
  } catch {
    return value;
  }
}

export function resolveImageUrl(image: HTMLImageElement): string {
  const attributes = [
    'data-original',
    'data-src',
    'data-lazy-src',
    'origin-src',
    'src',
    'fileid',
  ];
  const baseUrl = image.ownerDocument.baseURI;
  if (image.currentSrc) return resolveUrl(image.currentSrc, baseUrl);
  for (const attribute of attributes) {
    const value = image.getAttribute(attribute)?.trim();
    if (value) return resolveUrl(value, baseUrl);
  }
  return resolveUrl(image.src, baseUrl);
}

export function extractBackgroundImages(element: Element): ImagePart[] {
  if (!(element instanceof HTMLElement)) return [];
  let background = element.style.backgroundImage;
  try {
    if (!background || background === 'none') {
      background = element.ownerDocument.defaultView?.getComputedStyle(element).backgroundImage ?? '';
    }
  } catch {
    // Computed styles may be unavailable for detached or cross-origin documents.
  }
  if (!background || background === 'none') return [];

  const images: ImagePart[] = [];
  const pattern = /url\(\s*(["']?)(.*?)\1\s*\)/giu;
  for (const match of background.matchAll(pattern)) {
    const value = match[2]?.trim();
    if (!value) continue;
    const url = resolveUrl(value, element.ownerDocument.baseURI);
    if (url) images.push({ type: 'image', url, alt: element.getAttribute('aria-label') ?? '' });
  }
  return images;
}

interface TextStyle {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly subScript: boolean;
  readonly superScript: boolean;
}

const EMPTY_STYLE: TextStyle = {
  bold: false,
  italic: false,
  subScript: false,
  superScript: false,
};

export function extractRichContent(element: Element | null): RichPart[] {
  if (!element) return [];
  const parts: RichPart[] = [];

  const walk = (node: Node, style: TextStyle): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      parts.push({ type: 'text', text: node.nodeValue ?? '', ...style });
      return;
    }
    if (!(node instanceof Element)) return;

    const tag = node.tagName.toUpperCase();
    if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE'].includes(tag)) return;
    if (tag === 'IMG') {
      const url = resolveImageUrl(node as HTMLImageElement);
      if (url) parts.push({ type: 'image', url, alt: node.getAttribute('alt') ?? '' });
      return;
    }
    if (tag === 'BR') {
      parts.push({ type: 'break' });
      return;
    }

    const childStyle: TextStyle = {
      bold: style.bold || tag === 'STRONG' || tag === 'B',
      italic: style.italic || tag === 'EM' || tag === 'I',
      subScript: style.subScript || tag === 'SUB',
      superScript: style.superScript || tag === 'SUP',
    };
    const lengthBefore = parts.length;
    node.childNodes.forEach((child) => walk(child, childStyle));
    if (BLOCK_TAGS.has(tag) && parts.length > lengthBefore) parts.push({ type: 'break' });
  };

  element.childNodes.forEach((child) => walk(child, EMPTY_STYLE));
  const existingImageUrls = new Set(
    parts.filter((part): part is ImagePart => part.type === 'image').map((part) => part.url),
  );
  for (const image of extractBackgroundImages(element)) {
    if (!existingImageUrls.has(image.url)) parts.push(image);
  }
  let normalized = normalizeRichContent(parts);

  if (normalized.length === 0) {
    const fallback = [
      element.getAttribute('aria-label'),
      element.getAttribute('title'),
      element.getAttribute('data-content'),
      element.getAttribute('data-text'),
      element.getAttribute('data-value'),
    ].find((value) => Boolean(value?.trim()));
    if (fallback) normalized = [{ type: 'text', text: fallback }];
  }
  return normalized;
}

export function richContentToText(
  content: RichContent,
  imageFormatter: (url: string, alt: string) => string = () => '',
): string {
  let output = '';
  for (const part of content) {
    if (part.type === 'text') output += part.text;
    else if (part.type === 'image') output += imageFormatter(part.url, part.alt);
    else output += '\n';
  }
  return normalizeWhitespace(output);
}

export function isRichContentEmpty(content: RichContent): boolean {
  return !content.some((part) => {
    return part.type === 'image' || (part.type === 'text' && Boolean(part.text.trim()));
  });
}

function cloneContent(content: RichContent): RichPart[] {
  return content.map((part) => ({ ...part }));
}

export function stripOptionPrefix(content: RichContent): {
  readonly key: string;
  readonly content: RichPart[];
} {
  const output = cloneContent(content);
  let key = '';
  for (let index = 0; index < output.length; index += 1) {
    const part = output[index];
    if (part?.type !== 'text' || !part.text.trim()) continue;
    const match = part.text.match(/^\s*([A-Z])\s*[.、．:：]?\s*/iu);
    if (match?.[1]) {
      key = match[1].toUpperCase();
      output[index] = { ...part, text: part.text.slice(match[0].length) };
    }
    break;
  }
  return { key, content: normalizeRichContent(output) };
}

export function stripQuestionPrefix(content: RichContent): RichPart[] {
  const output = cloneContent(content);
  let numberHandled = false;
  let typeHandled = false;
  for (let index = 0; index < output.length; index += 1) {
    const part = output[index];
    if (part?.type !== 'text') continue;
    let text = part.text.replace(/\s*[（(]\s*\d+(?:\.\d+)?\s*分?\s*[）)]\s*$/u, '');
    if (!numberHandled) {
      text = text.replace(/^\s*\d+\s*[.、．]\s*/u, '');
      if (text.trim()) numberHandled = true;
    }
    if (!typeHandled) {
      text = text.replace(
        /^\s*(?:[（(【[]\s*)?(?:单选题|单项选择题|多选题|多项选择题|填空题|判断题|简答题|论述题|问答题|选择题)(?:\s*[）)】\]])?\s*/u,
        '',
      );
      if (text.trim()) typeHandled = true;
    }
    output[index] = { ...part, text };
  }
  return normalizeRichContent(output);
}

export function stripAnswerLabel(content: RichContent): RichPart[] {
  const output = cloneContent(content);
  for (let index = 0; index < output.length; index += 1) {
    const part = output[index];
    if (part?.type !== 'text' || !part.text.trim()) continue;
    output[index] = {
      ...part,
      text: part.text.replace(
        /^\s*(?:正确答案|参考答案|答案|我的答案|你的答案|学生答案|答案解析|解析)\s*[:：]?\s*/u,
        '',
      ),
    };
    break;
  }
  return normalizeRichContent(output);
}

export function joinRichContents(
  contents: readonly RichContent[],
  separator = '；',
): RichPart[] {
  const output: RichPart[] = [];
  for (const content of contents) {
    if (isRichContentEmpty(content)) continue;
    if (output.length > 0) output.push({ type: 'text', text: separator });
    output.push(...content.map((part) => ({ ...part })));
  }
  return normalizeRichContent(output);
}

export function textContent(value: string): RichPart[] {
  const normalized = value.trim();
  return normalized ? [{ type: 'text', text: normalized }] : [];
}
