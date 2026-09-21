/**
 * font-cxsecret 字体反爬解码 —— DOM 接入层
 *
 * 职责：
 *   1. 从文档里找出内联的 @font-face 字体（学习通把字体以 base64 内联在 <style> 中）
 *   2. 用纯逻辑层算出「混淆码位 → 真实汉字」映射并缓存
 *   3. 按作用域把文本还原：只处理带 cxsecret 类名（或其祖先带该类名）的节点
 *
 * 为什么要按作用域过滤：混淆只作用于页面标记为 font-cxsecret 的子树，子树外的
 * 同码位文字属于正常内容，若一并替换会把它改坏。若页面里找不到任何该类名标记
 * （例如对方改名），则退化为整页替换——此时不替换只会输出乱码，替换更有利。
 */

import { buildGlyphDecodeMap, decodeBase64, replaceByCodePointMap } from './cxsecret-font';

export interface CxSecretDecoder {
  /** 可还原的码位数量（映射条目数） */
  readonly size: number;
  /** 文档中是否存在混淆作用域标记；为 false 时对所有文本生效 */
  readonly scoped: boolean;
  decode(text: string): string;
}

/** 与字体家族名无关的作用域判定：类名里含 cxsecret 即视为混淆子树 */
const SCOPE_MARKER_SELECTOR = '[class*="cxsecret"]';
const CXSECRET_FAMILY = /cxsecret/i;
/** @font-face 块内的家族名与字体数据 */
const FONT_FACE_PATTERN = /@font-face\s*\{([^}]*)\}/gi;
const FAMILY_PATTERN = /font-family\s*:\s*(['"]?)([^;'"]+)\1/i;
const SOURCE_URL_PATTERN = /url\(\s*(['"]?)(data:[^'")]+)\1\s*\)/i;
const BASE64_PAYLOAD_PATTERN = /;base64,([A-Za-z0-9+/=\s]+)$/i;

const decoderCache = new WeakMap<Document, CxSecretDecoder | null>();
const scopeCache = new WeakMap<Element, boolean>();

function resolveDocument(doc?: Document | null): Document | null {
  if (doc) return doc;
  return typeof document === 'undefined' ? null : document;
}

/** 取父元素：用 parentNode + nodeType 判定，比 parentElement 在各 DOM 实现上更可靠 */
export function parentElementOf(node: Node | null | undefined): Element | null {
  const parent = node?.parentNode ?? null;
  return parent && parent.nodeType === Node.ELEMENT_NODE ? (parent as Element) : null;
}

function tryQuery(root: Document, selector: string): Element[] {
  try {
    return Array.from(root.querySelectorAll(selector));
  } catch {
    return [];
  }
}

interface FontSource {
  readonly family: string;
  readonly base64: string;
}

function parseFontFaceBlock(block: string): FontSource | null {
  const family = block.match(FAMILY_PATTERN)?.[2]?.trim() ?? '';
  const url = block.match(SOURCE_URL_PATTERN)?.[2] ?? '';
  const base64 = url.match(BASE64_PAYLOAD_PATTERN)?.[1]?.replace(/\s+/g, '') ?? '';
  if (!base64) return null;
  return { family, base64 };
}

/** 收集文档里所有内联 data: 字体的 @font-face */
function collectFontSources(doc: Document): FontSource[] {
  const sources: FontSource[] = [];
  const push = (css: string): void => {
    if (!css || !css.includes('@font-face')) return;
    const pattern = new RegExp(FONT_FACE_PATTERN.source, 'gi');
    for (const match of css.matchAll(pattern)) {
      const source = parseFontFaceBlock(match[1] ?? '');
      if (source) sources.push(source);
    }
  };

  for (const style of tryQuery(doc, 'style')) push(style.textContent ?? '');
  // 同源外链样式表：cssRules 不可读时（跨域）静默跳过
  try {
    for (const sheet of Array.from(doc.styleSheets ?? [])) {
      let rules: CSSRuleList | undefined;
      try {
        rules = sheet.cssRules;
      } catch {
        continue;
      }
      for (const rule of Array.from(rules ?? [])) {
        const cssText = (rule as CSSRule).cssText ?? '';
        if (cssText.includes('@font-face')) push(cssText);
      }
    }
  } catch {
    // styleSheets 在部分文档上不可用，忽略
  }
  return sources;
}

function buildMapFromSources(sources: readonly FontSource[]): Map<number, string> {
  const map = new Map<number, string>();
  for (const source of sources) {
    try {
      const partial = buildGlyphDecodeMap(decodeBase64(source.base64));
      for (const [codePoint, char] of partial) map.set(codePoint, char);
    } catch {
      // 单个字体解析失败不影响其他候选
    }
  }
  return map;
}

function createDecoder(doc: Document): CxSecretDecoder | null {
  const sources = collectFontSources(doc);
  if (sources.length === 0) return null;

  // 先按家族名筛选，命中不了再退回尝试全部内联字体
  const preferred = sources.filter((source) => CXSECRET_FAMILY.test(source.family));
  let map = preferred.length > 0 ? buildMapFromSources(preferred) : new Map<number, string>();
  if (map.size === 0) map = buildMapFromSources(sources);
  if (map.size === 0) return null;

  return {
    size: map.size,
    scoped: tryQuery(doc, SCOPE_MARKER_SELECTOR).length > 0,
    decode: (text: string) => replaceByCodePointMap(text, map),
  };
}

/** 取文档的解码器；无内联反爬字体时返回 null */
export function getCxSecretDecoder(doc?: Document | null): CxSecretDecoder | null {
  const target = resolveDocument(doc);
  if (!target) return null;
  if (decoderCache.has(target)) return decoderCache.get(target) ?? null;
  let decoder: CxSecretDecoder | null = null;
  try {
    decoder = createDecoder(target);
  } catch {
    decoder = null;
  }
  decoderCache.set(target, decoder);
  return decoder;
}

/** 清空缓存：同一文档内重载了内容（字体可能变化）后调用 */
export function clearCxSecretDecoderCache(doc?: Document | null): void {
  const target = resolveDocument(doc);
  if (target) decoderCache.delete(target);
}

/**
 * 判定元素是否处于混淆作用域内（自身或任一祖先带 cxsecret 类名）。
 * 结果按元素缓存，避免逐文本节点重复向上遍历。
 */
export function isCxSecretScope(element: Element | null): boolean {
  if (!element) return false;
  const cached = scopeCache.get(element);
  if (cached !== undefined) return cached;
  let result = false;
  let node: Element | null = element;
  while (node) {
    const className = node.getAttribute('class') ?? '';
    if (className && CXSECRET_FAMILY.test(className)) {
      result = true;
      break;
    }
    node = parentElementOf(node);
  }
  scopeCache.set(element, result);
  return result;
}

/**
 * 还原元素文本里的混淆汉字。
 * @param text  已读出的原始文本
 * @param scope 文本所属元素（用于判定作用域）
 */
export function decodeCxSecretText(text: string, scope: Element | null, doc?: Document | null): string {
  if (!text) return text;
  const decoder = getCxSecretDecoder(doc ?? scope?.ownerDocument ?? null);
  if (!decoder) return text;
  if (decoder.scoped && !isCxSecretScope(scope)) return text;
  return decoder.decode(text);
}
