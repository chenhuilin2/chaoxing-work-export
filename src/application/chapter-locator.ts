// 学习通目录树的真实结构是两级（见 src/domain/chapter.ts 的说明）：
//   章：`.posCatalog_select.firstLayer > .posCatalog_title`（分组标题，点不开、不能提取）
//   小节：`.posCatalog_level` 下的 `.posCatalog_name`（唯一挂 onclick、真正能提取的条目）
// 两级都把编号印在 `<em class="posCatalog_sbar">` 里（章是「1」、小节是「1.1」），
// 因此这里读页面的真编号，而不是按顺序现编。
import type { ChapterGroupInfo } from '../domain/chapter';
import { normalizeInlineWhitespace } from '../utils/text';

export interface ChapterDescriptor {
  readonly index: number;
  readonly id: string;
  /** 小节名（不含编号） */
  readonly title: string;
  /** 页面上的小节编号（「1.1」）；页面没给编号时为空串 */
  readonly label: string;
  readonly active: boolean;
  /** 所属章；扁平目录（没有第一级分组）时为 null */
  readonly group: ChapterGroupInfo | null;
}

interface CatalogEntry {
  readonly element: HTMLElement;
  readonly group: ChapterGroupInfo | null;
}

const CHAPTER_CONTAINER_SELECTORS = [
  '#coursetree',
  '.catalog_points',
  '.chapter-list',
  '.catalog-list',
  '.posCatalog_list',
] as const;

// 章节条目：优先取真正承载章节名的叶子节点（学习通真实结构是 span.posCatalog_name，
// 它是唯一挂 onclick 的节点），容器 div 只作为兜底候选
const CHAPTER_ITEM_SELECTORS = [
  '.posCatalog_name',
  '.catalog_name',
  '.chapter-item',
  '[data-chapter-id]',
  '[data-id][class*="catalog"]',
  '.posCatalog_select',
] as const;

// 分组标题：第一级目录项（真实结构为 .posCatalog_select.firstLayer > span.posCatalog_title），
// 它们不是可提取章节，但要作为「章」标题参与两级导出
const CHAPTER_GROUP_SELECTORS = ['.firstLayer', '.posCatalog_title'] as const;

// 编号节点：章与小节都用 .posCatalog_sbar，老结构里小节用 .posCatalog_num
const CHAPTER_NUMBER_SELECTORS = ['.posCatalog_sbar', '.posCatalog_num', '.chapter-num'] as const;

const ACTIVE_CLASSES = ['active', 'cur', 'current', 'on', 'selected', 'posCatalog_active'] as const;

/** 「1.1」「2.3」这类编号前缀；必须是纯数字点号（避免把「3D 打印」当成编号） */
const SECTION_NUMBER_PATTERN = /^([0-9]+(?:\.[0-9]+)*)(?:\s|$)/;

/** Finds and activates course chapter entries without embedding extraction logic. */
export class ChapterLocator {
  constructor(private readonly root: Document = document) {}

  list(): ChapterDescriptor[] {
    return this.entries().map((entry, index) => {
      const identity = this.identityOf(entry.element);
      return {
        index,
        id: this.idOf(entry.element, index),
        // 名字读不出时退回「第 N 章」，与改动前一致（面板上不能出现空标题）
        title: identity.title || `第 ${index + 1} 章`,
        label: identity.label,
        active: this.isActive(entry.element),
        group: entry.group,
      };
    });
  }

  activeIndex(): number | null {
    const index = this.elements().findIndex((element) => this.isActive(element));
    return index >= 0 ? index : null;
  }

  activate(index: number): boolean {
    const element = this.elements()[index];
    if (!element) return false;

    const target = this.clickableTarget(element);
    target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
    target.click();
    return true;
  }

  private elements(): HTMLElement[] {
    return this.entries().map((entry) => entry.element);
  }

  /**
   * 解析目录选择器：沿用「按优先级挑第一个命中数 > 1 的选择器」的既有判定，
   * 保证「哪些节点算章节」与改动前完全一致。
   */
  private resolveSearch(): { root: ParentNode; selector: string } | null {
    const searchRoots: ParentNode[] = [this.root];
    for (const selector of CHAPTER_CONTAINER_SELECTORS) {
      const container = this.root.querySelector(selector);
      if (container) searchRoots.unshift(container);
    }

    for (const searchRoot of searchRoots) {
      for (const selector of CHAPTER_ITEM_SELECTORS) {
        const raw = [...searchRoot.querySelectorAll<HTMLElement>(selector)];
        const filtered = raw.filter((element) => this.isChapterItem(element));
        if (filtered.length > 1) return { root: searchRoot, selector };
      }
    }
    return null;
  }

  /**
   * 按文档顺序走一遍目录：每个小节挂到它前面最近的那个「章」上。
   *
   * 只依赖文档顺序、不依赖具体层级，所以「章 → .posCatalog_level > 小节」的真实结构
   * 与「分组标题和小节平铺成兄弟节点」的老结构都能正确归属。
   */
  private entries(): CatalogEntry[] {
    const search = this.resolveSearch();
    if (!search) return [];

    const entries: CatalogEntry[] = [];
    const scope: { current: ChapterGroupInfo | null } = { current: null };
    // 命中即跳过整棵子树：分组标题内部不再识别下级分组，小节命中后也不再下钻
    //（等价于原来的 removeNestedDuplicates 加「最内层选择器优先」）
    const consumed = new Set<Element>();

    for (const element of search.root.querySelectorAll<HTMLElement>('*')) {
      if (this.isInsideConsumed(element, consumed)) continue;

      if (this.isGroupElement(element)) {
        scope.current = this.groupInfoOf(element);
        consumed.add(element);
        continue;
      }
      if (element.matches(search.selector) && this.isChapterItem(element)) {
        entries.push({ element, group: scope.current });
        consumed.add(element);
      }
    }
    return entries;
  }

  private isInsideConsumed(element: HTMLElement, consumed: ReadonlySet<Element>): boolean {
    if (consumed.size === 0) return false;
    for (let node = element.parentElement; node; node = node.parentElement) {
      if (consumed.has(node)) return true;
    }
    return false;
  }

  private isGroupElement(element: HTMLElement): boolean {
    return CHAPTER_GROUP_SELECTORS.some((selector) => element.matches(selector));
  }

  /** 是否为可提取章节：排除分组标题与不可见占位项 */
  private isChapterItem(element: HTMLElement): boolean {
    // 分组标题（.firstLayer / .posCatalog_title）不是章节
    const isGroup = CHAPTER_GROUP_SELECTORS.some(
      (selector) => element.matches(selector) || Boolean(element.closest(selector)),
    );
    if (isGroup) return false;
    if (!this.isVisible(element)) return false;

    const title = this.identityOf(element).title;
    return title.length > 0 && !/^(?:目录|章节|返回)$/u.test(title);
  }

  /** 可见性判定：测试环境（linkedom）没有布局信息，此时不做隐藏判定 */
  private isVisible(element: HTMLElement): boolean {
    if (typeof element.getClientRects !== 'function') return true;
    if (element.getClientRects().length > 0) return true;
    return Boolean(element.offsetParent);
  }

  /**
   * 选中真正可点击的节点。
   * 学习通把 onclick 挂在章节名节点（span.posCatalog_name）上，容器 div 自身没有事件，
   * 对容器派发点击不会向下传递到子节点，因此必须精确定位到带 onclick 的元素。
   */
  private clickableTarget(element: HTMLElement): HTMLElement {
    if (element.hasAttribute('onclick')) return element;
    const handler = element.querySelector<HTMLElement>('[onclick]');
    if (handler) return handler;
    if (element.matches('a, button, [role="button"]')) return element;
    return element.querySelector<HTMLElement>('a, button, [role="button"]') ?? element;
  }

  /** 章信息：编号读页面上的 em，名字优先 title 属性 */
  private groupInfoOf(element: HTMLElement): ChapterGroupInfo {
    const identity = this.identityOf(element);
    return { id: this.idOf(element, 0), label: identity.label, title: identity.title };
  }

  /**
   * 一次读出「编号 + 名字」。
   *
   * 名字优先取 title 属性（学习通把它写成不含编号的纯名字，也不带「已完成」这类提示文本），
   * 没有 title 时才退回文本，并把文本里的编号（`em.posCatalog_sbar`）从名字里去掉 ——
   * 编号单独存在 label 里、拼标题时再组合，避免出现「1.1 1.1 导论1」。
   */
  private identityOf(element: HTMLElement): { label: string; title: string } {
    const fromNode = normalizeInlineWhitespace(this.numberNodeOf(element)?.textContent ?? '');

    const explicit = normalizeInlineWhitespace(
      element.getAttribute('title') ??
        element.querySelector<HTMLElement>('[title]')?.getAttribute('title') ??
        '',
    );
    if (explicit) return { label: fromNode || this.numberPrefixOf(explicit), title: explicit };

    const named = element.querySelector<HTMLElement>(
      '.catalog_name, .chapter-title, .posCatalog_name',
    );
    const raw = normalizeInlineWhitespace((named ?? element).textContent ?? '');
    if (fromNode && raw.startsWith(fromNode)) {
      return { label: fromNode, title: raw.slice(fromNode.length).trim() };
    }
    const prefix = this.numberPrefixOf(raw);
    return prefix
      ? { label: prefix, title: raw.slice(prefix.length).trim() }
      : { label: '', title: raw };
  }

  private numberPrefixOf(text: string): string {
    return SECTION_NUMBER_PATTERN.exec(text)?.[1] ?? '';
  }

  private numberNodeOf(element: HTMLElement): HTMLElement | null {
    for (const selector of CHAPTER_NUMBER_SELECTORS) {
      if (element.matches(selector)) return element;
      const found = element.querySelector<HTMLElement>(selector);
      if (found) return found;
    }
    return null;
  }

  private idOf(element: HTMLElement, index: number): string {
    return (
      element.dataset.chapterId ??
      element.dataset.id ??
      element.getAttribute('id') ??
      `chapter-${index + 1}`
    );
  }

  private isActive(element: HTMLElement): boolean {
    if (element.getAttribute('aria-current') === 'true') return true;
    if (element.getAttribute('aria-selected') === 'true') return true;

    // 激活态类名挂在容器 div 上（class="posCatalog_select posCatalog_active"），
    // 而章节条目现在是它内部的名字节点，因此必须连同所属容器一起判断
    const owner = element.closest<HTMLElement>('.posCatalog_select');
    const nodes = owner && owner !== element ? [element, owner] : [element];
    if (
      nodes.some((node) => ACTIVE_CLASSES.some((className) => node.classList.contains(className)))
    ) {
      return true;
    }
    return Boolean(element.closest(ACTIVE_CLASSES.map((className) => `.${className}`).join(',')));
  }
}
