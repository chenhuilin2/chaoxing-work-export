import { normalizeInlineWhitespace } from '../utils/text';

export interface ChapterDescriptor {
  readonly index: number;
  readonly id: string;
  readonly title: string;
  readonly active: boolean;
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
// 它们不是可提取章节，必须排除
const CHAPTER_GROUP_SELECTORS = ['.firstLayer', '.posCatalog_title'] as const;

const ACTIVE_CLASSES = ['active', 'cur', 'current', 'on', 'selected', 'posCatalog_active'] as const;

/** Finds and activates course chapter entries without embedding extraction logic. */
export class ChapterLocator {
  constructor(private readonly root: Document = document) {}

  list(): ChapterDescriptor[] {
    return this.elements().map((element, index) => ({
      index,
      id: this.idOf(element, index),
      title: this.titleOf(element, index),
      active: this.isActive(element),
    }));
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
    const searchRoots: ParentNode[] = [this.root];
    for (const selector of CHAPTER_CONTAINER_SELECTORS) {
      const container = this.root.querySelector(selector);
      if (container) searchRoots.unshift(container);
    }

    for (const searchRoot of searchRoots) {
      for (const selector of CHAPTER_ITEM_SELECTORS) {
        const raw = [...searchRoot.querySelectorAll<HTMLElement>(selector)];
        const filtered = raw.filter((element) => this.isChapterItem(element));
        if (filtered.length > 1) return this.removeNestedDuplicates(filtered);
      }
    }
    return [];
  }

  /** 是否为可提取章节：排除分组标题与不可见占位项 */
  private isChapterItem(element: HTMLElement): boolean {
    // 分组标题（.firstLayer / .posCatalog_title）不是章节
    const isGroup = CHAPTER_GROUP_SELECTORS.some(
      (selector) => element.matches(selector) || Boolean(element.closest(selector)),
    );
    if (isGroup) return false;
    if (!this.isVisible(element)) return false;

    const title = this.titleOf(element, 0);
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

  private removeNestedDuplicates(elements: readonly HTMLElement[]): HTMLElement[] {
    return elements.filter((element) => {
      return !elements.some((other) => other !== element && other.contains(element));
    });
  }

  private idOf(element: HTMLElement, index: number): string {
    return (
      element.dataset.chapterId ??
      element.dataset.id ??
      element.getAttribute('id') ??
      `chapter-${index + 1}`
    );
  }

  private titleOf(element: HTMLElement, index: number): string {
    const explicit =
      element.getAttribute('title') ??
      element.querySelector<HTMLElement>('[title]')?.getAttribute('title') ??
      element.querySelector<HTMLElement>('.catalog_name, .chapter-title, .posCatalog_name')
        ?.innerText ??
      element.innerText ??
      // 兜底：章节名节点不带 title 时也要取到真实名称，避免退化成「第 N 章」
      element.textContent;
    const title = normalizeInlineWhitespace(explicit ?? '');
    return title || `第 ${index + 1} 章`;
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
