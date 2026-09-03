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

const CHAPTER_ITEM_SELECTORS = [
  '.posCatalog_select',
  '.posCatalog_name',
  '.catalog_name',
  '.chapter-item',
  '[data-chapter-id]',
  '[data-id][class*="catalog"]',
] as const;

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

    const target =
      element.matches('a, button, [role="button"]')
        ? element
        : element.querySelector<HTMLElement>('a, button, [role="button"]') ?? element;
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
        const filtered = raw.filter((element) => {
          if (!element.offsetParent && element.getClientRects().length === 0) return false;
          const title = this.titleOf(element, 0);
          return title.length > 0 && !/^(?:目录|章节|返回)$/u.test(title);
        });
        if (filtered.length > 1) return this.removeNestedDuplicates(filtered);
      }
    }
    return [];
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
      element.innerText;
    const title = normalizeInlineWhitespace(explicit ?? '');
    return title || `第 ${index + 1} 章`;
  }

  private isActive(element: HTMLElement): boolean {
    if (element.getAttribute('aria-current') === 'true') return true;
    if (element.getAttribute('aria-selected') === 'true') return true;
    return ACTIVE_CLASSES.some((className) => element.classList.contains(className)) ||
      Boolean(element.closest(ACTIVE_CLASSES.map((className) => `.${className}`).join(',')));
  }
}
