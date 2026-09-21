import { collectAccessibleDocuments } from '../utils/dom';
import { normalizeInlineWhitespace } from '../utils/text';

export interface TaskTabDescriptor {
  readonly index: number;
  readonly id: string;
  readonly title: string;
  readonly active: boolean;
}

/** 任务卡切换结果：已激活 / 本次完成切换 / 页面无可用任务卡 */
export type TaskTabState = 'active' | 'switched' | 'missing';

// 任务卡切换栏容器：知识卡片页顶部（视频 / 章节测验 / 作业 …）所在位置
const TAB_BAR_SELECTORS = ['#prev_tab', '.prev_list', '.prev_ul'] as const;

// 任务卡条目：真实结构是 li[role="option"]，文案在 span.spanText 中
const TAB_ITEM_SELECTORS = ['li[role="option"]', '.prev_ul li', 'li'] as const;

// 承载题目的任务卡关键词（章节测验 / 作业 / 考试 / 练习…）
const QUESTION_TAB_PATTERN = /测验|作业|考试|练习|测试|自测/u;

// 激活态类名：学习通实际使用 .active，其余作为兼容候选
const ACTIVE_CLASSES = ['active', 'on', 'cur', 'current', 'selected'] as const;

// 补切任务卡的默认节流参数：单章最多补切几次、两次之间至少间隔多久。
// 冷却时间只需覆盖「新页面重建任务卡栏」的短暂窗口，太长会让章节切换变慢。
const DEFAULT_MAX_CLICKS = 4;
const DEFAULT_CLICK_COOLDOWN_MS = 400;

/**
 * 定位知识卡片页的「任务点切换栏」。
 *
 * 任务卡内容是懒加载的：卡片的真实地址写在内部 iframe 的 `_src` 上，只有点击该卡触发
 * `changeDisplayContent()` 之后才会写回 `src` 并加载内容。因此任务卡不切换时，
 * 页面里根本不存在题目 DOM —— 这正是「必须手动切到章节测验才能提取」的根因。
 */
export class TaskTabLocator {
  constructor(private readonly root: Document = document) {}

  /** 承载任务卡栏的文档：知识卡片页通常位于同源 iframe 内 */
  holder(): Document | null {
    for (const { document: candidate } of collectAccessibleDocuments(this.root)) {
      if (TAB_BAR_SELECTORS.some((selector) => candidate.querySelector(selector))) return candidate;
    }
    return null;
  }

  /** 任务卡栏是否已经渲染出来 */
  present(): boolean {
    return this.holder() !== null;
  }

  list(): TaskTabDescriptor[] {
    return this.elements().map((element, index) => ({
      index,
      id: element.getAttribute('id') ?? `tab-${index + 1}`,
      title: this.titleOf(element),
      active: this.isActive(element),
    }));
  }

  activeIndex(): number | null {
    const index = this.elements().findIndex((element) => this.isActive(element));
    return index >= 0 ? index : null;
  }

  /** 题目类任务卡在列表中的位置（章节测验 / 作业 / 考试…） */
  questionTabIndex(): number | null {
    const index = this.elements().findIndex((element) =>
      QUESTION_TAB_PATTERN.test(this.titleOf(element)),
    );
    return index >= 0 ? index : null;
  }

  /**
   * 确保题目类任务卡处于激活状态。
   * 已激活返回 active；本次完成切换返回 switched；没有任务卡栏或没有题目卡返回 missing。
   */
  ensureQuestionTab(): TaskTabState {
    const element = this.elements().find((candidate) => {
      return QUESTION_TAB_PATTERN.test(this.titleOf(candidate));
    });
    if (!element) return 'missing';
    if (this.isActive(element)) return 'active';
    return this.click(element) ? 'switched' : 'missing';
  }

  activate(index: number): boolean {
    const element = this.elements()[index];
    return element ? this.click(element) : false;
  }

  /** 触发任务卡自身的 onclick（学习通挂在 li 上，对应 changeDisplayContent()） */
  private click(element: HTMLElement): boolean {
    element.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
    element.click();
    return true;
  }

  private elements(): HTMLElement[] {
    const holder = this.holder();
    if (!holder) return [];

    for (const barSelector of TAB_BAR_SELECTORS) {
      const bar = holder.querySelector(barSelector);
      if (!bar) continue;

      for (const itemSelector of TAB_ITEM_SELECTORS) {
        const items = [...bar.querySelectorAll<HTMLElement>(itemSelector)].filter((element) => {
          return this.titleOf(element).length > 0;
        });
        if (items.length > 0) return items;
      }
    }
    return [];
  }

  private titleOf(element: HTMLElement): string {
    const candidates = [
      element.getAttribute('title'),
      element.querySelector<HTMLElement>('.spanText')?.textContent,
      // 兜底：直接读文本，并去掉左侧序号（结构为 <span class="num">2</span><span class="spanText">章节测验</span>）
      (element.textContent ?? '').replace(/^\d+\s*/u, ''),
    ];
    for (const candidate of candidates) {
      const title = normalizeInlineWhitespace(candidate ?? '');
      if (title) return title;
    }
    return '';
  }

  private isActive(element: HTMLElement): boolean {
    if (element.getAttribute('aria-current') === 'true') return true;
    if (element.getAttribute('aria-selected') === 'true') return true;
    if (ACTIVE_CLASSES.some((className) => element.classList.contains(className))) return true;
    return Boolean(element.closest(ACTIVE_CLASSES.map((className) => `.${className}`).join(',')));
  }
}

/**
 * 任务卡补切节流器，配合「采样等待」使用。
 *
 * 章节切换会重建卡片页，新页面默认停在「视频」卡，只切一次的任务卡会随旧页面一起失效；
 * 而在页面加载过程中频繁点击又会触发重复加载。这里限定「最多补切几次 + 冷却多久」，
 * 让采样循环可以放心地每轮都尝试补切一次。
 */
export class QuestionTabGuard {
  private clicks = 0;
  private lastClickedAt = Number.NEGATIVE_INFINITY;

  constructor(
    private readonly tabs: TaskTabLocator,
    private readonly maxClicks = DEFAULT_MAX_CLICKS,
    private readonly cooldownMs = DEFAULT_CLICK_COOLDOWN_MS,
  ) {}

  /** 需要时补切一次题目卡，返回本次是否真的点击了 */
  ensure(now = Date.now()): boolean {
    if (this.clicks >= this.maxClicks) return false;
    if (now - this.lastClickedAt < this.cooldownMs) return false;
    if (this.tabs.ensureQuestionTab() !== 'switched') return false;
    this.clicks += 1;
    this.lastClickedAt = now;
    return true;
  }
}
