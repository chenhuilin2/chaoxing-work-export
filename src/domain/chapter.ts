// 学习通目录树其实是两级：
//   第一级「章」——`.posCatalog_select.firstLayer` 里的分组标题，页面上显示「1 担当复兴大任 成就时代新人」；
//   第二级「小节」——`.posCatalog_level` 下真正可点开的 `.posCatalog_name`，页面上显示「1.1 如何认识新方位」。
//
// 编号本来就印在页面上（两级都用 `<em class="posCatalog_sbar">`），所以导出与面板都直接用页面编号，
// 不再按导出顺序现编「第 N 章」。这个文件集中放两级标题的拼法，避免面板和三种导出各写一套。
import { normalizeInlineWhitespace } from '../utils/text';

/** 小节所属的章（目录第一级）；扁平目录没有这一级 */
export interface ChapterGroupInfo {
  readonly id: string;
  /** 页面上的章号（「1」）；页面没给编号时为空串 */
  readonly label: string;
  /** 章名（不含编号） */
  readonly title: string;
}

/** 章节标题的两级信息：面板列表、导出计划与三种渲染共用 */
export interface ChapterRef {
  /** 小节名（不含编号） */
  readonly title: string;
  /** 页面上的小节编号（「1.1」）；页面没给编号时为空串 */
  readonly label?: string;
  /** 所属章；扁平目录为 null / undefined */
  readonly group?: ChapterGroupInfo | null;
}

/** 「第一章」「第2章：」「第3讲」这类自带编号的名字：原样使用，避免叠成「第 1 章 第一章 …」 */
const CHAPTER_NUMBER_PATTERN = /^第\s*[0-9一二三四五六七八九十百千]+\s*[章节篇讲]/;

/** 「1.1」「2.3.1」这类小节编号（后面必须跟空白或结束，避免把「3D 打印」当成编号） */
const SECTION_NUMBER_PATTERN = /^[0-9]+(?:\.[0-9]+)*(?:\s|$)/;

/**
 * 小节标题：「1.1 如何认识新方位」。
 * 页面没给编号（老结构把编号写进了名字里，如 title="1.1 小节"）时退回名字本身，不叠加编号。
 */
export function sectionHeadingText(ref: ChapterRef): string {
  const name = normalizeInlineWhitespace(ref.title);
  const label = normalizeInlineWhitespace(ref.label ?? '');
  if (!name) return label;
  if (!label || name.startsWith(label) || SECTION_NUMBER_PATTERN.test(name)) return name;
  return `${label} ${name}`;
}

/**
 * 章标题：「第 1 章 担当复兴大任 成就时代新人」。
 *
 * 章名自带「第一章」这类编号时原样使用；页面没给章号时用导出内的章序兜底
 * （fallbackIndex 从 1 开始，传 0 表示「只要名字」）。
 */
export function groupHeadingText(group: ChapterGroupInfo, fallbackIndex = 0): string {
  const name = normalizeInlineWhitespace(group.title);
  if (!name) return '';
  if (CHAPTER_NUMBER_PATTERN.test(name)) return name;
  const label = normalizeInlineWhitespace(group.label);
  if (label) return `第 ${label} 章 ${name}`;
  return fallbackIndex > 0 ? `第 ${fallbackIndex} 章 ${name}` : name;
}
