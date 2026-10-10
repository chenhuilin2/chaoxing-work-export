// 导出文件的「文件划分 + 章节分组」纯逻辑（Word / TXT / Markdown 共用）。
//
// 单独成模块的原因：`word-exporter.ts` 依赖 CDN 注入的全局 `docx`，Node 单测里加载不了；
// 而「拆成几个文件、每个文件装什么、文件名怎么起、什么时候该打成一个压缩包」是最容易出错的一环，
// 所以把这段判定从渲染里拆出来，做成不碰 docx 的纯函数以便单测覆盖。
//
// 三种格式的文件划分必须只有这一处：TXT/Markdown 曾经在 export-service 里另写了一套
// 「合并导出 / 按章节拆分」，结果 Word 支持按章分节而 TXT/MD 仍把各章同题型混在一起。
//
// 目录树是两级（章 → 小节，见 domain/chapter.ts）：
//   - 正文/汇总里的标题：**章标题只在同一份文档含多个章时出现**（单章时文件大标题已经写明章名）；
//     小节标题只要页面给了编号就出现；老数据（扁平目录、历史记录）没有章信息，退回改动前的
//     「第 N 章 名字」单级写法，输出与改动前逐字一致。
//   - 「按章节拆分文件」按**章**拆：一章一个文件，文件里装该章的所有小节；压缩包阈值按章数算，
//     文件名就是章名本身（「第 1 章 担当复兴大任 成就时代新人.txt」），不再拼课程名与序号前缀。
//   - 标题上方的空行数由 `SectionGaps` 描述（章 3 行 / 小节 2 行，见 READABLE_GAPS），
//     TXT 与 Markdown 共用同一套间距规则，只有标题标记（`##` / `###`）各写各的。
import type { ChapterGroupInfo } from '../domain/chapter';
import { groupHeadingText, sectionHeadingText } from '../domain/chapter';
import type { ExportOptions } from '../domain/export-options';
import type { ChapterExtraction, ExtractionResult } from '../domain/question';
import {
  type LegacyResults,
  type LegacyTypeKey,
  shuffleQuestions,
  toLegacyChapter,
  toLegacyResult,
} from './legacy-bridge';

/**
 * 一份导出文件里的一个小节片段。
 *
 * 两级标题分别存放：`chapterHeading` 是所属章的标题，`heading` 是本节的小节标题，
 * 两者都可能为 null（老数据没有章信息、扁平目录没有小节编号）。
 */
export interface ExportSection {
  /** 所属章标题（如「第 1 章 担当复兴大任 成就时代新人」）；只有一级目录时为 null */
  readonly chapterHeading: string | null;
  /** 小节标题（含页面真编号，如「1.1 如何认识新方位」）；老数据/无编号时为 null */
  readonly heading: string | null;
  readonly results: LegacyResults;
  readonly typeOrder: readonly LegacyTypeKey[];
}

/** 一节内容在渲染序列里的位置（正文 / 答案 / 错题共用） */
export interface SectionGroupEntry<T> {
  /** 小节标题；null 表示这一节不再细分（老数据或扁平目录） */
  readonly heading: string | null;
  readonly value: T;
}

/** 一章的内容：章标题 + 它下面的各小节 */
export interface SectionGroup<T> {
  /** 章标题；null 表示这份文档只有一章（或目录只有一级），正文里不出现 */
  readonly heading: string | null;
  readonly entries: readonly SectionGroupEntry<T>[];
}

/** 一次导出要产出的一个文件 */
export interface ExportDocumentPlan {
  /** 文件大标题 */
  readonly title: string;
  /** 正文包含的章节；章数 > 1 时每章前插章标题 */
  readonly sections: readonly ExportSection[];
  /** 输出文件名（含扩展名） */
  readonly filename: string;
}

/** 逐章文件太多时改成的压缩包 */
export interface ExportArchivePlan {
  /** 压缩包文件名 */
  readonly filename: string;
  /** 包里装了各章的多少个文件 */
  readonly documentCount: number;
}

export interface ExportPlanSet {
  readonly documents: readonly ExportDocumentPlan[];
  /** 非空时表示这批文件应打成一个压缩包（太多文件逐个下载会被浏览器拦截） */
  readonly archive: ExportArchivePlan | null;
}

/**
 * 逐章导出的文件数阈值：超过它就不再逐个下载，改成一个压缩包。
 *
 * 章节练习动辄几十章，逐个下载既会触发浏览器的「多文件下载」拦截，也会把下载目录塞满；
 * 少数几章时逐个下载反而更方便直接打开。注意现在按「章」拆，所以阈值比的是**章数**。
 */
export const CHAPTER_ARCHIVE_THRESHOLD = 5;

/** 去掉用户可能输入的扩展名，得到输出文件的基础名（与 TXT/MD 导出规则一致） */
function baseName(result: ExtractionResult, options: ExportOptions): string {
  return (options.filename || result.title || '学习通题目').replace(/\.(txt|md|docx)$/, '');
}

/** 章节名可能含路径分隔符等文件名非法字符，替换后再落盘 */
function safeChapterTitle(title: string, index: number): string {
  return (title || `章节${index + 1}`).replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * 章名重名时补 `_2` / `_3` 后缀。
 *
 * 文件名不再带序号前缀，重名就会互相覆盖 —— 非法字符清洗（`1/2` → `1_2`）最容易撞车，
 * 所以这里兜一层。正常目录（章名自带「第 N 章」编号）不会触发。
 */
function uniqueChapterNames(names: readonly string[]): string[] {
  const used = new Map<string, number>();
  return names.map((name) => {
    const count = (used.get(name) ?? 0) + 1;
    used.set(name, count);
    return count === 1 ? name : `${name}_${count}`;
  });
}

/** 章节名自带编号（「第一章」「第2章：」「第3讲」）的判定，用于避免重复编号 */
const CHAPTER_NUMBER_PATTERN = /^第\s*[0-9一二三四五六七八九十百千]+\s*[章节篇讲]/;

/**
 * 没有章 / 小节信息时的标题（老数据、扁平目录）。
 *
 * 章节名常自带编号（「第一章 马克思主义」「第2章：马克思/恩格斯」），此时原样使用，
 * 避免出现「第 1 章 第一章 马克思主义」这种叠字；没有编号时才补上「第 N 章」。
 */
export function chapterHeadingText(title: string, index: number): string {
  const name = (title || '').trim() || `章节${index + 1}`;
  return CHAPTER_NUMBER_PATTERN.test(name) ? name : `第 ${index + 1} 章 ${name}`;
}

/** 同一份文档里含几个章（章标题只在 > 1 时打进正文） */
export function countChapters(sections: readonly ExportSection[]): number {
  const headings = new Set<string>();
  for (const section of sections) {
    if (section.chapterHeading) headings.add(section.chapterHeading);
  }
  return headings.size;
}

/**
 * 把 sections 折成「章 → 小节」两级，并借机跳过没有内容的节（`pick` 返回 null）。
 *
 * 三种格式的正文 / 答案汇总 / 错题汇总都走这一个函数，所以「章标题只在多章时出现」
 * 「没有错题的章连标题一起省掉」这类规则只有一处实现。
 */
export function groupSections<T>(
  sections: readonly ExportSection[],
  pick: (section: ExportSection, index: number) => T | null,
): SectionGroup<T>[] {
  const multiChapter = countChapters(sections) > 1;
  const groups: { heading: string | null; entries: SectionGroupEntry<T>[] }[] = [];

  sections.forEach((section, index) => {
    const value = pick(section, index);
    if (value === null) return;

    const heading = multiChapter ? section.chapterHeading : null;
    const last = groups[groups.length - 1];
    if (last && last.heading === heading) {
      last.entries.push({ heading: section.heading, value });
      return;
    }
    groups.push({ heading, entries: [{ heading: section.heading, value }] });
  });

  return groups;
}

/**
 * 按章切分 sections：每章一个区块，区块内是该章按顺序排列的小节。
 *
 * 三种格式的「章内」渲染（题目 → 该章答案 → 该章错题）都以块为单位，
 * 章标题为 null 表示这份文档只有一章（或目录只有一级）。
 */
export function chapterBlocks(sections: readonly ExportSection[]): SectionGroup<ExportSection>[] {
  return groupSections(sections, (section) => section);
}

/** 把「标题 + 该节正文」拼成一段；该节没有正文时只留标题，不留多余空行 */
export function joinHeading(heading: string | null, body: string): string {
  if (!heading) return body;
  return body ? `${heading}\n\n${body}` : heading;
}

/** n 个空行对应的换行串（段落之间空 n 行 = n + 1 个换行） */
export function blankLines(count: number): string {
  return '\n'.repeat(count + 1);
}

/**
 * 标题上方的空行数。
 *
 * 拿「标题上方空几行」而不是「用什么分隔符」来描述，是因为 TXT 与 Markdown 的
 * 空行规则完全一致（要求「章标题上方多空两行、小节标题上方多空一行」），
 * 而两者只有 Markdown 的 `#` 标记不同 —— 标记由各格式自己加，间距规则共用。
 */
export interface SectionGaps {
  /** 章标题上方空几行（文档第一处不空） */
  readonly chapter: number;
  /** 小节标题上方空几行（每章第一个小节不空） */
  readonly section: number;
}

/** 改动前的间距：章 / 小节标题上方各 1 个空行 */
export const DEFAULT_GAPS: SectionGaps = { chapter: 1, section: 1 };

/**
 * 面向阅读的间距：章标题上方空 3 行、小节标题上方空 2 行。
 *
 * 多章节导出时层级本来就深（章 → 小节 → 题型），只空一行时分不出「换章」与「换小节」，
 * 章之间多空一行、小节之间多空一行，滚动时一眼能看出自己在哪一级。
 */
export const READABLE_GAPS: SectionGaps = { chapter: 3, section: 2 };

/**
 * 章标题 → 小节标题 → 正文，逐级拼接；标题文本由各格式自己带上标记
 * （TXT 是纯文本、Markdown 是 `##` / `###`），分段规则共用。
 *
 * 标题与其正文之间始终空 1 行（`joinHeading`）；`gaps` 只管**标题上方**的空行，
 * 也就是「上一段正文」与「下一个标题」之间的留白。
 */
export function renderSectionGroups(
  groups: readonly SectionGroup<string>[],
  gaps: SectionGaps = DEFAULT_GAPS,
): string {
  return groups
    .map((group) => {
      const body = group.entries
        .map((entry) => joinHeading(entry.heading, entry.value))
        .join(blankLines(gaps.section));
      return joinHeading(group.heading, body);
    })
    .join(blankLines(gaps.chapter));
}

/**
 * 把章节数据摊成渲染用的 sections。
 *
 * 有章信息的两级目录：章标题用页面上的章号（「第 1 章 …」），小节标题用页面上的小节编号
 * （「1.1 …」）；没有章信息时退回改动前的「第 N 章 名字」单级写法。
 */
function buildSections(chapters: readonly ChapterExtraction[], shuffle: boolean): ExportSection[] {
  const groupOrder = new Map<string, number>();

  return chapters.map((chapter, index) => {
    const data = toLegacyChapter(chapter);
    const results = shuffle ? shuffleQuestions(data.results, data.typeOrder) : data.results;
    const group: ChapterGroupInfo | null = chapter.group ?? null;

    if (!group) {
      // 扁平目录 / 老数据：这一节自成一段，标题沿用「第 N 章 名字」
      return {
        chapterHeading: chapterHeadingText(data.title, index),
        heading: null,
        results,
        typeOrder: data.typeOrder,
      };
    }

    let order = groupOrder.get(group.id);
    if (order === undefined) {
      order = groupOrder.size + 1;
      groupOrder.set(group.id, order);
    }
    return {
      chapterHeading: groupHeadingText(group, order) || null,
      heading: sectionHeadingText(chapter) || null,
      results,
      typeOrder: data.typeOrder,
    };
  });
}

/** 按「章」把小节分组，保持首次出现的顺序（按章拆文件时一章一份） */
function groupChapters(
  chapters: readonly ChapterExtraction[],
): { readonly group: ChapterGroupInfo | null; readonly chapters: ChapterExtraction[] }[] {
  const groups: {
    readonly key: string;
    readonly group: ChapterGroupInfo | null;
    readonly chapters: ChapterExtraction[];
  }[] = [];

  chapters.forEach((chapter, index) => {
    const group = chapter.group ?? null;
    const key = group ? group.id : `#${index}`;
    const last = groups[groups.length - 1];
    if (last && last.key === key) {
      last.chapters.push(chapter);
      return;
    }
    groups.push({ key, group, chapters: [chapter] });
  });

  return groups.map(({ group, chapters: items }) => ({ group, chapters: items }));
}

/**
 * 规划本次导出要产出哪些文件：
 * - 只有一章（或页面没有章节信息）→ 一个文件、正文不分节，与改动前逐字一致；
 * - 勾选「按章节拆分文件」且多章节 → **一章一个文件**（文件内含该章的所有小节）；
 *   章数超过阈值时另附一个压缩包计划；
 * - 否则多章节单文件 → 正文按「章 → 小节 → 题型」分节，避免把各章节的同题型题目混在同一大题里。
 *
 * 题库导入（智能导入）格式不认章节标题，所以多章节时也不在单文件里插章节标题，
 * 只按需要拆文件——拆出来的每个文件仍然只有一章的题目。
 */
export function planExportDocuments(
  result: ExtractionResult,
  options: ExportOptions,
  ext: string,
): ExportPlanSet {
  const base = baseName(result, options);
  // 题库导入必带答案并禁用打乱，与 export-service 的既有规则保持一致
  const shuffle = !options.bankImport && options.shuffle;
  const chapters = result.chapters ?? [];
  const legacy = toLegacyResult(result);

  // 单章节（或没有章节信息）：一个文件、正文不分节
  if (chapters.length <= 1) {
    const section: ExportSection = {
      chapterHeading: null,
      heading: null,
      results: shuffle ? shuffleQuestions(legacy.results, legacy.typeOrder) : legacy.results,
      typeOrder: legacy.typeOrder,
    };
    return {
      documents: [{ title: legacy.title, sections: [section], filename: `${base}${ext}` }],
      archive: null,
    };
  }

  // 按章节拆分：一章一个文件（章内的多个小节装在同一个文件里）
  if (options.splitByChapter) {
    const blocks = groupChapters(chapters).map((item, index) => {
      const sections = buildSections(item.chapters, shuffle);
      const first = sections[0];
      // 文件大标题即章名：有章信息时用「第 N 章 …」，扁平目录仍用该节自己的名字
      const title = (item.group ? first?.chapterHeading : null) ?? item.chapters[0]?.title ?? '';
      return { title: title || `章节${index + 1}`, sections };
    });
    // 文件名直接用章名（「第 1 章 担当复兴大任 成就时代新人.txt」）：
    // 课程名与序号前缀在解压后的文件夹里只会碍事，章名本身已经含编号、彼此不重名。
    const names = uniqueChapterNames(
      blocks.map((block, index) => safeChapterTitle(block.title, index)),
    );
    const documents = blocks.map((block, index) => ({
      title: block.title,
      sections: block.sections,
      filename: `${names[index]}${ext}`,
    }));
    return {
      documents,
      // 章太多：不再逐个下载，全部打进一个压缩包
      archive:
        documents.length > CHAPTER_ARCHIVE_THRESHOLD
          ? { filename: `${base}.zip`, documentCount: documents.length }
          : null,
    };
  }

  // 多章节单文件：正文按章节分节
  const sections: ExportSection[] = options.bankImport
    ? [
        {
          chapterHeading: null,
          heading: null,
          results: shuffle ? shuffleQuestions(legacy.results, legacy.typeOrder) : legacy.results,
          typeOrder: legacy.typeOrder,
        },
      ]
    : buildSections(chapters, shuffle);

  return {
    documents: [{ title: legacy.title, sections, filename: `${base}${ext}` }],
    archive: null,
  };
}
