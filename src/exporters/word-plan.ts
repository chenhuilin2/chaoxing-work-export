// Word 导出的「文件划分 + 章节分组」纯逻辑。
//
// 单独成模块的原因：`word-exporter.ts` 依赖 CDN 注入的全局 `docx`，Node 单测里加载不了；
// 而「按章节拆分」最容易出错的一环恰恰是「拆成几个文件、每个文件装什么、文件名怎么起」，
// 所以把这段判定从渲染里拆出来，做成不碰 docx 的纯函数以便单测覆盖。
import type { ExportOptions } from '../domain/export-options';
import type { ExtractionResult } from '../domain/question';
import {
  type LegacyChapterData,
  type LegacyResults,
  type LegacyTypeKey,
  shuffleQuestions,
  toLegacyChapter,
  toLegacyResult,
} from './legacy-bridge';

/** 一份 Word 文档里的一个章节片段（单章节文档时 title 即文档标题） */
export interface WordSection {
  readonly title: string;
  readonly results: LegacyResults;
  readonly typeOrder: readonly LegacyTypeKey[];
}

/** 一次 Word 导出要产出的一个 .docx */
export interface WordDocumentPlan {
  /** 文档大标题 */
  readonly title: string;
  /** 正文包含的章节；长度 > 1 时每章前插章节标题 */
  readonly sections: readonly WordSection[];
  /** 输出文件名（含扩展名） */
  readonly filename: string;
}

/** 去掉用户可能输入的扩展名，得到输出文件的基础名（与 TXT/MD 导出规则一致） */
function baseName(result: ExtractionResult, options: ExportOptions): string {
  return (options.filename || result.title || '学习通题目').replace(/\.(txt|md|docx)$/, '');
}

/** 章节名可能含路径分隔符等文件名非法字符，替换后再落盘 */
function safeChapterTitle(title: string, index: number): string {
  return (title || `章节${index + 1}`).replace(/[\\/:*?"<>|]/g, '_');
}

function toSection(chapter: LegacyChapterData, shuffle: boolean): WordSection {
  return {
    title: chapter.title,
    results: shuffle ? shuffleQuestions(chapter.results, chapter.typeOrder) : chapter.results,
    typeOrder: chapter.typeOrder,
  };
}

/**
 * 规划本次 Word 导出要产出哪些文件：
 * - 勾选「按章节拆分文件」且结果是多章节 → **一章一个 .docx**（与 TXT/Markdown 行为一致）；
 * - 否则只产出一个 .docx：**多章节时正文按「章节 → 题型」分节**，单章节时与原先完全一致。
 *
 * 题库导入（智能导入）格式不认章节标题，所以多章节时也不在单文件里插章节标题，
 * 只按需要拆文件——拆出来的每个文件仍然只有一章的题目。
 */
export function planWordDocuments(
  result: ExtractionResult,
  options: ExportOptions,
): WordDocumentPlan[] {
  const base = baseName(result, options);
  // 题库导入必带答案并禁用打乱，与 export-service 的既有规则保持一致
  const shuffle = !options.bankImport && options.shuffle;
  const chapters = result.chapters ?? [];

  // 分章节下载：每个章节单独一个文件
  if (options.splitByChapter && chapters.length > 1) {
    return chapters.map((chapter, index) => {
      const data = toLegacyChapter(chapter);
      const title = safeChapterTitle(data.title, index);
      return {
        title: data.title || title,
        sections: [toSection(data, shuffle)],
        filename: `${base}_${index + 1}_${title}.docx`,
      };
    });
  }

  const legacy = toLegacyResult(result);
  // 多章节单文件：正文按章节分节，避免把各章节的同题型题目混在同一大题里
  const groupedByChapter = chapters.length > 1 && !options.bankImport;
  const sections: WordSection[] = groupedByChapter
    ? chapters.map((chapter) => toSection(toLegacyChapter(chapter), shuffle))
    : [
        {
          title: legacy.title,
          results: shuffle ? shuffleQuestions(legacy.results, legacy.typeOrder) : legacy.results,
          typeOrder: legacy.typeOrder,
        },
      ];

  return [{ title: legacy.title, sections, filename: `${base}.docx` }];
}
