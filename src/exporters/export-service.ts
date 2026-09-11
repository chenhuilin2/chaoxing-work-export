// 导出服务：组织逻辑移植自主脚本 chaoxing-work-export.user.js 的
// btnDownload / getOutputText / getChapterText，保持行为完全一致。
import type { ExportArtifact, ExportOptions } from '../domain/export-options';
import type { ExtractionResult } from '../domain/question';
import {
  type LegacyChapterData,
  type LegacyResultData,
  shuffleQuestions,
  toLegacyChapter,
  toLegacyResult,
} from './legacy-bridge';
import {
  formatOutput,
  formatOutputWithAnswers,
  formatWrongQuestionsTXT,
} from './text-formatter';
import {
  formatOutputMD,
  formatOutputWithAnswersMD,
  formatWrongQuestionsMD,
} from './markdown-formatter';
import { generateWordBlob, WORD_MIME } from './word-exporter';

// 原版 getOutputText：根据选项生成输出文本
// （打乱时题目和答案用打乱顺序重新生成，错题汇总始终用原始顺序）
function getOutputText(legacy: LegacyResultData, options: ExportOptions): string {
  const fmt = options.format;
  const withAnswers = options.withAnswers;
  const withWrong = options.withWrong;
  const doShuffle = options.shuffle;

  let base = '';
  if (doShuffle) {
    const shuffled = shuffleQuestions(legacy.results, legacy.typeOrder);
    if (fmt === 'md') {
      base = withAnswers
        ? formatOutputWithAnswersMD(shuffled, legacy.typeOrder)
        : formatOutputMD(shuffled, legacy.typeOrder);
    } else {
      base = withAnswers
        ? formatOutputWithAnswers(shuffled, legacy.typeOrder)
        : formatOutput(shuffled, legacy.typeOrder);
    }
  } else {
    if (fmt === 'md') {
      base = withAnswers
        ? formatOutputWithAnswersMD(legacy.results, legacy.typeOrder)
        : formatOutputMD(legacy.results, legacy.typeOrder);
    } else {
      base = withAnswers
        ? formatOutputWithAnswers(legacy.results, legacy.typeOrder)
        : formatOutput(legacy.results, legacy.typeOrder);
    }
  }

  if (withWrong) {
    const wrong =
      fmt === 'md'
        ? formatWrongQuestionsMD(legacy.results, legacy.typeOrder)
        : formatWrongQuestionsTXT(legacy.results, legacy.typeOrder);
    return base + wrong;
  }
  return base;
}

// 原版 getChapterText：为单个章节生成输出文本（用于分章节下载）
function getChapterText(chapter: LegacyChapterData, options: ExportOptions): string {
  const results = chapter.results || {};
  const typeOrder = chapter.typeOrder || [];
  const withAnswers = options.withAnswers;
  const withWrong = options.withWrong;
  const doShuffle = options.shuffle;
  const activeResults = doShuffle ? shuffleQuestions(results, typeOrder) : results;
  if (options.format === 'md') {
    if (withWrong) return formatWrongQuestionsMD(activeResults, typeOrder);
    if (withAnswers) return formatOutputWithAnswersMD(activeResults, typeOrder);
    return formatOutputMD(activeResults, typeOrder);
  }
  if (withWrong) return formatWrongQuestionsTXT(activeResults, typeOrder);
  if (withAnswers) return formatOutputWithAnswers(activeResults, typeOrder);
  return formatOutput(activeResults, typeOrder);
}

export class ExportService {
  async createArtifacts(
    sourceResult: ExtractionResult,
    options: ExportOptions,
  ): Promise<ExportArtifact[]> {
    const legacy = toLegacyResult(sourceResult);
    // 原版规则：文件名输入去掉扩展名后作为基础名
    const baseFilename = (options.filename || legacy.title || '学习通题目').replace(
      /\.(txt|md|docx)$/,
      '',
    );

    // Word 试卷导出（原版逻辑：题库导入必带答案并禁用打乱/错题；Word 暂不支持分章节，合并导出）
    if (options.format === 'word') {
      const isBankImport = options.bankImport;
      const doShuffle = !isBankImport && options.shuffle;
      const withWrong = !isBankImport && options.withWrong;
      const activeResults = doShuffle
        ? shuffleQuestions(legacy.results, legacy.typeOrder)
        : legacy.results;
      window.__xxt_failed_image_count = 0;
      const blob = await generateWordBlob(
        activeResults,
        legacy.typeOrder,
        legacy.title,
        isBankImport || options.withAnswers,
        withWrong,
        isBankImport,
      );
      const failedImages = window.__xxt_failed_image_count || 0;
      return [
        {
          blob,
          filename: `${baseFilename}.docx`,
          mimeType: WORD_MIME,
          failedImages,
        },
      ];
    }

    // TXT / MD 导出
    const ext = options.format === 'md' ? '.md' : '.txt';
    const mime = options.format === 'md' ? 'text/markdown;charset=utf-8' : 'text/plain;charset=utf-8';

    // 分章节下载：每个章节单独一个文件
    if (options.splitByChapter && sourceResult.chapters && sourceResult.chapters.length > 1) {
      const artifacts: ExportArtifact[] = [];
      for (let i = 0; i < sourceResult.chapters.length; i++) {
        const chapter = sourceResult.chapters[i];
        if (!chapter) continue;
        const chapterData = toLegacyChapter(chapter);
        // 生成单章文本
        const chText = getChapterText(chapterData, options);
        const safeTitle = (chapterData.title || `章节${i + 1}`).replace(/[\\/:*?"<>|]/g, '_');
        artifacts.push({
          blob: new Blob([chText], { type: mime }),
          filename: `${baseFilename}_${i + 1}_${safeTitle}${ext}`,
          mimeType: mime,
        });
      }
      return artifacts;
    }

    // 合并下载
    const text = getOutputText(legacy, options);
    return [
      {
        blob: new Blob([text], { type: mime }),
        filename: `${baseFilename}${ext}`,
        mimeType: mime,
      },
    ];
  }

  previewText(result: ExtractionResult, options: ExportOptions): string {
    // 原版行为：Word 格式不支持复制文本
    if (options.format === 'word') return '';
    return getOutputText(toLegacyResult(result), options);
  }
}
