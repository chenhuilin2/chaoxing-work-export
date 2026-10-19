// 导出服务：组织逻辑移植自主脚本 chaoxing-work-export.user.js 的
// btnDownload / getOutputText / getChapterText，保持行为完全一致。
//
// 文件划分（拆几个文件、每个文件装哪些章节、要不要打成一个压缩包）全部由 `export-plan` 决定，
// 这里只负责「按计划逐个渲染，再按需要打包下载」——三种格式走同一份计划，规则不会再各自跑偏。
import type { ExportArtifact, ExportFormat, ExportOptions } from '../domain/export-options';
import type { ExtractionResult } from '../domain/question';
import { type ExportDocumentPlan, planExportDocuments } from './export-plan';
import { formatSectionsOutputMD } from './markdown-formatter';
import { formatSectionsOutput } from './text-formatter';
import { WORD_MIME, generateWordBlob } from './word-exporter';
import { type ZipEntry, createZipBlob } from './zip-archive';

const EXTENSIONS: Record<ExportFormat, string> = {
  word: '.docx',
  txt: '.txt',
  md: '.md',
};

const TEXT_MIME: Record<'txt' | 'md', string> = {
  txt: 'text/plain;charset=utf-8',
  md: 'text/markdown;charset=utf-8',
};

const ARCHIVE_MIME = 'application/zip';

/**
 * 把逐章文件打成一个压缩包。
 *
 * 包内先放一个与压缩包同名的目录条目，章节文件都挂在这个目录下 ——
 * 用户解压后拿到的是「一个装着所有章节文件的文件夹」，而不是几十个散文件。
 */
async function bundleArchive(
  artifacts: readonly ExportArtifact[],
  archiveFilename: string,
): Promise<ExportArtifact> {
  const folder = archiveFilename.replace(/\.zip$/i, '') || '章节文件';
  const entries: ZipEntry[] = [{ name: `${folder}/`, data: new Uint8Array(0) }];
  for (const artifact of artifacts) {
    entries.push({
      name: `${folder}/${artifact.filename}`,
      data: new Uint8Array(await artifact.blob.arrayBuffer()),
    });
  }
  return {
    blob: createZipBlob(entries),
    filename: archiveFilename,
    mimeType: ARCHIVE_MIME,
    // 图片计数取各文件的合计（Word 逐章生成时是全局累加的）
    failedImages: artifacts.reduce((total, artifact) => total + (artifact.failedImages ?? 0), 0),
    archiveEntries: artifacts.length,
  };
}

export class ExportService {
  async createArtifacts(
    sourceResult: ExtractionResult,
    options: ExportOptions,
  ): Promise<ExportArtifact[]> {
    const plan = planExportDocuments(sourceResult, options, EXTENSIONS[options.format]);
    const artifacts =
      options.format === 'word'
        ? await this.renderWord(plan.documents, options)
        : this.renderText(plan.documents, options);

    // 逐章拆出的文件太多：不再逐个下载，打成一个压缩包
    if (plan.archive) return [await bundleArchive(artifacts, plan.archive.filename)];
    return artifacts;
  }

  previewText(result: ExtractionResult, options: ExportOptions): string {
    // 原版行为：Word 格式不支持复制文本
    if (options.format === 'word') return '';
    // 复制的内容与下载的合并文件一致：强制走「多章节单文件」这一份计划
    const { documents } = planExportDocuments(
      result,
      { ...options, splitByChapter: false, splitBySection: false },
      EXTENSIONS[options.format],
    );
    const document = documents[0];
    if (!document) return '';
    return options.format === 'md'
      ? formatSectionsOutputMD(document.sections, options)
      : formatSectionsOutput(document.sections, options);
  }

  /** Word：每个计划一份 .docx（题库导入必带答案、禁打乱/错题） */
  private async renderWord(
    documents: readonly ExportDocumentPlan[],
    options: ExportOptions,
  ): Promise<ExportArtifact[]> {
    const isBankImport = options.bankImport;
    const withWrong = !isBankImport && options.withWrong;
    const artifacts: ExportArtifact[] = [];
    window.__xxt_failed_image_count = 0;
    for (const document of documents) {
      const failedBefore = window.__xxt_failed_image_count || 0;
      const blob = await generateWordBlob(document.title, document.sections, {
        withAnswers: isBankImport || options.withAnswers,
        withWrong,
        bankImport: isBankImport,
      });
      artifacts.push({
        blob,
        filename: document.filename,
        mimeType: WORD_MIME,
        // 图片计数是全局累加的，这里取本文件的增量
        failedImages: (window.__xxt_failed_image_count || 0) - failedBefore,
      });
    }
    return artifacts;
  }

  /** TXT / Markdown：计划里的每个文件各渲染一份文本 */
  private renderText(
    documents: readonly ExportDocumentPlan[],
    options: ExportOptions,
  ): ExportArtifact[] {
    const format = options.format === 'md' ? 'md' : 'txt';
    const mime = TEXT_MIME[format];
    return documents.map((document) => {
      const text =
        format === 'md'
          ? formatSectionsOutputMD(document.sections, options)
          : formatSectionsOutput(document.sections, options);
      return {
        blob: new Blob([text], { type: mime }),
        filename: document.filename,
        mimeType: mime,
      };
    });
  }
}
