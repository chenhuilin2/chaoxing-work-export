import type { ExportArtifact, ExportOptions } from '../domain/export-options';
import {
  buildStatistics,
  deriveTypeOrder,
  type ChapterExtraction,
  type ExtractionResult,
} from '../domain/question';
import { sanitizeFilename } from '../utils/text';
import { formatMarkdown } from './markdown-formatter';
import { transformResult } from './result-transformer';
import { formatText } from './text-formatter';
import { WordExporter } from './word-exporter';

function extension(format: ExportOptions['format']): string {
  if (format === 'word') return 'docx';
  return format;
}

function mimeType(format: ExportOptions['format']): string {
  if (format === 'word') return WordExporter.mimeType;
  if (format === 'md') return 'text/markdown;charset=utf-8';
  return 'text/plain;charset=utf-8';
}

function chapterResult(parent: ExtractionResult, chapter: ChapterExtraction): ExtractionResult {
  return {
    title: `${parent.title} - ${chapter.title}`,
    questions: chapter.questions,
    typeOrder: deriveTypeOrder(chapter.questions),
    statistics: buildStatistics(chapter.questions),
    sourceUrl: chapter.sourceUrl,
    extractor: chapter.extractor,
    extractedAt: parent.extractedAt,
  };
}

export class ExportService {
  async createArtifacts(
    sourceResult: ExtractionResult,
    options: ExportOptions,
  ): Promise<ExportArtifact[]> {
    const result = transformResult(sourceResult, options.shuffle);
    const baseName = sanitizeFilename(options.filename || result.title);

    if (options.splitByChapter && result.chapters && result.chapters.length > 1) {
      const artifacts: ExportArtifact[] = [];
      for (let index = 0; index < result.chapters.length; index += 1) {
        const chapter = result.chapters[index] as ChapterExtraction;
        const chapterName = sanitizeFilename(
          `${baseName}-${String(index + 1).padStart(2, '0')}-${chapter.title}`,
        );
        artifacts.push(await this.createOne(chapterResult(result, chapter), options, chapterName));
      }
      return artifacts;
    }

    return [await this.createOne(result, options, baseName)];
  }

  previewText(result: ExtractionResult, options: ExportOptions): string {
    const transformed = transformResult(result, options.shuffle);
    return options.format === 'md'
      ? formatMarkdown(transformed, options)
      : formatText(transformed, options);
  }

  private async createOne(
    result: ExtractionResult,
    options: ExportOptions,
    baseName: string,
  ): Promise<ExportArtifact> {
    const filename = `${baseName}.${extension(options.format)}`;
    if (options.format === 'word') {
      const generated = await new WordExporter().export(result, options);
      return {
        blob: generated.blob,
        filename,
        mimeType: mimeType(options.format),
        failedImages: generated.failedImages,
      };
    }

    const content =
      options.format === 'md' ? formatMarkdown(result, options) : formatText(result, options);
    return {
      blob: new Blob([content], { type: mimeType(options.format) }),
      filename,
      mimeType: mimeType(options.format),
    };
  }
}
