import type { ExportOptions } from '../domain/export-options';
import {
  QUESTION_TYPE_LONG_LABELS,
  groupQuestions,
  type ExtractionResult,
  type Question,
  type RichContent,
  type RichPart,
} from '../domain/question';
import { richContentToText } from '../extractors/rich-content';
import { normalizeAnswer } from '../utils/text';

interface ImageAsset {
  readonly data: Uint8Array;
  readonly type: 'png' | 'jpg' | 'gif' | 'bmp';
  readonly width: number;
  readonly height: number;
}

interface RichRunOptions {
  readonly color?: string;
  readonly bold?: boolean;
}

const WORD_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function imageType(mimeType: string, url: string): ImageAsset['type'] {
  const value = `${mimeType} ${url}`.toLowerCase();
  if (value.includes('gif')) return 'gif';
  if (value.includes('bmp')) return 'bmp';
  if (value.includes('jpeg') || value.includes('jpg')) return 'jpg';
  return 'png';
}

function absoluteImageUrl(url: string): string {
  try {
    return new URL(url, document.baseURI).href;
  } catch {
    return url;
  }
}

async function imageDimensions(blob: Blob): Promise<{ readonly width: number; readonly height: number }> {
  if (typeof createImageBitmap === 'function') {
    const bitmap = await createImageBitmap(blob);
    const dimensions = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return dimensions;
  }

  return new Promise((resolve, reject) => {
    const objectUrl = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      resolve({ width: image.naturalWidth || 300, height: image.naturalHeight || 200 });
    };
    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error('无法读取图片尺寸'));
    };
    image.src = objectUrl;
  });
}

function fitImage(
  width: number,
  height: number,
  maxWidth = 420,
  maxHeight = 260,
): { readonly width: number; readonly height: number } {
  const safeWidth = Math.max(1, width || maxWidth);
  const safeHeight = Math.max(1, height || maxHeight);
  const ratio = Math.min(1, maxWidth / safeWidth, maxHeight / safeHeight);
  return {
    width: Math.max(1, Math.round(safeWidth * ratio)),
    height: Math.max(1, Math.round(safeHeight * ratio)),
  };
}

class ImageRepository {
  private readonly cache = new Map<string, Promise<ImageAsset | null>>();
  private failures = 0;

  get failedCount(): number {
    return this.failures;
  }

  load(url: string): Promise<ImageAsset | null> {
    const absolute = absoluteImageUrl(url);
    const existing = this.cache.get(absolute);
    if (existing) return existing;
    const request = this.fetch(absolute);
    this.cache.set(absolute, request);
    return request;
  }

  private async fetch(url: string): Promise<ImageAsset | null> {
    const controller = new AbortController();
    const timer = window.setTimeout(() => controller.abort(), 10_000);
    try {
      const response = await fetch(url, { credentials: 'include', signal: controller.signal });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const blob = await response.blob();
      const sourceDimensions = await imageDimensions(blob);
      const dimensions = fitImage(sourceDimensions.width, sourceDimensions.height);
      return {
        data: new Uint8Array(await blob.arrayBuffer()),
        type: imageType(blob.type, url),
        width: dimensions.width,
        height: dimensions.height,
      };
    } catch {
      this.failures += 1;
      return null;
    } finally {
      window.clearTimeout(timer);
    }
  }
}

function runText(value: string, options: Record<string, unknown> = {}): unknown {
  return new docx.TextRun({ text: value, font: 'Microsoft YaHei', ...options });
}

function pageBreakParagraph(): unknown {
  return new docx.Paragraph({ children: [new docx.PageBreak()] });
}

function normalizeBankAnswer(question: Question): string {
  const plain = richContentToText(question.correctAnswer).trim();
  if (!plain) return '';
  const normalized = normalizeAnswer(plain);
  if (question.type === 'true-false') {
    if (normalized === 'TRUE') return '对';
    if (normalized === 'FALSE') return '错';
  }
  if (question.type === 'multiple-choice') {
    const letters = normalized.match(/[A-H]/g);
    if (letters && letters.length > 0) return [...new Set(letters)].join('，');
  }
  return plain.replace(/^(?:正确答案|参考答案|答案)\s*[:：]?\s*/u, '');
}

function bankStem(content: RichContent): RichContent {
  return content.map((part) => {
    if (part.type !== 'text') return part;
    return {
      ...part,
      text: part.text.replace(/[（(]\s{2,}[）)]/gu, '（ ）'),
    };
  });
}

/** Converts a normalized extraction snapshot into a browser-generated .docx file. */
export class WordExporter {
  private readonly images = new ImageRepository();

  async export(result: ExtractionResult, options: ExportOptions): Promise<{ blob: Blob; failedImages: number }> {
    if (typeof docx === 'undefined') {
      throw new Error('Word 组件未加载，请刷新页面后重试');
    }

    const children = options.bankImport
      ? await this.buildBankImport(result)
      : await this.buildStandardDocument(result, options);
    const wordDocument = new docx.Document({
      creator: 'Chaoxing Work Export',
      title: result.title,
      description: '由 Chaoxing Work Export 生成的学习资料',
      styles: {
        default: {
          document: {
            run: { font: 'Microsoft YaHei', size: 22, color: '222222' },
            paragraph: { spacing: { after: 120, line: 360 } },
          },
        },
        paragraphStyles: [
          {
            id: 'CweTitle',
            name: 'CWE Title',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: 'Microsoft YaHei', size: 36, bold: true, color: '1D4ED8' },
            paragraph: { spacing: { before: 120, after: 220 }, alignment: 'center' },
          },
          {
            id: 'CweHeading1',
            name: 'CWE Heading 1',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: 'Microsoft YaHei', size: 30, bold: true, color: '2563EB' },
            paragraph: { spacing: { before: 260, after: 140 }, keepNext: true },
          },
          {
            id: 'CweHeading2',
            name: 'CWE Heading 2',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: { font: 'Microsoft YaHei', size: 26, bold: true, color: '1E3A8A' },
            paragraph: { spacing: { before: 220, after: 100 }, keepNext: true },
          },
        ],
      },
      sections: [
        {
          properties: {
            page: {
              size: { width: 11906, height: 16838 },
              margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
            },
          },
          children,
        },
      ],
    });

    return {
      blob: await docx.Packer.toBlob(wordDocument),
      failedImages: this.images.failedCount,
    };
  }

  static readonly mimeType = WORD_MIME;

  private async buildStandardDocument(
    result: ExtractionResult,
    options: ExportOptions,
  ): Promise<unknown[]> {
    const children: unknown[] = [
      new docx.Paragraph({ style: 'CweTitle', children: [runText(result.title, { bold: true })] }),
      new docx.Paragraph({
        alignment: docx.AlignmentType?.CENTER,
        children: [
          runText(
            `共 ${result.statistics.total} 题　错题 ${result.statistics.wrong} 题　含答案 ${result.statistics.withCorrectAnswer} 题`,
            { color: '64748B', size: 19 },
          ),
        ],
      }),
    ];

    let number = 1;
    if (result.chapters && result.chapters.length > 0) {
      for (const chapter of result.chapters) {
        children.push(
          new docx.Paragraph({
            style: 'CweHeading1',
            children: [runText(chapter.title, { bold: true })],
          }),
        );
        number = await this.appendGroupedQuestions(
          children,
          chapter.questions,
          chapter.typeOrder,
          number,
        );
      }
    } else {
      await this.appendGroupedQuestions(children, result.questions, result.typeOrder, number);
    }

    if (options.withAnswers) {
      children.push(pageBreakParagraph());
      await this.appendAnswerAppendix(children, result, options.includeAnalysis);
    }
    if (options.withWrong && result.statistics.wrong > 0) {
      children.push(pageBreakParagraph());
      await this.appendWrongAppendix(children, result, options.includeAnalysis);
    }
    return children;
  }

  private async appendGroupedQuestions(
    children: unknown[],
    questions: readonly Question[],
    typeOrder: ExtractionResult['typeOrder'],
    startNumber: number,
  ): Promise<number> {
    const groups = groupQuestions(questions);
    let number = startNumber;
    for (const type of typeOrder) {
      const typedQuestions = groups[type];
      if (typedQuestions.length === 0) continue;
      children.push(
        new docx.Paragraph({
          style: 'CweHeading2',
          children: [
            runText(`${QUESTION_TYPE_LONG_LABELS[type]}（共 ${typedQuestions.length} 题）`, {
              bold: true,
            }),
          ],
        }),
      );
      for (const question of typedQuestions) {
        children.push(...(await this.buildQuestion(question, number)));
        number += 1;
      }
    }
    return number;
  }

  private async buildQuestion(question: Question, number: number): Promise<unknown[]> {
    const stemRuns = [
      runText(`${number}. `, { bold: true }),
      ...(await this.richRuns(question.stem)),
    ];
    const children: unknown[] = [
      new docx.Paragraph({
        children: stemRuns,
        border: {
          top: {
            color: 'CBD5E1',
            space: 8,
            style: docx.BorderStyle?.SINGLE ?? 'single',
            size: 6,
          },
        },
        spacing: { before: 140, after: 100 },
        keepNext: question.options.length > 0,
      }),
    ];

    for (const option of question.options) {
      children.push(
        new docx.Paragraph({
          children: [
            runText(`${option.key}. `, { bold: true }),
            ...(await this.richRuns(option.content)),
          ],
          indent: { left: 620, hanging: 220 },
          spacing: { after: 60 },
        }),
      );
    }
    return children;
  }

  private async appendAnswerAppendix(
    children: unknown[],
    result: ExtractionResult,
    includeAnalysis: boolean,
  ): Promise<void> {
    children.push(
      new docx.Paragraph({ style: 'CweHeading1', children: [runText('参考答案', { bold: true })] }),
    );
    const groups = groupQuestions(result.questions);
    const numberById = new Map(result.questions.map((question, index) => [question.id, index + 1]));
    for (const type of result.typeOrder) {
      const questions = groups[type];
      if (questions.length === 0) continue;
      children.push(
        new docx.Paragraph({
          style: 'CweHeading2',
          children: [runText(QUESTION_TYPE_LONG_LABELS[type], { bold: true })],
        }),
      );
      for (const question of questions) {
        const number = numberById.get(question.id) ?? 0;
        const answerRuns = question.correctAnswer.length
          ? await this.richRuns(question.correctAnswer, { color: '00A870', bold: true })
          : [runText('暂无答案', { color: '94A3B8' })];
        children.push(
          new docx.Paragraph({
            children: [runText(`${number}. `, { bold: true }), ...answerRuns],
            spacing: { after: includeAnalysis && question.analysis.length > 0 ? 30 : 100 },
          }),
        );
        if (includeAnalysis && question.analysis.length > 0) {
          children.push(
            new docx.Paragraph({
              children: [
                runText('解析：', { bold: true, color: '475569' }),
                ...(await this.richRuns(question.analysis, { color: '475569' })),
              ],
              indent: { left: 360 },
              spacing: { after: 120 },
            }),
          );
        }
      }
    }
  }

  private async appendWrongAppendix(
    children: unknown[],
    result: ExtractionResult,
    includeAnalysis: boolean,
  ): Promise<void> {
    children.push(
      new docx.Paragraph({ style: 'CweHeading1', children: [runText('错题汇总', { bold: true })] }),
    );
    const wrong = result.questions
      .map((question, index) => ({ question, number: index + 1 }))
      .filter(({ question }) => question.isWrong);

    for (const { question, number } of wrong) {
      children.push(...(await this.buildQuestion(question, number)));
      children.push(
        new docx.Paragraph({
          children: [
            runText('我的答案：', { bold: true, color: 'DC2626' }),
            ...(question.userAnswer.length
              ? await this.richRuns(question.userAnswer, { color: 'DC2626' })
              : [runText('未作答', { color: 'DC2626' })]),
          ],
          indent: { left: 360 },
        }),
        new docx.Paragraph({
          children: [
            runText('正确答案：', { bold: true, color: '00A870' }),
            ...(question.correctAnswer.length
              ? await this.richRuns(question.correctAnswer, { color: '00A870' })
              : [runText('暂无答案', { color: '94A3B8' })]),
          ],
          indent: { left: 360 },
        }),
      );
      if (includeAnalysis && question.analysis.length > 0) {
        children.push(
          new docx.Paragraph({
            children: [
              runText('解析：', { bold: true, color: '475569' }),
              ...(await this.richRuns(question.analysis, { color: '475569' })),
            ],
            indent: { left: 360 },
          }),
        );
      }
    }
  }

  private async buildBankImport(result: ExtractionResult): Promise<unknown[]> {
    const children: unknown[] = [
      new docx.Paragraph({ style: 'CweTitle', children: [runText(result.title, { bold: true })] }),
    ];
    for (let index = 0; index < result.questions.length; index += 1) {
      const question = result.questions[index] as Question;
      children.push(
        new docx.Paragraph({
          children: [
            runText(`${index + 1}.【${QUESTION_TYPE_LONG_LABELS[question.type]}】`, { bold: true }),
            ...(await this.richRuns(bankStem(question.stem))),
          ],
          spacing: { before: 140, after: 80 },
        }),
      );
      for (const option of question.options) {
        children.push(
          new docx.Paragraph({
            children: [runText(`${option.key}. `), ...(await this.richRuns(option.content))],
            indent: { left: 420 },
            spacing: { after: 40 },
          }),
        );
      }
      children.push(
        new docx.Paragraph({
          children: [
            runText('答案：', { bold: true }),
            runText(normalizeBankAnswer(question) || '暂无答案'),
          ],
          spacing: { after: 140 },
        }),
      );
    }
    return children;
  }

  private async richRuns(content: RichContent, options: RichRunOptions = {}): Promise<unknown[]> {
    const runs: unknown[] = [];
    for (const part of content) {
      runs.push(...(await this.partRuns(part, options)));
    }
    return runs.length > 0 ? runs : [runText('')];
  }

  private async partRuns(part: RichPart, options: RichRunOptions): Promise<unknown[]> {
    if (part.type === 'break') return [new docx.TextRun({ break: 1 })];
    if (part.type === 'text') {
      return [
        runText(part.text, {
          bold: options.bold || part.bold,
          italics: part.italic,
          subScript: part.subScript,
          superScript: part.superScript,
          color: options.color,
        }),
      ];
    }

    const asset = await this.images.load(part.url);
    if (!asset) {
      return [runText(`[图片加载失败: ${part.alt || part.url}]`, { color: '94A3B8', italics: true })];
    }
    return [
      new docx.ImageRun({
        data: asset.data,
        type: asset.type,
        transformation: { width: asset.width, height: asset.height },
        altText: {
          title: part.alt || '题目图片',
          description: part.alt || '题目图片',
          name: part.alt || '题目图片',
        },
      }),
    ];
  }
}
