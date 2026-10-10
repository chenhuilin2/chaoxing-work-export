export type ExportFormat = 'word' | 'txt' | 'md';

export interface ExportOptions {
  readonly format: ExportFormat;
  readonly filename: string;
  readonly withAnswers: boolean;
  readonly withWrong: boolean;
  readonly shuffle: boolean;
  readonly bankImport: boolean;
  readonly splitByChapter: boolean;
}

export interface ExportArtifact {
  readonly blob: Blob;
  readonly filename: string;
  readonly mimeType: string;
  readonly failedImages?: number;
  /** 逐章文件打包成压缩包时，包里装了几个章节文件（用于下载提示） */
  readonly archiveEntries?: number;
}
