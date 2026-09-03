export type ExportFormat = 'word' | 'txt' | 'md';

export interface ExportOptions {
  readonly format: ExportFormat;
  readonly filename: string;
  readonly withAnswers: boolean;
  readonly withWrong: boolean;
  readonly includeAnalysis: boolean;
  readonly shuffle: boolean;
  readonly bankImport: boolean;
  readonly splitByChapter: boolean;
}

export interface ExportArtifact {
  readonly blob: Blob;
  readonly filename: string;
  readonly mimeType: string;
  readonly failedImages?: number;
}
