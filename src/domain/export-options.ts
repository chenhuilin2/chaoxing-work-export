export type ExportFormat = 'word' | 'txt' | 'md';

export interface ExportOptions {
  readonly format: ExportFormat;
  readonly filename: string;
  readonly withAnswers: boolean;
  readonly withWrong: boolean;
  readonly shuffle: boolean;
  readonly bankImport: boolean;
  /** 按「章」拆文件：一章一个文件，文件里装该章的所有小节 */
  readonly splitByChapter: boolean;
  /**
   * 按「小节」拆文件：一个小节一个文件。
   *
   * 与 `splitByChapter` 互斥（面板上勾一个会取消另一个）；两者同时为真时以小节为准。
   * 存在的理由：目录树只有一个章、下面挂十几个小节时（如「1 章节测试」下 1.1~1.11），
   * 按章拆只会产出 1 个文件，用户真正想要的是按小节拆开。
   */
  readonly splitBySection: boolean;
}

export interface ExportArtifact {
  readonly blob: Blob;
  readonly filename: string;
  readonly mimeType: string;
  readonly failedImages?: number;
  /** 逐章文件打包成压缩包时，包里装了几个章节文件（用于下载提示） */
  readonly archiveEntries?: number;
}
