import type { ExportOptions } from './export-options';
import type { ExtractionResult } from './question';

export interface HistoryEntry {
  readonly id: string;
  readonly createdAt: string;
  readonly title: string;
  readonly result: ExtractionResult;
  readonly options: ExportOptions;
  readonly textSnapshot?: string;
}
