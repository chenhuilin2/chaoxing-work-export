import type { ExportFormat } from './export-options';

export type ThemeMode = 'auto' | 'light' | 'dark';

export interface Shortcut {
  readonly ctrl: boolean;
  readonly shift: boolean;
  readonly alt: boolean;
  readonly key: string;
}

export interface PanelPosition {
  readonly left: number;
  readonly top: number;
}

export interface ExportPreferences {
  readonly format: ExportFormat;
  readonly withAnswers: boolean;
  readonly withWrong: boolean;
  readonly includeAnalysis: boolean;
  readonly shuffle: boolean;
  readonly bankImport: boolean;
  readonly splitByChapter: boolean;
}

export interface UserSettings {
  readonly theme: ThemeMode;
  readonly shortcut: Shortcut;
  readonly hideShortcut: Shortcut;
  readonly exportPreferences: ExportPreferences;
  readonly enableDrag: boolean;
  readonly rememberPanelPosition: boolean;
  readonly panelPosition: PanelPosition | null;
}

export const DEFAULT_SETTINGS: UserSettings = {
  theme: 'auto',
  shortcut: { ctrl: true, shift: true, alt: false, key: 'e' },
  hideShortcut: { ctrl: true, shift: true, alt: false, key: 'h' },
  exportPreferences: {
    format: 'word',
    withAnswers: false,
    withWrong: false,
    includeAnalysis: true,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
  },
  enableDrag: false,
  rememberPanelPosition: true,
  panelPosition: null,
};
