import {
  DEFAULT_SETTINGS,
  type ExportPreferences,
  type PanelPosition,
  type Shortcut,
  type ThemeMode,
  type UserSettings,
} from '../domain/settings';
import { parseShortcut } from '../utils/shortcut';
import { SafeStorage } from './safe-storage';

const KEY = 'chaoxing-work-export:settings:v3';
const LEGACY_KEY = 'xxt_settings';
const LEGACY_EXPORT_KEYS = ['xxt_export_config', 'xxt_export_options'] as const;
// 「打开窗口自动提取」默认值升级的标记（只升级一次，之后尊重用户的显式选择）
const AUTO_EXTRACT_DEFAULT_KEY = 'chaoxing-work-export:auto-extract-default:v1';

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function booleanValue(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

function shortcutValue(value: unknown, fallback: Shortcut): Shortcut {
  if (typeof value === 'string') return parseShortcut(value, fallback);
  if (!isRecord(value)) return { ...fallback };
  return {
    ctrl: booleanValue(value.ctrl, fallback.ctrl),
    shift: booleanValue(value.shift, fallback.shift),
    alt: booleanValue(value.alt, fallback.alt),
    key: typeof value.key === 'string' && value.key ? value.key.toLowerCase() : fallback.key,
  };
}

function panelPositionValue(value: unknown): PanelPosition | null {
  if (!isRecord(value)) return null;
  return typeof value.left === 'number' && typeof value.top === 'number'
    ? { left: value.left, top: value.top }
    : null;
}

function themeValue(value: unknown): ThemeMode {
  if (value === 'system') return 'auto';
  return value === 'light' || value === 'dark' || value === 'auto' ? value : 'auto';
}

function firstBoolean(
  source: Record<string, unknown>,
  keys: readonly string[],
  fallback: boolean,
): boolean {
  for (const key of keys) {
    if (typeof source[key] === 'boolean') return source[key];
  }
  return fallback;
}

function exportPreferencesValue(value: unknown): ExportPreferences {
  const source = isRecord(value) ? value : {};
  const rawFormat = source.format ?? source.exportFormat;
  const format = rawFormat === 'txt' || rawFormat === 'md' ? rawFormat : 'word';
  return {
    format,
    withAnswers: firstBoolean(
      source,
      ['withAnswers', 'includeAnswers', 'appendAnswers'],
      DEFAULT_SETTINGS.exportPreferences.withAnswers,
    ),
    withWrong: firstBoolean(
      source,
      ['withWrong', 'includeWrong', 'appendWrong'],
      DEFAULT_SETTINGS.exportPreferences.withWrong,
    ),
    shuffle: firstBoolean(
      source,
      ['shuffle', 'randomOrder'],
      DEFAULT_SETTINGS.exportPreferences.shuffle,
    ),
    bankImport: firstBoolean(
      source,
      ['bankImport', 'questionBankFormat', 'bankMode'],
      DEFAULT_SETTINGS.exportPreferences.bankImport,
    ),
    splitByChapter: firstBoolean(
      source,
      ['splitByChapter', 'splitChapters'],
      DEFAULT_SETTINGS.exportPreferences.splitByChapter,
    ),
  };
}

function settingsValue(value: unknown): UserSettings {
  const source = isRecord(value) ? value : {};
  const legacyExport = source.exportConfig ?? source.exportOptions;
  return {
    theme: themeValue(source.theme),
    shortcut: shortcutValue(source.shortcut ?? source.extractShortcut, DEFAULT_SETTINGS.shortcut),
    hideShortcut: shortcutValue(
      source.hideShortcut ?? source.toggleShortcut,
      DEFAULT_SETTINGS.hideShortcut,
    ),
    exportPreferences: exportPreferencesValue(source.exportPreferences ?? legacyExport ?? source),
    enableDrag: booleanValue(source.enableDrag ?? source.draggable, DEFAULT_SETTINGS.enableDrag),
    rememberPanelPosition: booleanValue(
      source.rememberPanelPosition ?? source.rememberPosition,
      DEFAULT_SETTINGS.rememberPanelPosition,
    ),
    panelPosition: panelPositionValue(source.panelPosition ?? source.panelPos),
    autoExtractOnLoad: booleanValue(source.autoExtractOnLoad, DEFAULT_SETTINGS.autoExtractOnLoad),
  };
}

export class SettingsRepository {
  constructor(private readonly storage = new SafeStorage()) {}

  load(): UserSettings {
    const current = this.storage.read(KEY);
    if (current !== null) return this.upgradeAutoExtractDefault(settingsValue(current));

    const legacySettings = this.storage.read(LEGACY_KEY);
    const legacyExport = LEGACY_EXPORT_KEYS.map((key) => this.storage.read(key)).find(
      (value) => value !== null,
    );
    if (legacySettings !== null || legacyExport !== undefined) {
      const merged = isRecord(legacySettings) ? { ...legacySettings } : {};
      if (legacyExport !== undefined) merged.exportPreferences = legacyExport;
      const migrated = settingsValue(merged);
      this.save(migrated);
      return migrated;
    }
    return settingsValue(DEFAULT_SETTINGS);
  }

  save(settings: UserSettings): void {
    this.storage.write(KEY, settings);
  }

  /**
   * 一次性默认值升级：「打开窗口自动提取」的默认值由关闭改为打开。
   *
   * 只改 DEFAULT_SETTINGS 对存量配置无效 —— 该字段早就被显式写入过（值为旧的默认值 false），
   * 会一直盖住新的默认值。这里在首次加载时补成新默认值并记下标记；标记置位后
   * 用户在设置里主动关掉会被正常保存，不会再次被改回。
   */
  private upgradeAutoExtractDefault(settings: UserSettings): UserSettings {
    if (this.storage.read(AUTO_EXTRACT_DEFAULT_KEY) !== null) return settings;
    this.storage.write(AUTO_EXTRACT_DEFAULT_KEY, true);
    if (settings.autoExtractOnLoad) return settings;
    const upgraded: UserSettings = { ...settings, autoExtractOnLoad: true };
    this.save(upgraded);
    return upgraded;
  }
}
