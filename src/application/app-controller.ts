import type { ExportOptions } from '../domain/export-options';
import type { HistoryEntry } from '../domain/history';
import type { ExtractionResult } from '../domain/question';
import type { ExportPreferences, PanelPosition, UserSettings } from '../domain/settings';
import type { ExportService } from '../exporters/export-service';
import { ClipboardService } from '../infrastructure/clipboard-service';
import { DownloadService } from '../infrastructure/download-service';
import type { FrameBridge } from '../infrastructure/frame-bridge';
import { HistoryRepository } from '../infrastructure/history-repository';
import { Logger } from '../infrastructure/logger';
import { SettingsRepository } from '../infrastructure/settings-repository';
import type { PanelView } from '../ui/panel-view';
import { debounce } from '../utils/async';
import { stableHash } from '../utils/hash';
import { isEditableTarget } from '../utils/dom';
import { matchesShortcut } from '../utils/shortcut';
import { humanizeError, sanitizeFilename } from '../utils/text';
import type { ChapterExtractionService } from './chapter-extraction-service';
import type { ExtractionService } from './extraction-service';

export class AppController {
  private currentResult: ExtractionResult | null = null;
  private settings: UserSettings;
  private history: HistoryEntry[];
  private observer: MutationObserver | null = null;
  private readonly logger = new Logger();
  private readonly refreshChaptersDebounced = debounce(() => this.refreshChapterAvailability(), 700);
  private readonly onShortcut = (event: KeyboardEvent): void => {
    if (isEditableTarget(event.target)) return;
    if (matchesShortcut(event, this.settings.hideShortcut)) {
      event.preventDefault();
      const hidden = this.view.toggleGlobalVisibility();
      if (!hidden) this.view.setStatus('浮窗已恢复', 'success');
      return;
    }
    if (matchesShortcut(event, this.settings.shortcut)) {
      event.preventDefault();
      this.view.open();
      void this.extractCurrent();
    }
  };

  constructor(
    private readonly view: PanelView,
    private readonly extractionService: ExtractionService,
    private readonly chapterService: ChapterExtractionService,
    private readonly exportService: ExportService,
    private readonly frameBridge: FrameBridge,
    private readonly settingsRepository = new SettingsRepository(),
    private readonly historyRepository = new HistoryRepository(),
    private readonly downloadService = new DownloadService(),
    private readonly clipboardService = new ClipboardService(),
  ) {
    this.settings = this.settingsRepository.load();
    this.history = this.historyRepository.load();
  }

  start(): void {
    this.frameBridge.start();
    this.view.setCallbacks({
      onExtract: () => void this.extractCurrent(),
      onOpenChapters: () => this.openChapterDialog(),
      onExtractChapters: (indexes) => void this.extractChapters(indexes),
      onDownload: () => void this.downloadCurrent(),
      onCopy: () => void this.copyCurrent(),
      onPreferencesChange: (preferences) => this.savePreferences(preferences),
      onSettingsSave: (settings) => this.saveSettings(settings),
      onHistoryRestore: (id) => this.restoreHistory(id),
      onHistoryDownload: (id) => void this.downloadHistory(id),
      onHistoryDelete: (id) => this.deleteHistory(id),
      onPanelPositionChange: (position) => this.savePanelPosition(position),
    });
    this.view.applySettings(this.settings);
    this.view.renderHistory(this.history);
    this.refreshChapterAvailability();
    window.addEventListener('keydown', this.onShortcut, true);

    if (document.body) {
      this.observer = new MutationObserver(() => this.refreshChaptersDebounced());
      this.observer.observe(document.body, { childList: true, subtree: true });
    }

    // 自动提取：进入页面时仅提取一次，后续页面变化不再自动提取
    if (this.settings.autoExtractOnLoad) {
      window.setTimeout(() => void this.extractCurrent(), 800);
    }
  }

  stop(): void {
    window.removeEventListener('keydown', this.onShortcut, true);
    this.observer?.disconnect();
    this.observer = null;
    this.frameBridge.stop();
    this.view.destroy();
  }

  private async extractCurrent(): Promise<void> {
    this.view.setBusy('extract', '正在识别当前页面和已加载的 iframe…');
    try {
      const requestedAt = Date.now();
      this.frameBridge.requestRefresh();
      let result = this.extractionService.extractAccessibleDocuments();
      if (!result) {
        result = await this.extractionService.waitForChangedResult({
          afterTimestamp: requestedAt,
          timeoutMs: 4_500,
          intervalMs: 300,
        });
      }
      if (!result || result.questions.length === 0) {
        throw new Error('当前页面未识别到题目，请确认题目已经加载完成');
      }
      this.acceptResult(result);
    } catch (error) {
      this.logger.warn('Extraction failed', error);
      this.view.setStatus(humanizeError(error), 'error');
    } finally {
      this.view.setBusy(null);
      this.refreshChapterAvailability();
    }
  }

  private openChapterDialog(): void {
    const chapters = this.chapterService.listChapters();
    this.view.setChapterCount(chapters.length);
    if (chapters.length < 2) {
      this.view.setStatus('当前页面未识别到可批量遍历的章节列表', 'warning');
      return;
    }
    this.view.showChapterDialog(chapters);
  }

  private async extractChapters(indexes: readonly number[]): Promise<void> {
    if (indexes.length === 0) {
      this.view.setStatus('请至少选择一个章节', 'warning');
      return;
    }
    this.view.setBusy('chapters', `准备提取 ${indexes.length} 个章节…`);
    try {
      const result = await this.chapterService.extractSelected(indexes, (progress) => {
        this.view.updateChapterProgress(progress);
      });
      this.acceptResult(result);
      this.view.closeChapterDialog();
    } catch (error) {
      this.logger.warn('Chapter extraction failed', error);
      this.view.setStatus(humanizeError(error), 'error');
    } finally {
      this.view.setBusy(null);
      this.refreshChapterAvailability();
    }
  }

  private acceptResult(result: ExtractionResult): void {
    this.currentResult = result;
    this.view.renderResult(result);
    const options = this.view.readExportOptions();
    this.view.applyExportOptions({ ...options, filename: sanitizeFilename(result.title) });
  }

  private async downloadCurrent(): Promise<void> {
    if (!this.currentResult) {
      this.view.setStatus('请先提取题目', 'warning');
      return;
    }
    const options = this.view.readExportOptions();
    this.view.setBusy('download', options.format === 'word' ? '正在生成 Word 文档并处理图片…' : '正在生成文件…');
    try {
      const artifacts = await this.exportService.createArtifacts(this.currentResult, options);
      await this.downloadService.downloadMany(artifacts);
      const failedImages = artifacts.reduce((total, artifact) => total + (artifact.failedImages ?? 0), 0);
      this.addHistory(this.currentResult, options);
      // 下载状态提示与主脚本保持一致
      let message: string;
      if (options.format === 'word') {
        const warnMsg = failedImages > 0 ? `（${failedImages} 张图片加载失败）` : '';
        const withAnswers = options.bankImport || options.withAnswers;
        const withWrong = !options.bankImport && options.withWrong;
        message =
          (options.bankImport
            ? '题库导入格式已下载'
            : 'Word 试卷' + (withAnswers ? '（含答案）' : '') + (withWrong ? '（含错题）' : '') + '已下载') +
          warnMsg;
      } else if (artifacts.length > 1) {
        message = `已下载 ${artifacts.length} 个章节文件`;
      } else {
        message = '文件已下载';
      }
      this.view.setStatus(message, failedImages > 0 ? 'warning' : 'success');
    } catch (error) {
      this.logger.error('Export failed', error);
      const prefix = options.format === 'word' ? 'Word 导出失败: ' : '导出失败：';
      this.view.setStatus(`${prefix}${humanizeError(error)}`, 'error');
    } finally {
      this.view.setBusy(null);
    }
  }

  private async copyCurrent(): Promise<void> {
    if (!this.currentResult) {
      this.view.setStatus('请先提取题目', 'warning');
      return;
    }
    const options = this.view.readExportOptions();
    // 原版行为：Word 格式不支持复制
    if (options.format === 'word') {
      this.view.setStatus('Word 格式不支持复制，请使用下载', 'warning');
      return;
    }
    this.view.setBusy('copy', '正在生成可复制文本…');
    try {
      const text = this.exportService.previewText(this.currentResult, options);
      await this.clipboardService.writeText(text);
      this.view.setStatus('已复制到剪贴板', 'success');
    } catch (error) {
      this.view.setStatus(`复制失败：${humanizeError(error)}`, 'error');
    } finally {
      this.view.setBusy(null);
    }
  }

  private addHistory(result: ExtractionResult, options: ExportOptions): void {
    const createdAt = new Date().toISOString();
    const entry: HistoryEntry = {
      id: `history-${stableHash(`${createdAt}|${result.title}|${result.questions.length}`)}`,
      createdAt,
      title: options.filename || result.title,
      result,
      options,
    };
    this.history = this.historyRepository.add(entry);
    this.view.renderHistory(this.history);
  }

  private restoreHistory(id: string): void {
    const entry = this.history.find((candidate) => candidate.id === id);
    if (!entry) {
      this.view.setStatus('历史记录不存在或已被清理', 'warning');
      return;
    }
    this.currentResult = entry.result;
    this.view.renderResult(entry.result);
    this.view.applyExportOptions(entry.options);
    this.view.closeHistory();
    this.view.open();
    this.view.setStatus(`已恢复历史记录：${entry.title}`, 'success');
  }

  private async downloadHistory(id: string): Promise<void> {
    const entry = this.history.find((candidate) => candidate.id === id);
    if (!entry) return;
    this.view.closeHistory();
    this.view.setBusy('download', '正在重新生成历史文件…');
    try {
      const artifacts = await this.exportService.createArtifacts(entry.result, entry.options);
      await this.downloadService.downloadMany(artifacts);
      this.view.setStatus(`历史文件已重新下载`, 'success');
    } catch (error) {
      this.view.setStatus(`重新下载失败：${humanizeError(error)}`, 'error');
    } finally {
      this.view.setBusy(null);
    }
  }

  private deleteHistory(id: string): void {
    this.history = this.historyRepository.remove(id);
    this.view.renderHistory(this.history);
    this.view.setStatus('历史记录已删除', 'success');
  }

  private savePreferences(preferences: ExportPreferences): void {
    this.settings = { ...this.settings, exportPreferences: preferences };
    this.settingsRepository.save(this.settings);
  }

  private saveSettings(settings: UserSettings): void {
    this.settings = settings;
    this.settingsRepository.save(settings);
    this.view.applySettings(settings);
  }

  private savePanelPosition(position: PanelPosition): void {
    if (!this.settings.rememberPanelPosition) return;
    this.settings = { ...this.settings, panelPosition: position };
    this.settingsRepository.save(this.settings);
    this.view.applySettings(this.settings);
  }

  private refreshChapterAvailability(): void {
    try {
      this.view.setChapterCount(this.chapterService.listChapters().length);
    } catch (error) {
      this.logger.debug('Chapter discovery skipped', error);
      this.view.setChapterCount(0);
    }
  }
}
