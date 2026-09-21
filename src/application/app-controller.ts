import type { ExportOptions } from '../domain/export-options';
import type { HistoryEntry } from '../domain/history';
import type { ExtractionResult } from '../domain/question';
import type { ExportPreferences, PanelPosition, UserSettings } from '../domain/settings';
import { buildExtractionDiagnostics } from '../extractors/dom-diagnostics';
import type { ExportService } from '../exporters/export-service';
import { ClipboardService } from '../infrastructure/clipboard-service';
import { DownloadService } from '../infrastructure/download-service';
import type { FrameBridge } from '../infrastructure/frame-bridge';
import { HistoryRepository } from '../infrastructure/history-repository';
import { Logger } from '../infrastructure/logger';
import { SettingsRepository } from '../infrastructure/settings-repository';
import type { PanelView } from '../ui/panel-view';
import { debounce } from '../utils/async';
import { APP_VERSION } from '../app-config';
import { stableHash } from '../utils/hash';
import { isEditableTarget } from '../utils/dom';
import { matchesShortcut } from '../utils/shortcut';
import { humanizeError, sanitizeFilename } from '../utils/text';
import type { ChapterExtractionService } from './chapter-extraction-service';
import { incompleteChoiceCount, type ExtractionService } from './extraction-service';
import { QuestionTabGuard, TaskTabLocator } from './task-tab-locator';

// 页面存在题目卡时等待题目加载的上限（切换任务卡后答题页还要 AJAX 渲染）
const QUESTION_SETTLE_TIMEOUT_MS = 12_000;

// 页面没有题目卡时不再空等：只留出页面自身渲染完成的时间（与原 4.5s 等待的容错度一致）
const PAGE_IDLE_TIMEOUT_MS = 4_500;

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
    private readonly taskTabs = new TaskTabLocator(),
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
      onDiagnose: () => void this.copyDiagnostics(),
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
      this.frameBridge.requestRefresh();
      this.acceptResult(await this.resolveCurrentResult());
    } catch (error) {
      this.logger.warn('Extraction failed', error);
      this.view.setStatus(humanizeError(error), 'error');
    } finally {
      this.view.setBusy(null);
      this.refreshChapterAvailability();
    }
  }

  /**
   * 取当前页面的题目。
   *
   * 默认停留在「视频/文档」卡时，题目 iframe 的真实地址还留在 `_src` 上，页面里没有题目 DOM：
   * 此时采样等待并在需要时补切题目卡，避免固定等满整段超时。
   *
   * 已经取到题目时也不是立刻收工：答题页会分几批渲染，中途取到的结果可能只有题干、
   * 选项还没补上，这类「选择题缺选项」的结果要交给采样等待补齐。
   */
  private async resolveCurrentResult(): Promise<ExtractionResult> {
    const immediate = this.extractionService.extractAccessibleDocuments();
    if (immediate && immediate.questions.length > 0 && incompleteChoiceCount(immediate) === 0) {
      return immediate;
    }

    const guard = new QuestionTabGuard(this.taskTabs);
    // 页面根本没有题目卡时，结构不会自己变出题目，只留出页面自身渲染的时间
    const timeoutMs =
      this.taskTabs.questionTabIndex() === null ? PAGE_IDLE_TIMEOUT_MS : QUESTION_SETTLE_TIMEOUT_MS;
    const settled = await this.extractionService.settleQuestions({
      timeoutMs,
      repair: () => {
        if (guard.ensure()) {
          this.view.setStatus('已自动切换到题目所在任务卡，正在等待题目加载…', 'neutral');
        }
      },
    });

    if (!settled) throw new Error('当前页面未识别到题目，请确认题目已经加载完成');
    return settled;
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
      const warnMsg = failedImages > 0 ? `（${failedImages} 张图片加载失败）` : '';
      let message: string;
      if (artifacts.length > 1) {
        // 勾选「按章节拆分文件」时，各种格式都是一章一个文件
        message = `已下载 ${artifacts.length} 个章节文件`;
      } else if (options.format === 'word') {
        const withAnswers = options.bankImport || options.withAnswers;
        const withWrong = !options.bankImport && options.withWrong;
        message = options.bankImport
          ? '题库导入格式已下载'
          : 'Word 试卷' +
            (withAnswers ? '（含答案）' : '') +
            (withWrong ? '（含错题）' : '') +
            '已下载';
      } else {
        message = '文件已下载';
      }
      this.view.setStatus(message + warnMsg, failedImages > 0 ? 'warning' : 'success');
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

  /**
   * 复制提取诊断信息。
   *
   * 「提取到题干、提取不到选项」在离线快照上复现不出来（快照是渲染完成后的 DOM），
   * 所以把当前页面各层文档的题目/选项结构一次性复制出来，供定位成因：
   * 选择器没命中、选项在题目容器之外，还是选项节点本身就解析为空。
   */
  private async copyDiagnostics(): Promise<void> {
    try {
      await this.clipboardService.writeText(
        buildExtractionDiagnostics({ version: APP_VERSION }),
      );
      this.view.setStatus('诊断信息已复制到剪贴板，直接粘贴发给维护者即可', 'success');
    } catch (error) {
      this.view.setStatus(`复制诊断失败：${humanizeError(error)}`, 'error');
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
