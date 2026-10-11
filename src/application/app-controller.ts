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
import { ChapterChangeWatcher } from './chapter-change-watcher';
import { incompleteChoiceCount, type ExtractionService } from './extraction-service';
import { QuestionTabGuard, TaskTabLocator } from './task-tab-locator';

// 页面存在题目卡时等待题目加载的上限（切换任务卡后答题页还要 AJAX 渲染）
const QUESTION_SETTLE_TIMEOUT_MS = 12_000;

// 页面没有题目卡时不再空等：只留出页面自身渲染完成的时间（与原 4.5s 等待的容错度一致）
const PAGE_IDLE_TIMEOUT_MS = 4_500;

// 章节切换的轮询间隔：切章是 AJAX 局部刷新（URL 不变），只能靠轮询发现；
// 每次轮询只做一次 querySelector，开销可忽略
const CHAPTER_POLL_INTERVAL_MS = 500;

// 确认切章后等这么久再看任务卡栏：新章节的卡片页（含任务卡栏）是异步重建的
const CHAPTER_TAB_SETTLE_MS = 600;

/** 一次提取的调用选项 */
interface ExtractOptions {
  /** 上一份结果的指纹：传入时要求页面内容确实变过，避免读到上一章残留的题目 */
  readonly previousFingerprint?: string;
  /** 后台自动触发：失败时只给一句中性提示，不按错误弹给用户 */
  readonly quiet?: boolean;
}

export class AppController {
  private currentResult: ExtractionResult | null = null;
  private settings: UserSettings;
  private history: HistoryEntry[];
  private observer: MutationObserver | null = null;
  private chapterWatcher: ChapterChangeWatcher | null = null;
  private chapterPollTimer: number | null = null;
  /** >0 表示正处于批量章节提取中：期间的章节切换是我们自己点的，不能触发自动提取 */
  private chapterWatchSuspend = 0;
  /** 提取进行中；期间到来的自动提取请求合并成一次待办，结束后补跑 */
  private extracting = false;
  private pendingAutoExtract = false;
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

    // 自动提取：进入页面时提取一次
    if (this.settings.autoExtractOnLoad) {
      window.setTimeout(() => void this.extractCurrent(), 800);
    }
    this.startChapterWatch();
  }

  /**
   * 开始监听章节切换。
   *
   * 章节练习页切章是 AJAX 局部刷新（`getTeacherAjax()` 重建 `#mainid`），URL 不变、
   * 脚本不重跑，所以「进入页面提取一次」之后就没有触发点了 —— 这里补上：切到新章节后
   * 自动提取一次当前章节的题目，用户不必再手点按钮。
   */
  private startChapterWatch(): void {
    this.chapterWatcher = new ChapterChangeWatcher(() => this.chapterService.activeChapterKey());
    // 以「进入页面时的章节」为基线，避免把首次渲染出来的目录树当成切章
    this.chapterWatcher.reset();
    this.chapterPollTimer = window.setInterval(
      () => this.detectChapterChange(),
      CHAPTER_POLL_INTERVAL_MS,
    );
  }

  /**
   * 轮询到章节已被切换时自动提取当前章节的题目。
   *
   * 新章节没有题目卡时直接跳过：纯视频/文档章节的题目永远不会出现，没必要空等十几秒
   * 再报一次「未识别到题目」。任务卡栏不存在时（非知识卡片页）不跳过，照常尝试。
   */
  private detectChapterChange(): void {
    if (!this.chapterWatcher?.poll()) return;
    if (!this.settings.autoExtractOnLoad) return;
    if (this.chapterWatchSuspend > 0) return;

    window.setTimeout(() => {
      if (this.chapterWatchSuspend > 0) return;
      const tabs = this.taskTabs;
      if (tabs.present() && tabs.questionTabIndex() === null) {
        this.logger.debug('Chapter has no question tab, skip auto extraction');
        return;
      }
      void this.extractCurrent({
        previousFingerprint: this.currentFingerprint(),
        quiet: true,
      });
    }, CHAPTER_TAB_SETTLE_MS);
  }

  /** 上一份结果的指纹：切章后用它确认「读到的确实不是上一章残留」 */
  private currentFingerprint(): string | undefined {
    return this.currentResult ? this.extractionService.fingerprint(this.currentResult) : undefined;
  }

  stop(): void {
    window.removeEventListener('keydown', this.onShortcut, true);
    this.observer?.disconnect();
    this.observer = null;
    if (this.chapterPollTimer !== null) {
      window.clearInterval(this.chapterPollTimer);
      this.chapterPollTimer = null;
    }
    this.chapterWatcher = null;
    this.frameBridge.stop();
    this.view.destroy();
  }

  private async extractCurrent(options: ExtractOptions = {}): Promise<void> {
    // 提取期间又被触发（用户连点章节、或快捷键）：合并成一次待办，等这次结束后补跑，
    // 避免两次提取互相抢任务卡、抢 DOM
    if (this.extracting) {
      this.pendingAutoExtract = true;
      return;
    }
    this.extracting = true;
    this.view.setBusy('extract', '正在识别当前页面和已加载的 iframe…');
    try {
      this.frameBridge.requestRefresh();
      this.acceptResult(await this.resolveCurrentResult(options));
    } catch (error) {
      this.logger.warn('Extraction failed', error);
      // 后台自动触发时用户并没有发出请求，「未识别到题目」是正常情况而非错误
      this.view.setStatus(
        options.quiet ? '当前章节未识别到题目' : humanizeError(error),
        options.quiet ? 'warning' : 'error',
      );
    } finally {
      this.view.setBusy(null);
      this.refreshChapterAvailability();
      this.extracting = false;
      if (this.pendingAutoExtract) {
        this.pendingAutoExtract = false;
        void this.extractCurrent({
          previousFingerprint: this.currentFingerprint(),
          quiet: true,
        });
      }
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
  private async resolveCurrentResult(options: ExtractOptions = {}): Promise<ExtractionResult> {
    const previousFingerprint = options.previousFingerprint;
    const immediate = this.extractionService.extractAccessibleDocuments();
    // 指纹与上一份结果一致，说明新章节的内容还没进 DOM，此刻读到的仍是上一章残留的题目
    const stale =
      previousFingerprint !== undefined &&
      immediate !== null &&
      this.extractionService.fingerprint(immediate) === previousFingerprint;

    if (
      !stale &&
      immediate &&
      immediate.questions.length > 0 &&
      incompleteChoiceCount(immediate) === 0
    ) {
      return immediate;
    }

    const guard = new QuestionTabGuard(this.taskTabs);
    // 页面根本没有题目卡时，结构不会自己变出题目，只留出页面自身渲染的时间
    const timeoutMs =
      this.taskTabs.questionTabIndex() === null ? PAGE_IDLE_TIMEOUT_MS : QUESTION_SETTLE_TIMEOUT_MS;
    const settled = await this.extractionService.settleQuestions({
      timeoutMs,
      previousFingerprint,
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
    // 批量提取会自己逐章切换章节，期间的切章不能被当成「用户切章」而触发自动提取，
    // 否则每切一章都会再抢一次任务卡，与批量流程互相打断
    this.chapterWatchSuspend += 1;
    this.pendingAutoExtract = false;
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
      this.chapterWatchSuspend -= 1;
      // 批量结束会把章节切回原处，重新对准基线，避免这次遍历被记成一次「切章」
      this.chapterWatcher?.reset();
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
      // 逐章文件太多时导出的是压缩包：报「打包了几个章节文件」比「下载了 1 个文件」有用
      const archived = artifacts[0]?.archiveEntries ?? 0;
      let message: string;
      if (archived > 0) {
        message = `章节较多，已把 ${archived} 个章节文件打包成 1 个 zip`;
      } else if (artifacts.length > 1) {
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
