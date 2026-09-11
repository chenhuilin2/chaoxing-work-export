import { APP_NAME, UI_HOST_ID } from '../app-config';
import type { ExportOptions } from '../domain/export-options';
import type { HistoryEntry } from '../domain/history';
import type { ExtractionResult } from '../domain/question';
import type { ExportPreferences, PanelPosition, UserSettings } from '../domain/settings';
import type { ChapterDescriptor } from '../application/chapter-locator';
import type { ChapterProgress } from '../application/chapter-extraction-service';
import { queryRequired } from '../utils/assert';
import { formatShortcut, parseShortcut } from '../utils/shortcut';
import { sanitizeFilename } from '../utils/text';
import { PANEL_STYLES } from './styles';

export type StatusKind = 'neutral' | 'success' | 'warning' | 'error';
export type BusyAction = 'extract' | 'chapters' | 'download' | 'copy' | null;

export interface PanelCallbacks {
  readonly onExtract: () => void;
  readonly onOpenChapters: () => void;
  readonly onExtractChapters: (indexes: readonly number[]) => void;
  readonly onDownload: () => void;
  readonly onCopy: () => void;
  readonly onPreferencesChange: (preferences: ExportPreferences) => void;
  readonly onSettingsSave: (settings: UserSettings) => void;
  readonly onHistoryRestore: (id: string) => void;
  readonly onHistoryDownload: (id: string) => void;
  readonly onHistoryDelete: (id: string) => void;
  readonly onPanelPositionChange: (position: PanelPosition) => void;
}

const NOOP_CALLBACKS: PanelCallbacks = {
  onExtract: () => undefined,
  onOpenChapters: () => undefined,
  onExtractChapters: () => undefined,
  onDownload: () => undefined,
  onCopy: () => undefined,
  onPreferencesChange: () => undefined,
  onSettingsSave: () => undefined,
  onHistoryRestore: () => undefined,
  onHistoryDownload: () => undefined,
  onHistoryDelete: () => undefined,
  onPanelPositionChange: () => undefined,
};

const ICONS = {
  settings:
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.62l-1.92-3.32a.49.49 0 0 0-.59-.22l-2.39.96a7.2 7.2 0 0 0-1.62-.94L14.4 2.8a.48.48 0 0 0-.48-.4h-3.84a.48.48 0 0 0-.48.4l-.36 2.54c-.59.24-1.13.56-1.62.94l-2.39-.96a.48.48 0 0 0-.59.22L2.72 8.86a.48.48 0 0 0 .12.62l2.03 1.58c-.05.3-.08.63-.08.94s.03.64.08.94l-2.03 1.58a.49.49 0 0 0-.12.62l1.92 3.32c.12.22.37.3.59.22l2.39-.96c.49.38 1.03.7 1.62.94l.36 2.54c.04.23.24.4.48.4h3.84c.24 0 .44-.17.48-.4l.36-2.54a7.2 7.2 0 0 0 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32a.49.49 0 0 0-.12-.62l-2.02-1.58ZM12 15.5A3.5 3.5 0 1 1 12 8a3.5 3.5 0 0 1 0 7.5Z"/></svg>',
  history:
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M13 3a9 9 0 1 0 8.48 6H19.3A7 7 0 1 1 13 5c1.93 0 3.68.78 4.95 2.05L15 10h7V3l-2.63 2.63A8.96 8.96 0 0 0 13 3Zm-1 4v6l5 3 .75-1.23-4.25-2.52V7H12Z"/></svg>',
  close:
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m18.3 5.7-1-1L12 10l-5.3-5.3-1 1L11 11l-5.3 5.3 1 1L12 12l5.3 5.3 1-1L13 11l5.3-5.3Z"/></svg>',
};

// 概览统计区展示的题型，与卡片顺序一致
const TYPE_STAT_KEYS = [
  'single-choice',
  'multiple-choice',
  'fill-blank',
  'true-false',
  'short-answer',
] as const;

function template(): string {
  return `
    <style>${PANEL_STYLES}</style>
    <div class="cwe-shell">
      <button class="cwe-launcher" type="button" aria-label="打开题目导出面板">提取题目</button>
      <section class="cwe-panel" data-open="false" role="dialog" aria-label="${APP_NAME}">
        <header class="cwe-header">
          <div class="cwe-header-main">
            <h2 class="cwe-title">学习通题目导出</h2>
          </div>
          <div class="cwe-header-actions">
            <button class="cwe-icon-button" data-action="history" type="button" title="下载历史">${ICONS.history}</button>
            <button class="cwe-icon-button" data-action="settings" type="button" title="设置">${ICONS.settings}</button>
            <button class="cwe-icon-button" data-action="close" type="button" title="关闭">${ICONS.close}</button>
          </div>
        </header>
        <div class="cwe-scroll">
          <div class="cwe-status" data-kind="neutral">进入作业、考试或章节练习页面后开始提取</div>
          <div class="cwe-extract-grid">
            <button class="cwe-button cwe-button-primary" data-action="extract" type="button">提取当前页面</button>
            <button class="cwe-button cwe-button-soft" data-action="chapters" type="button" disabled>提取多个章节</button>
          </div>

          <div class="cwe-stats" aria-label="提取统计">
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="single-choice">0</span><span class="cwe-stat-label">单选</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="multiple-choice">0</span><span class="cwe-stat-label">多选</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="fill-blank">0</span><span class="cwe-stat-label">填空</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="true-false">0</span><span class="cwe-stat-label">判断</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="short-answer">0</span><span class="cwe-stat-label">简答</span></div>
          </div>

          <section class="cwe-section">
            <h3 class="cwe-section-title">导出文件</h3>
            <label class="cwe-field">
              <span class="cwe-label">文件名</span>
              <input class="cwe-input" data-field="filename" type="text" maxlength="120" value="学习通题目" />
            </label>
            <div class="cwe-formats" style="margin-top: 9px">
              <label class="cwe-format"><input type="radio" name="cwe-format" value="word" checked><span>Word 试卷</span></label>
              <label class="cwe-format"><input type="radio" name="cwe-format" value="txt"><span>TXT 文本</span></label>
              <label class="cwe-format"><input type="radio" name="cwe-format" value="md"><span>Markdown</span></label>
            </div>
            <div class="cwe-options">
              <label class="cwe-check"><input data-option="withAnswers" type="checkbox">附加答案</label>
              <label class="cwe-check"><input data-option="includeAnalysis" type="checkbox">附加解析</label>
              <label class="cwe-check"><input data-option="withWrong" type="checkbox">附加错题</label>
              <label class="cwe-check"><input data-option="shuffle" type="checkbox">题型内乱序</label>
              <label class="cwe-check"><input data-option="bankImport" type="checkbox">题库导入</label>
              <label class="cwe-check"><input data-option="splitByChapter" type="checkbox" disabled>按章节拆分文件</label>
            </div>
          </section>

          <div class="cwe-actions">
            <button class="cwe-button cwe-button-primary" data-action="download" type="button" disabled>下载文件</button>
            <button class="cwe-button" data-action="copy" type="button" disabled>复制文本</button>
          </div>
          <div class="cwe-footer">仅处理当前浏览器中已加载的题目；请遵守课程与平台规则。</div>
        </div>
      </section>

      <div class="cwe-modal-backdrop" data-modal="settings" data-open="false">
        <section class="cwe-modal" role="dialog" aria-modal="true" aria-label="脚本设置">
          <header class="cwe-modal-header">
            <h3 class="cwe-modal-title">脚本设置</h3>
            <button class="cwe-icon-button" data-modal-close="settings" type="button">${ICONS.close}</button>
          </header>
          <div class="cwe-modal-body">
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">界面主题</span>
              <select class="cwe-select" data-setting="theme">
                <option value="auto">跟随系统</option><option value="light">浅色</option><option value="dark">深色</option>
              </select>
            </label>
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">提取快捷键</span>
              <input class="cwe-input" data-setting="shortcut" placeholder="Ctrl+Shift+E" />
            </label>
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">隐藏快捷键</span>
              <input class="cwe-input" data-setting="hideShortcut" placeholder="Ctrl+Shift+H" />
            </label>
            <label class="cwe-switch-row"><span>允许拖动面板</span><input data-setting="enableDrag" type="checkbox"></label>
            <label class="cwe-switch-row"><span>记忆面板位置</span><input data-setting="rememberPanelPosition" type="checkbox"></label>
            <label class="cwe-switch-row"><span>打开窗口自动提取</span><input data-setting="autoExtractOnLoad" type="checkbox"></label>
          </div>
          <footer class="cwe-modal-footer">
            <button class="cwe-button" data-modal-close="settings" type="button">取消</button>
            <button class="cwe-button cwe-button-primary" data-action="save-settings" type="button">保存设置</button>
          </footer>
        </section>
      </div>

      <div class="cwe-modal-backdrop" data-modal="history" data-open="false">
        <section class="cwe-history-modal" role="dialog" aria-modal="true" aria-label="下载历史">
          <header class="cwe-history-modal-header">
            <h3 class="cwe-history-modal-title">历史记录</h3>
            <button class="cwe-history-modal-close" data-modal-close="history" type="button" aria-label="关闭">&times;</button>
          </header>
          <div class="cwe-history-list" data-history-list></div>
        </section>
      </div>

      <div class="cwe-modal-backdrop" data-modal="chapters" data-open="false">
        <section class="cwe-modal" role="dialog" aria-modal="true" aria-label="选择章节">
          <header class="cwe-modal-header">
            <h3 class="cwe-modal-title">选择要提取的章节</h3>
            <button class="cwe-icon-button" data-modal-close="chapters" type="button">${ICONS.close}</button>
          </header>
          <div class="cwe-modal-body">
            <div class="cwe-chapter-toolbar">
              <button class="cwe-mini-button" data-action="chapter-all" type="button">全选</button>
              <button class="cwe-mini-button" data-action="chapter-none" type="button">取消全选</button>
            </div>
            <div class="cwe-chapter-list" data-chapter-list></div>
            <div class="cwe-chapter-progress" data-chapter-progress></div>
          </div>
          <footer class="cwe-modal-footer">
            <button class="cwe-button" data-modal-close="chapters" type="button">取消</button>
            <button class="cwe-button cwe-button-primary" data-action="start-chapters" type="button">开始提取</button>
          </footer>
        </section>
      </div>
    </div>
  `;
}

export class PanelView {
  private readonly host: HTMLElement;
  private readonly shadow: ShadowRoot;
  private callbacks: PanelCallbacks = NOOP_CALLBACKS;
  private settings: UserSettings | null = null;
  private chapterCount = 0;
  private busy: BusyAction = null;
  private globallyHidden = false;
  // 提取结果是否含参考答案（无答案时隐藏“附加参考答案/附加错题汇总”选项）
  private answerOptionsAvailable = true;
  private mediaQuery: MediaQueryList | null = null;
  private readonly onColorSchemeChange = (): void => this.applyResolvedTheme();

  constructor() {
    const existing = document.getElementById(UI_HOST_ID);
    existing?.remove();
    this.host = document.createElement('div');
    this.host.id = UI_HOST_ID;
    this.shadow = this.host.attachShadow({ mode: 'open' });
    this.shadow.innerHTML = template();
    document.body.appendChild(this.host);
    this.bindEvents();
    this.enableDragging();
  }

  setCallbacks(callbacks: PanelCallbacks): void {
    this.callbacks = callbacks;
  }

  open(): void {
    this.panel().dataset.open = 'true';
    this.launcher().classList.add('cwe-hidden');
  }

  close(): void {
    this.closeAllModals();
    this.panel().dataset.open = 'false';
    if (!this.globallyHidden) this.launcher().classList.remove('cwe-hidden');
  }

  toggle(): void {
    if (this.panel().dataset.open === 'true') this.close();
    else this.open();
  }

  toggleGlobalVisibility(): boolean {
    this.globallyHidden = !this.globallyHidden;
    this.shell().hidden = this.globallyHidden;
    return this.globallyHidden;
  }

  setStatus(message: string, kind: StatusKind = 'neutral'): void {
    const status = this.element<HTMLElement>('.cwe-status');
    status.textContent = message;
    status.dataset.kind = kind;
  }

  setBusy(action: BusyAction, message?: string): void {
    this.busy = action;
    const controls = this.shadow.querySelectorAll<HTMLButtonElement>(
      '[data-action="extract"], [data-action="chapters"], [data-action="download"], [data-action="copy"]',
    );
    controls.forEach((button) => {
      button.disabled = action !== null || this.shouldDisableButton(button.dataset.action ?? '');
    });
    const chapterBusy = action === 'chapters';
    this.element<HTMLButtonElement>('[data-action="start-chapters"]').disabled = chapterBusy;
    this.shadow.querySelectorAll<HTMLInputElement>('[data-chapter-list] input').forEach((input) => {
      input.disabled = chapterBusy;
    });

    const extract = this.element<HTMLButtonElement>('[data-action="extract"]');
    const chapters = this.element<HTMLButtonElement>('[data-action="chapters"]');
    const download = this.element<HTMLButtonElement>('[data-action="download"]');
    const copy = this.element<HTMLButtonElement>('[data-action="copy"]');
    extract.innerHTML = action === 'extract' ? '<span class="cwe-progress"></span>提取中' : '提取当前页面';
    chapters.innerHTML = action === 'chapters' ? '<span class="cwe-progress"></span>遍历中' : '提取多个章节';
    download.innerHTML = action === 'download' ? '<span class="cwe-progress"></span>生成中' : '下载文件';
    copy.innerHTML = action === 'copy' ? '<span class="cwe-progress"></span>复制中' : '复制文本';
    if (message) this.setStatus(message, 'neutral');
  }

  renderResult(result: ExtractionResult): void {
    for (const name of TYPE_STAT_KEYS) {
      const count = result.statistics.byType[name];
      this.stat(name).textContent = String(count);
      this.setStatVisible(name, count > 0);
    }
    this.element<HTMLInputElement>('[data-field="filename"]').value = sanitizeFilename(result.title);
    // 提取结果无参考答案时，隐藏“附加参考答案/附加错题汇总”选项
    this.setAnswerOptionsAvailable(result.statistics.withCorrectAnswer > 0);
    this.refreshActionAvailability(true);
    this.updateFilenamePreview();
    const chapterMessage = result.chapters?.length ? `，来自 ${result.chapters.length} 个章节` : '';
    this.setStatus(`已提取 ${result.statistics.total} 道题${chapterMessage}`, 'success');
  }

  clearResult(): void {
    for (const name of TYPE_STAT_KEYS) {
      this.stat(name).textContent = '0';
      this.setStatVisible(name, false);
    }
    // 重置为默认显示，等待下次提取结果决定
    this.setAnswerOptionsAvailable(true);
    this.refreshActionAvailability(false);
  }

  setChapterCount(count: number): void {
    this.chapterCount = count;
    const button = this.element<HTMLButtonElement>('[data-action="chapters"]');
    button.disabled = this.busy !== null || count < 2;
    button.title = count < 2 ? '当前页面未识别到多个章节入口' : `已识别 ${count} 个章节`;
    const split = this.option('splitByChapter');
    split.disabled = count < 2;
    if (count < 2) split.checked = false;
  }

  readExportOptions(): ExportOptions {
    const formatElement = this.shadow.querySelector<HTMLInputElement>('input[name="cwe-format"]:checked');
    const format = formatElement?.value === 'txt' || formatElement?.value === 'md' ? formatElement.value : 'word';
    return {
      format,
      filename: this.element<HTMLInputElement>('[data-field="filename"]').value,
      withAnswers: this.option('withAnswers').checked,
      withWrong: this.option('withWrong').checked,
      includeAnalysis: this.option('includeAnalysis').checked,
      shuffle: this.option('shuffle').checked,
      bankImport: format === 'word' && this.option('bankImport').checked,
      splitByChapter: this.option('splitByChapter').checked,
    };
  }

  readExportPreferences(): ExportPreferences {
    const options = this.readExportOptions();
    return {
      format: options.format,
      withAnswers: options.withAnswers,
      withWrong: options.withWrong,
      includeAnalysis: options.includeAnalysis,
      shuffle: options.shuffle,
      bankImport: options.bankImport,
      splitByChapter: options.splitByChapter,
    };
  }

  applyExportOptions(options: ExportOptions): void {
    this.element<HTMLInputElement>('[data-field="filename"]').value = options.filename;
    const format = this.shadow.querySelector<HTMLInputElement>(
      `input[name="cwe-format"][value="${options.format}"]`,
    );
    if (format) format.checked = true;
    this.option('withAnswers').checked = options.withAnswers;
    this.option('withWrong').checked = options.withWrong;
    this.option('includeAnalysis').checked = options.includeAnalysis;
    this.option('shuffle').checked = options.shuffle;
    this.option('bankImport').checked = options.bankImport;
    this.option('splitByChapter').checked = options.splitByChapter && this.chapterCount > 1;
    this.updateOptionAvailability();
  }

  applySettings(settings: UserSettings): void {
    this.settings = settings;
    this.element<HTMLSelectElement>('[data-setting="theme"]').value = settings.theme;
    this.element<HTMLInputElement>('[data-setting="shortcut"]').value = formatShortcut(settings.shortcut);
    this.element<HTMLInputElement>('[data-setting="hideShortcut"]').value = formatShortcut(
      settings.hideShortcut,
    );
    this.element<HTMLInputElement>('[data-setting="enableDrag"]').checked = settings.enableDrag;
    this.element<HTMLInputElement>('[data-setting="rememberPanelPosition"]').checked =
      settings.rememberPanelPosition;
    this.element<HTMLInputElement>('[data-setting="autoExtractOnLoad"]').checked =
      settings.autoExtractOnLoad;
    const preferences = settings.exportPreferences;
    this.applyExportOptions({ ...preferences, filename: this.element<HTMLInputElement>('[data-field="filename"]').value });
    this.applyResolvedTheme();
    this.applyPosition(settings.panelPosition);
  }

  renderHistory(entries: readonly HistoryEntry[]): void {
    const list = this.element<HTMLElement>('[data-history-list]');
    list.replaceChildren();
    if (entries.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'cwe-history-empty';
      empty.textContent = '暂无历史记录';
      list.appendChild(empty);
      return;
    }

    const fmtNames: Record<string, string> = { txt: 'TXT', md: 'MD', word: 'Word' };
    for (const entry of entries) {
      const item = document.createElement('article');
      item.className = 'cwe-history-item';
      item.dataset.historyId = entry.id;

      const info = document.createElement('div');
      info.className = 'cwe-history-info';

      const title = document.createElement('div');
      title.className = 'cwe-history-title';
      title.textContent = entry.title;
      title.title = entry.title;

      // 与主脚本一致：按导出配置显示“含答案、含错题、打乱、题库导入”标志
      const flags: string[] = [];
      if (entry.options.withAnswers) flags.push('含答案');
      if (entry.options.withWrong) flags.push('含错题');
      if (entry.options.shuffle) flags.push('打乱');
      if (entry.options.bankImport) flags.push('题库导入');
      const flagStr = flags.length > 0 ? ` · ${flags.join('、')}` : '';
      const fmtName = fmtNames[entry.options.format] || entry.options.format.toUpperCase();

      const meta = document.createElement('div');
      meta.className = 'cwe-history-meta';
      meta.textContent = `${new Date(entry.createdAt).toLocaleString()} · ${entry.result.statistics.total}题 · ${fmtName}${flagStr}`;

      const downloadBtn = document.createElement('button');
      downloadBtn.type = 'button';
      downloadBtn.className = 'cwe-history-download';
      downloadBtn.title = '重新下载';
      downloadBtn.textContent = '⤓';
      downloadBtn.dataset.historyAction = 'download';

      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'cwe-history-delete';
      deleteBtn.title = '删除';
      deleteBtn.textContent = '✕';
      deleteBtn.dataset.historyAction = 'delete';

      info.append(title, meta);
      item.append(info, downloadBtn, deleteBtn);
      list.appendChild(item);
    }
  }

  openHistory(): void {
    this.openModal('history');
  }

  closeHistory(): void {
    this.closeModal('history');
  }

  showChapterDialog(chapters: readonly ChapterDescriptor[]): void {
    const list = this.element<HTMLElement>('[data-chapter-list]');
    list.replaceChildren();
    chapters.forEach((chapter) => {
      const label = document.createElement('label');
      label.className = 'cwe-chapter';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = true;
      input.value = String(chapter.index);
      const text = document.createElement('span');
      text.textContent = chapter.title;
      label.append(input, text);
      list.appendChild(label);
    });
    this.element<HTMLElement>('[data-chapter-progress]').textContent = `共识别 ${chapters.length} 个章节`;
    this.openModal('chapters');
  }

  updateChapterProgress(progress: ChapterProgress): void {
    const stateLabel = progress.state === 'loading' ? '正在加载' : progress.state === 'success' ? '已完成' : '失败';
    this.element<HTMLElement>('[data-chapter-progress]').textContent = `${progress.completed}/${progress.total} · ${stateLabel}：${progress.chapter.title}${progress.message ? `（${progress.message}）` : ''}`;
  }

  closeChapterDialog(): void {
    this.closeModal('chapters');
  }

  destroy(): void {
    this.mediaQuery?.removeEventListener('change', this.onColorSchemeChange);
    this.host.remove();
  }

  private bindEvents(): void {
    this.launcher().addEventListener('click', () => this.open());
    this.element('[data-action="close"]').addEventListener('click', () => this.close());
    this.element('[data-action="extract"]').addEventListener('click', () => this.callbacks.onExtract());
    this.element('[data-action="chapters"]').addEventListener('click', () => this.callbacks.onOpenChapters());
    this.element('[data-action="download"]').addEventListener('click', () => this.callbacks.onDownload());
    this.element('[data-action="copy"]').addEventListener('click', () => this.callbacks.onCopy());
    this.element('[data-action="settings"]').addEventListener('click', () => this.openModal('settings'));
    this.element('[data-action="history"]').addEventListener('click', () => this.openHistory());

    this.shadow.querySelectorAll<HTMLInputElement>('input[name="cwe-format"], [data-option]').forEach((input) => {
      input.addEventListener('change', () => {
        this.updateOptionAvailability(input);
        this.callbacks.onPreferencesChange(this.readExportPreferences());
      });
    });

    this.element('[data-action="save-settings"]').addEventListener('click', () => this.saveSettings());
    this.element('[data-action="chapter-all"]').addEventListener('click', () => this.setAllChapters(true));
    this.element('[data-action="chapter-none"]').addEventListener('click', () => this.setAllChapters(false));
    this.element('[data-action="start-chapters"]').addEventListener('click', () => {
      const indexes = [...this.shadow.querySelectorAll<HTMLInputElement>('[data-chapter-list] input:checked')]
        .map((input) => Number.parseInt(input.value, 10))
        .filter(Number.isFinite);
      this.callbacks.onExtractChapters(indexes);
    });

    this.shadow.querySelectorAll<HTMLElement>('[data-modal-close]').forEach((button) => {
      button.addEventListener('click', () => this.closeModal(button.dataset.modalClose ?? ''));
    });
    this.shadow.querySelectorAll<HTMLElement>('.cwe-modal-backdrop').forEach((backdrop) => {
      backdrop.addEventListener('click', (event) => {
        if (event.target === backdrop) this.closeModal(backdrop.dataset.modal ?? '');
      });
    });

    this.element<HTMLElement>('[data-history-list]').addEventListener('click', (event) => {
      const item = event.target instanceof Element ? event.target.closest<HTMLElement>('[data-history-id]') : null;
      const id = item?.dataset.historyId;
      if (!id) return;
      const actionBtn = event.target instanceof Element
        ? event.target.closest<HTMLButtonElement>('[data-history-action]')
        : null;
      // 与主脚本一致：点击条目本身恢复，图标按钮分别触发重新下载/删除
      if (actionBtn) {
        const action = actionBtn.dataset.historyAction;
        if (action === 'download') this.callbacks.onHistoryDownload(id);
        else if (action === 'delete') this.callbacks.onHistoryDelete(id);
      } else {
        this.callbacks.onHistoryRestore(id);
      }
    });

    window.addEventListener('keydown', (event) => {
      if (event.key !== 'Escape') return;
      const openModal = this.shadow.querySelector<HTMLElement>('.cwe-modal-backdrop[data-open="true"]');
      if (openModal) this.closeModal(openModal.dataset.modal ?? '');
      else if (this.panel().dataset.open === 'true') this.close();
    });
  }

  private enableDragging(): void {
    const panel = this.panel();
    const header = this.element<HTMLElement>('.cwe-header');
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;
    let dragging = false;

    header.addEventListener('pointerdown', (event) => {
      if (!this.settings?.enableDrag || event.button !== 0) return;
      if (event.target instanceof Element && event.target.closest('button')) return;
      const rectangle = panel.getBoundingClientRect();
      dragging = true;
      startX = event.clientX;
      startY = event.clientY;
      startLeft = rectangle.left;
      startTop = rectangle.top;
      header.setPointerCapture(event.pointerId);
      event.preventDefault();
    });

    header.addEventListener('pointermove', (event) => {
      if (!dragging) return;
      const maxLeft = Math.max(0, window.innerWidth - panel.offsetWidth);
      const maxTop = Math.max(0, window.innerHeight - 48);
      const left = Math.min(maxLeft, Math.max(0, startLeft + event.clientX - startX));
      const top = Math.min(maxTop, Math.max(0, startTop + event.clientY - startY));
      panel.style.left = `${left}px`;
      panel.style.top = `${top}px`;
      panel.style.right = 'auto';
    });

    const finish = (event: PointerEvent): void => {
      if (!dragging) return;
      dragging = false;
      if (header.hasPointerCapture(event.pointerId)) header.releasePointerCapture(event.pointerId);
      const rectangle = panel.getBoundingClientRect();
      this.callbacks.onPanelPositionChange({ left: rectangle.left, top: rectangle.top });
    };
    header.addEventListener('pointerup', finish);
    header.addEventListener('pointercancel', finish);
  }

  private updateOptionAvailability(changed?: HTMLInputElement): void {
    const selectedFormat = this.shadow.querySelector<HTMLInputElement>('input[name="cwe-format"]:checked')?.value;
    const bank = this.option('bankImport');
    if (selectedFormat !== 'word') bank.checked = false;
    if (changed === bank && bank.checked) {
      const word = this.shadow.querySelector<HTMLInputElement>('input[name="cwe-format"][value="word"]');
      if (word) word.checked = true;
      this.option('withAnswers').checked = false;
      this.option('withWrong').checked = false;
      this.option('shuffle').checked = false;
      this.option('includeAnalysis').checked = false;
    }

    const bankEnabled = bank.checked;
    if (bankEnabled) {
      this.option('withAnswers').checked = false;
      this.option('withWrong').checked = false;
      this.option('shuffle').checked = false;
      this.option('includeAnalysis').checked = false;
    }
    // 无参考答案时强制取消勾选（覆盖历史记录恢复的偏好）
    if (!this.answerOptionsAvailable) {
      this.option('withAnswers').checked = false;
      this.option('withWrong').checked = false;
    }
    this.option('withAnswers').disabled = bankEnabled;
    this.option('withWrong').disabled = bankEnabled;
    this.option('shuffle').disabled = bankEnabled;
    this.option('includeAnalysis').disabled =
      bankEnabled || !(this.option('withAnswers').checked || this.option('withWrong').checked);
    // 题库导入仅在 Word 格式下可用，TXT/MD 格式下始终禁用
    this.option('bankImport').disabled = selectedFormat !== 'word';
    this.option('splitByChapter').disabled = this.chapterCount < 2;
    this.updateFilenamePreview();
  }

  // 文件名预览：附加答案时追加“（含答案）”，题库导入时追加“（题库导入）”，取消勾选时移除
  private updateFilenamePreview(): void {
    const input = this.element<HTMLInputElement>('[data-field="filename"]');
    const name = input.value.replace(/(?:（含答案）|（题库导入）)+$/, '');
    const suffix = this.option('bankImport').checked
      ? '（题库导入）'
      : this.option('withAnswers').checked
        ? '（含答案）'
        : '';
    input.value = name + suffix;
  }

  // 控制“附加参考答案/附加错题汇总”选项的显示与勾选状态
  private setAnswerOptionsAvailable(available: boolean): void {
    this.answerOptionsAvailable = available;
    this.option('withAnswers').closest<HTMLElement>('.cwe-check')?.classList.toggle('cwe-hidden', !available);
    this.option('withWrong').closest<HTMLElement>('.cwe-check')?.classList.toggle('cwe-hidden', !available);
    if (!available) {
      this.option('withAnswers').checked = false;
      this.option('withWrong').checked = false;
    }
  }

  private refreshActionAvailability(hasResult: boolean): void {
    const download = this.element<HTMLButtonElement>('[data-action="download"]');
    const copy = this.element<HTMLButtonElement>('[data-action="copy"]');
    download.dataset.hasResult = String(hasResult);
    copy.dataset.hasResult = String(hasResult);
    download.disabled = this.busy !== null || !hasResult;
    copy.disabled = this.busy !== null || !hasResult;
  }

  private shouldDisableButton(action: string): boolean {
    if (action === 'chapters') return this.chapterCount < 2;
    if (action === 'download' || action === 'copy') {
      return this.element<HTMLButtonElement>(`[data-action="${action}"]`).dataset.hasResult !== 'true';
    }
    return false;
  }

  private saveSettings(): void {
    if (!this.settings) return;
    const rememberPanelPosition = this.element<HTMLInputElement>(
      '[data-setting="rememberPanelPosition"]',
    ).checked;
    const settings: UserSettings = {
      ...this.settings,
      theme: this.element<HTMLSelectElement>('[data-setting="theme"]').value as UserSettings['theme'],
      shortcut: parseShortcut(
        this.element<HTMLInputElement>('[data-setting="shortcut"]').value,
        this.settings.shortcut,
      ),
      hideShortcut: parseShortcut(
        this.element<HTMLInputElement>('[data-setting="hideShortcut"]').value,
        this.settings.hideShortcut,
      ),
      enableDrag: this.element<HTMLInputElement>('[data-setting="enableDrag"]').checked,
      rememberPanelPosition,
      autoExtractOnLoad:
        this.element<HTMLInputElement>('[data-setting="autoExtractOnLoad"]').checked,
      panelPosition: rememberPanelPosition ? this.settings.panelPosition : null,
      exportPreferences: this.readExportPreferences(),
    };
    this.callbacks.onSettingsSave(settings);
    this.closeModal('settings');
    this.setStatus('设置已保存', 'success');
  }

  private applyResolvedTheme(): void {
    const mode = this.settings?.theme ?? 'auto';
    this.mediaQuery?.removeEventListener('change', this.onColorSchemeChange);
    this.mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    if (mode === 'auto') this.mediaQuery.addEventListener('change', this.onColorSchemeChange);
    const dark = mode === 'dark' || (mode === 'auto' && this.mediaQuery.matches);
    this.host.dataset.theme = dark ? 'dark' : 'light';
  }

  private applyPosition(position: PanelPosition | null): void {
    const panel = this.panel();
    if (!position || !this.settings?.rememberPanelPosition) {
      panel.style.removeProperty('left');
      panel.style.removeProperty('top');
      panel.style.removeProperty('right');
      return;
    }
    const left = Math.min(Math.max(0, position.left), Math.max(0, window.innerWidth - panel.offsetWidth));
    const top = Math.min(Math.max(0, position.top), Math.max(0, window.innerHeight - 48));
    panel.style.left = `${left}px`;
    panel.style.top = `${top}px`;
    panel.style.right = 'auto';
  }

  private setAllChapters(checked: boolean): void {
    this.shadow.querySelectorAll<HTMLInputElement>('[data-chapter-list] input').forEach((input) => {
      input.checked = checked;
    });
  }

  private openModal(name: string): void {
    this.closeAllModals();
    const modal = this.shadow.querySelector<HTMLElement>(`[data-modal="${name}"]`);
    if (modal) modal.dataset.open = 'true';
  }

  private closeModal(name: string): void {
    const modal = this.shadow.querySelector<HTMLElement>(`[data-modal="${name}"]`);
    if (modal) modal.dataset.open = 'false';
  }

  private closeAllModals(): void {
    this.shadow.querySelectorAll<HTMLElement>('[data-modal]').forEach((modal) => {
      modal.dataset.open = 'false';
    });
  }

  private shell(): HTMLElement {
    return this.element('.cwe-shell');
  }

  private launcher(): HTMLButtonElement {
    return this.element('.cwe-launcher');
  }

  private panel(): HTMLElement {
    return this.element('.cwe-panel');
  }

  private option(name: keyof ExportPreferences): HTMLInputElement {
    return this.element(`[data-option="${name}"]`);
  }

  private stat(name: string): HTMLElement {
    return this.element(`[data-stat="${name}"]`);
  }

  private setStatVisible(name: string, visible: boolean): void {
    this.stat(name).closest<HTMLElement>('.cwe-stat')?.classList.toggle('cwe-hidden', !visible);
  }

  private element<E extends Element = HTMLElement>(selector: string): E {
    return queryRequired<E>(this.shadow, selector);
  }
}
