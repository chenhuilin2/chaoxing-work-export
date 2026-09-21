// 面板交互测试：章节弹窗的「全选」框与状态文本、导出选项的可用性、设置弹窗的行顺序。
// 面板挂在 Shadow DOM 上，linkedom 支持 attachShadow，因此可以按真实 DOM 语义驱动事件。
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { parseHTML, Event } = require('linkedom');

const { document, window } = parseHTML('<!doctype html><html><head></head><body></body></html>');

// 面板按浏览器语义使用全局 window/document，并在应用主题时读取 matchMedia
globalThis.window = window;
globalThis.document = document;
globalThis.Event = Event;
window.matchMedia = () => ({
  matches: false,
  addEventListener() {},
  removeEventListener() {},
});

const { PanelView } = require('../.tmp/test/src/ui/panel-view.js');

const HOST_ID = 'chaoxing-work-export-root';

/** 挂一个面板并取出它的 Shadow DOM */
function mountPanel() {
  const panel = new PanelView();
  const host = document.getElementById(HOST_ID);
  return { panel, shadow: host.shadowRoot };
}

function chapter(index, title) {
  return { index, id: `chapter-${index}`, title, active: false };
}

/** 面板要求完整导出选项，这里给一份「全不选」的基线 */
function exportOptions(overrides = {}) {
  return {
    format: 'word',
    filename: '卷子',
    withAnswers: false,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
    ...overrides,
  };
}

test('章节弹窗：全选框兼作状态指示，右侧显示已选数量', () => {
  const { panel, shadow } = mountPanel();
  panel.showChapterDialog([chapter(0, '第一章'), chapter(1, '第二章')]);

  const toggle = shadow.querySelector('[data-action="chapter-all-toggle"]');
  const status = shadow.querySelector('[data-chapter-status]');
  assert.ok(shadow.querySelector('[data-modal="chapters"] .cwe-modal-wide'), '章节弹窗应加宽');
  assert.equal(status.textContent, '已选 2 / 共 2 个章节');
  assert.equal(toggle.checked, true);
  assert.equal(toggle.indeterminate, false);
  // 原来的「全选 / 取消全选」两个按钮已被全选框取代
  assert.equal(shadow.querySelector('[data-action="chapter-all"]'), null);
  assert.equal(shadow.querySelector('[data-action="chapter-none"]'), null);

  // 取消一个章节 → 半选态 + 计数变化
  const list = shadow.querySelector('[data-chapter-list]');
  const inputs = [...list.querySelectorAll('input')];
  inputs[1].checked = false;
  list.dispatchEvent(new Event('change'));
  assert.equal(status.textContent, '已选 1 / 共 2 个章节');
  assert.equal(toggle.indeterminate, true);
  assert.equal(toggle.checked, false);

  // 勾上全选框 → 全部选中
  toggle.checked = true;
  toggle.dispatchEvent(new Event('change'));
  assert.equal(
    inputs.every((input) => input.checked),
    true,
  );
  assert.equal(status.textContent, '已选 2 / 共 2 个章节');
  assert.equal(toggle.indeterminate, false);

  panel.destroy();
});

test('导出选项：附加解析已整体移除，其余选项顺序不变', () => {
  const { panel, shadow } = mountPanel();
  assert.deepEqual(
    [...shadow.querySelectorAll('.cwe-options [data-option]')].map((input) => input.dataset.option),
    ['withAnswers', 'withWrong', 'shuffle', 'bankImport', 'splitByChapter'],
  );
  assert.equal(shadow.querySelector('[data-option="includeAnalysis"]'), null);
  // 面板读出的导出选项里也不再带该字段
  assert.deepEqual(Object.keys(panel.readExportOptions()).sort(), [
    'bankImport',
    'filename',
    'format',
    'shuffle',
    'splitByChapter',
    'withAnswers',
    'withWrong',
  ]);

  panel.destroy();
});

test('设置弹窗：开关组在上、快捷键设置移到下方', () => {
  const { panel, shadow } = mountPanel();
  assert.deepEqual(
    [...shadow.querySelectorAll('[data-modal="settings"] [data-setting]')].map(
      (element) => element.dataset.setting,
    ),
    [
      'theme',
      'enableDrag',
      'rememberPanelPosition',
      'autoExtractOnLoad',
      'shortcut',
      'hideShortcut',
    ],
  );
  // 三行开关的标签字重与上方「界面主题」等设置行一致（11.5px / 700）
  const styles = shadow.querySelector('style').textContent;
  assert.match(styles, /\.cwe-switch-row\s*\{[^}]*font-weight:\s*700/);
  assert.match(styles, /\.cwe-setting-label\s*\{[^}]*font-weight:\s*700/);
  panel.destroy();
});

test('章节遍历进度：显示「已提取 N/M 章节」并同步状态栏', () => {
  const { panel, shadow } = mountPanel();
  const progress = shadow.querySelector('[data-chapter-progress]');
  const status = shadow.querySelector('.cwe-status');

  panel.updateChapterProgress({
    completed: 3,
    total: 54,
    chapter: chapter(3, '导论4'),
    state: 'loading',
  });
  assert.equal(progress.textContent, '已提取 3/54 章节 · 正在加载：导论4');
  assert.match(status.textContent, /已提取 3\/54 章节/);

  panel.updateChapterProgress({
    completed: 4,
    total: 54,
    chapter: chapter(4, '导论5'),
    state: 'failed',
    message: '未在限定时间内识别到题目',
  });
  assert.equal(progress.textContent, '已提取 4/54 章节 · 失败：导论5（未在限定时间内识别到题目）');
  assert.equal(status.dataset.kind, 'warning');

  panel.updateChapterProgress({
    completed: 4,
    total: 54,
    chapter: chapter(4, '导论5'),
    state: 'success',
  });
  assert.equal(progress.textContent, '已提取 4/54 章节 · 已完成：导论5');
  assert.equal(status.dataset.kind, 'neutral');

  panel.destroy();
});
