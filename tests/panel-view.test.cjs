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
    ['withAnswers', 'withWrong', 'shuffle', 'bankImport', 'splitByChapter', 'splitBySection'],
  );
  assert.equal(shadow.querySelector('[data-option="includeAnalysis"]'), null);
  // 面板读出的导出选项里也不再带该字段
  assert.deepEqual(Object.keys(panel.readExportOptions()).sort(), [
    'bankImport',
    'filename',
    'format',
    'shuffle',
    'splitByChapter',
    'splitBySection',
    'withAnswers',
    'withWrong',
  ]);

  panel.destroy();
});

test('「按章节拆分」与「按小节拆分」互斥：勾一个自动取消另一个', () => {
  const { panel, shadow } = mountPanel();
  panel.setChapterCount(5);
  const byChapter = shadow.querySelector('[data-option="splitByChapter"]');
  const bySection = shadow.querySelector('[data-option="splitBySection"]');
  assert.equal(byChapter.disabled, false);
  assert.equal(bySection.disabled, false);

  byChapter.checked = true;
  byChapter.dispatchEvent(new Event('change'));
  assert.equal(byChapter.checked, true);
  assert.equal(bySection.checked, false);

  bySection.checked = true;
  bySection.dispatchEvent(new Event('change'));
  assert.equal(bySection.checked, true);
  assert.equal(byChapter.checked, false, '勾「按小节」要把「按章节」取消');

  byChapter.checked = true;
  byChapter.dispatchEvent(new Event('change'));
  assert.equal(byChapter.checked, true);
  assert.equal(bySection.checked, false, '勾「按章节」要把「按小节」取消');

  // 章节数不足 2 时两个选项都禁用并清空（与原有「按章节」规则一致）
  panel.setChapterCount(1);
  assert.equal(byChapter.disabled, true);
  assert.equal(bySection.disabled, true);
  assert.equal(byChapter.checked, false);
  assert.equal(bySection.checked, false);
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

test('章节弹窗：两级目录按章分组，章行可整章勾选，状态报出覆盖章数', () => {
  const { panel, shadow } = mountPanel();
  const groupA = { id: 'g1', label: '1', title: '担当复兴大任 成就时代新人' };
  const groupB = { id: 'g2', label: '2', title: '领悟人生真谛 把握人生方向' };
  const chapters = [
    { index: 0, id: 'c1', title: '如何认识新方位', label: '1.1', active: false, group: groupA },
    { index: 1, id: 'c2', title: '如何认识新使命', label: '1.2', active: false, group: groupA },
    {
      index: 2,
      id: 'c3',
      title: '为什么要思考人生问题',
      label: '2.1',
      active: false,
      group: groupB,
    },
  ];
  panel.showChapterDialog(chapters);

  const list = shadow.querySelector('[data-chapter-list]');
  const status = shadow.querySelector('[data-chapter-status]');
  assert.equal(status.textContent, '已选 3/3 小节 · 覆盖 2/2 章');
  // 章行用页面上的章号拼标题，小节显示页面上的真编号
  assert.deepEqual(
    [...list.querySelectorAll('.cwe-chapter-group-head span')].map((node) => node.textContent),
    ['第 1 章 担当复兴大任 成就时代新人', '第 2 章 领悟人生真谛 把握人生方向'],
  );
  assert.deepEqual(
    [...list.querySelectorAll('.cwe-chapter span')].map((node) => node.textContent),
    ['1.1 如何认识新方位', '1.2 如何认识新使命', '2.1 为什么要思考人生问题'],
  );

  // 取消第二章的整章勾选：该章小节全部取消，章行与状态同步
  const groupToggles = [...list.querySelectorAll('input[data-chapter-group-toggle]')];
  groupToggles[1].checked = false;
  groupToggles[1].dispatchEvent(new Event('change', { bubbles: true }));
  assert.equal(status.textContent, '已选 2/3 小节 · 覆盖 1/2 章');
  assert.deepEqual(
    [...list.querySelectorAll('input[data-chapter-index]')].map((input) => input.checked),
    [true, true, false],
  );
  assert.equal(groupToggles[0].indeterminate, false);

  // 只勾第一章里的一个小节 → 章行半选、全选框半选
  list.querySelector('input[data-chapter-index]').checked = false;
  list.dispatchEvent(new Event('change'));
  assert.equal(groupToggles[0].indeterminate, true);
  assert.equal(shadow.querySelector('[data-action="chapter-all-toggle"]').indeterminate, true);
  assert.equal(status.textContent, '已选 1/3 小节 · 覆盖 1/2 章');

  // 全选框勾上后，所有小节与所有章行都选中
  const all = shadow.querySelector('[data-action="chapter-all-toggle"]');
  all.checked = true;
  all.dispatchEvent(new Event('change'));
  assert.equal(status.textContent, '已选 3/3 小节 · 覆盖 2/2 章');
  assert.deepEqual(
    [...list.querySelectorAll('input[data-chapter-group-toggle]')].map((input) => input.checked),
    [true, true],
  );

  // 提取进度里的小节名也带真编号
  panel.updateChapterProgress({ completed: 1, total: 3, chapter: chapters[0], state: 'success' });
  assert.equal(
    shadow.querySelector('[data-chapter-progress]').textContent,
    '已提取 1/3 章节 · 已完成：1.1 如何认识新方位',
  );

  panel.destroy();
});

test('章节弹窗：点「开始提取」后立即收起，一章都没选时保持打开', () => {
  const { panel, shadow } = mountPanel();
  const calls = [];
  panel.setCallbacks({
    onExtract: () => undefined,
    onOpenChapters: () => undefined,
    onExtractChapters: (indexes) => calls.push([...indexes]),
    onDownload: () => undefined,
    onCopy: () => undefined,
    onDiagnose: () => undefined,
    onPreferencesChange: () => undefined,
    onSettingsSave: () => undefined,
    onHistoryRestore: () => undefined,
    onHistoryDownload: () => undefined,
    onHistoryDelete: () => undefined,
    onPanelPositionChange: () => undefined,
  });

  panel.showChapterDialog([chapter(0, '第一章'), chapter(1, '第二章')]);
  const backdrop = shadow.querySelector('[data-modal="chapters"]');
  const start = shadow.querySelector('[data-action="start-chapters"]');
  assert.equal(backdrop.dataset.open, 'true');

  // 确认之后弹窗立刻收起（几十章要提取很久，进度改由面板状态栏显示）
  start.dispatchEvent(new Event('click'));
  assert.deepEqual(calls, [[0, 1]]);
  assert.equal(backdrop.dataset.open, 'false');

  // 一章都没选：保持打开让用户就地补选，等待应用层提示
  panel.showChapterDialog([chapter(0, '第一章'), chapter(1, '第二章')]);
  const inputs = [...shadow.querySelectorAll('[data-chapter-list] input')];
  inputs.forEach((input) => {
    input.checked = false;
  });
  start.dispatchEvent(new Event('click'));
  assert.deepEqual(calls, [[0, 1], []]);
  assert.equal(backdrop.dataset.open, 'true');

  panel.destroy();
});
