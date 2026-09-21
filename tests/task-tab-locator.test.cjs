// 任务卡定位器用例：学习通知识卡片页默认停在「视频」卡，
// 题目卡（章节测验/作业/考试）必须被切换激活，否则题目 iframe 不会被加载。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument } = require('./dom-setup.cjs');
const {
  QuestionTabGuard,
  TaskTabLocator,
} = require('../.tmp/test/src/application/task-tab-locator.js');
const { collectAccessibleDocuments } = require('../.tmp/test/src/utils/dom.js');

/** 与 html/章节测验.html 一致的任务卡栏，默认激活「视频」 */
const TAB_BAR_HTML = `
<div class="prev_list" id="prev_tab">
  <ul class="prev_ul clearfix">
    <li class="active" title="视频" cardid="1198024984" onclick="changeDisplayContent(1,2,'1200668686','265608653','151187867','');" tabindex="0" role="option" id="dct1">
      <div class="prev_white" aria-hidden="true"><span class="num">1</span><span class="spanText">视频</span></div>
    </li>
    <li title="章节测验" cardid="1198024985" onclick="changeDisplayContent(2,2,'1200668686','265608653','151187867','');" tabindex="0" role="option" id="dct2" class="c2">
      <div class="prev_white" aria-hidden="true"><span class="num">2</span><span class="spanText">章节测验</span></div>
    </li>
  </ul>
</div>`;

const VIDEO_ONLY_TAB_BAR_HTML = `
<div class="prev_list" id="prev_tab">
  <ul class="prev_ul clearfix">
    <li class="active" title="视频" role="option" id="dct1"><div class="prev_white" aria-hidden="true"><span class="num">1</span><span class="spanText">视频</span></div></li>
    <li title="文档" role="option" id="dct2"><div class="prev_white" aria-hidden="true"><span class="num">2</span><span class="spanText">文档</span></div></li>
  </ul>
</div>`;

/** 记录冒泡到任务卡栏的点击目标，确认点中的是任务卡 li 本身 */
function captureClickTargets(element) {
  const targets = [];
  element.addEventListener('click', (event) => {
    targets.push(event.target);
  });
  return targets;
}

test('解析任务卡列表与激活状态', () => {
  const locator = new TaskTabLocator(createDocument(TAB_BAR_HTML));
  assert.deepEqual(locator.list(), [
    { index: 0, id: 'dct1', title: '视频', active: true },
    { index: 1, id: 'dct2', title: '章节测验', active: false },
  ]);
  assert.equal(locator.activeIndex(), 0);
  assert.equal(locator.questionTabIndex(), 1);
});

test('默认停在视频卡时切换到题目卡', () => {
  const root = createDocument(TAB_BAR_HTML);
  const locator = new TaskTabLocator(root);
  const targets = captureClickTargets(root.getElementById('prev_tab'));

  assert.equal(locator.ensureQuestionTab(), 'switched');
  assert.deepEqual(targets, [root.getElementById('dct2')]);
});

test('题目卡已激活时不重复点击', () => {
  const root = createDocument(TAB_BAR_HTML.replace('class="c2"', 'class="c2 active"'));
  const locator = new TaskTabLocator(root);
  const targets = captureClickTargets(root.getElementById('prev_tab'));

  assert.equal(locator.ensureQuestionTab(), 'active');
  assert.deepEqual(targets, []);
});

test('页面没有任务卡栏时返回 missing', () => {
  const locator = new TaskTabLocator(createDocument('<div class="Zy_TItle">题目</div>'));
  assert.equal(locator.present(), false);
  assert.equal(locator.ensureQuestionTab(), 'missing');
  assert.deepEqual(locator.list(), []);
});

test('任务卡栏存在但没有题目卡时返回 missing', () => {
  const locator = new TaskTabLocator(createDocument(VIDEO_ONLY_TAB_BAR_HTML));
  assert.equal(locator.present(), true);
  assert.equal(locator.questionTabIndex(), null);
  assert.equal(locator.ensureQuestionTab(), 'missing');
});

test('任务卡文案取 title 属性，缺失时回退到 spanText 并去掉序号', () => {
  const locator = new TaskTabLocator(
    createDocument(`
      <div class="prev_list" id="prev_tab">
        <ul class="prev_ul clearfix">
          <li role="option" id="dct1"><div class="prev_white"><span class="num">1</span><span class="spanText">作业</span></div></li>
          <li role="option" id="dct2"><div class="prev_white"><span class="num">2</span>考试</div></li>
        </ul>
      </div>`),
  );
  assert.deepEqual(
    locator.list().map((tab) => tab.title),
    ['作业', '考试'],
  );
  assert.equal(locator.questionTabIndex(), 0);
});

test('补切节流器：冷却期内不重复点击，超过次数上限后不再补切', () => {
  const root = createDocument(TAB_BAR_HTML);
  const targets = captureClickTargets(root.getElementById('prev_tab'));
  const guard = new QuestionTabGuard(new TaskTabLocator(root));

  assert.equal(guard.ensure(1_000), true); // 第 1 次补切
  assert.equal(guard.ensure(1_200), false); // 冷却期内
  assert.equal(guard.ensure(1_400), true); // 冷却结束
  assert.equal(guard.ensure(1_600), false);
  assert.equal(guard.ensure(1_800), true);
  assert.equal(guard.ensure(2_200), true);
  // 已用满 4 次：即使冷却结束也不再补切，避免页面加载过程中反复触发加载
  assert.equal(guard.ensure(9_000), false);
  assert.equal(targets.length, 4);
});

test('任务卡栏位于同源 iframe 内时也能定位', () => {
  const outer = createDocument('<iframe id="frame"></iframe>');
  const inner = createDocument(TAB_BAR_HTML);
  Object.defineProperty(outer.getElementById('frame'), 'contentDocument', {
    get: () => inner,
  });

  // 公共遍历工具能跨层收集文档，定位器据此找到 iframe 内的任务卡栏
  assert.deepEqual(
    collectAccessibleDocuments(outer).map((entry) => entry.depth),
    [0, 1],
  );
  const locator = new TaskTabLocator(outer);
  assert.equal(locator.present(), true);
  assert.equal(locator.ensureQuestionTab(), 'switched');
});
