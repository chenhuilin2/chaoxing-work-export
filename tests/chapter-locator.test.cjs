// 章节定位器用例：覆盖学习通目录树的真实结构（span.posCatalog_name 才是可点章节，
// .firstLayer / .posCatalog_title 是分组标题；分组标题同时充当导出的「章」）。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument } = require('./dom-setup.cjs');
const { ChapterLocator } = require('../.tmp/test/src/application/chapter-locator.js');
const { groupHeadingText, sectionHeadingText } = require('../.tmp/test/src/domain/chapter.js');

/** 与 html/章节练习完整页.html 一致的目录树：一个分组标题 + 两个章节 */
const CATALOG_HTML = `
<div id="coursetree">
  <div class="posCatalog_list">
    <ul>
      <li>
        <div class="posCatalog_select firstLayer" id="1200668675">
          <span class="posCatalog_title posCatalog_rotate titleIcon" title="导论 马克思主义中国化时代化的历史进程与理论成果"><em class="posCatalog_sbar">1</em> 导论 马克思主义中国化时代化的历史进程与理论成果</span>
        </div>
        <div class="posCatalog_level" style="display: block;">
          <ul>
            <li>
              <div class="posCatalog_select posCatalog_active" id="cur1200668686">
                <span class="posCatalog_name" title="导论1" onclick="getTeacherAjax('265608653','151187867','1200668686');"><em class="posCatalog_num">1</em>导论1</span>
              </div>
            </li>
            <li>
              <div class="posCatalog_select" id="1200670000">
                <span class="posCatalog_name" title="导论2" onclick="getTeacherAjax('265608653','151187867','1200670000');"><em class="posCatalog_num">2</em>导论2</span>
              </div>
            </li>
          </ul>
        </div>
      </li>
    </ul>
  </div>
</div>`;

/**
 * 记录冒泡到此处的点击目标。
 * linkedom 不执行行内 onclick，也无法断言「点了哪个节点」，因此用事件冒泡的
 * event.target 来确认点击确实落在带 onclick 的节点上（点击事件会冒泡到容器）。
 */
function captureClickTargets(element) {
  const targets = [];
  element.addEventListener('click', (event) => {
    targets.push(event.target);
  });
  return targets;
}

test('章节列表排除分组标题，只保留真实章节', () => {
  const root = createDocument(CATALOG_HTML);
  const chapters = new ChapterLocator(root).list();
  assert.deepEqual(
    chapters.map((chapter) => chapter.title),
    ['导论1', '导论2'],
  );
});

test('识别当前激活章节：激活类名挂在容器 div 上也能命中', () => {
  const root = createDocument(CATALOG_HTML);
  assert.equal(new ChapterLocator(root).activeIndex(), 0);
});

test('切换章节时点击带 onclick 的名字节点，而不是无事件的容器 div', () => {
  const root = createDocument(CATALOG_HTML);
  const container = root.getElementById('1200670000');
  const nameNode = container.querySelector('.posCatalog_name');
  const targets = captureClickTargets(container);

  assert.equal(new ChapterLocator(root).activate(1), true);
  assert.deepEqual(targets, [nameNode]);
});

test('章节条目匹配到容器 div 时，仍会下钻到内部的 onclick 节点', () => {
  // 老结构：章节容器不带 .posCatalog_name，只能走兜底选择器
  const root = createDocument(`
    <div class="chapter-list">
      <div class="posCatalog_select" id="cur1200668686"><span id="chapter-name-a" onclick="getTeacherAjax('a');">导论1</span></div>
      <div class="posCatalog_select" id="1200670000"><span id="chapter-name-b" onclick="getTeacherAjax('b');">导论2</span></div>
    </div>`);
  const locator = new ChapterLocator(root);
  const container = root.getElementById('1200670000');
  const nameNode = root.getElementById('chapter-name-b');
  const targets = captureClickTargets(container);

  assert.deepEqual(
    locator.list().map((chapter) => chapter.title),
    ['导论1', '导论2'],
  );
  assert.equal(locator.activate(1), true);
  assert.deepEqual(targets, [nameNode]);
});

test('分组标题自身被识别为分组时不进入章节列表', () => {
  // 分组标题内部若也用了 posCatalog_name，仍应被 .firstLayer 排除
  const root = createDocument(`
    <div class="posCatalog_list">
      <div class="posCatalog_select firstLayer" id="group-1"><span class="posCatalog_name" title="第一章">第一章</span></div>
      <div class="posCatalog_select" id="cur-1"><span class="posCatalog_name" title="1.1 小节" onclick="getTeacherAjax('x');">1.1 小节</span></div>
      <div class="posCatalog_select" id="cur-2"><span class="posCatalog_name" title="1.2 小节" onclick="getTeacherAjax('y');">1.2 小节</span></div>
    </div>`);
  assert.deepEqual(
    new ChapterLocator(root).list().map((chapter) => chapter.title),
    ['1.1 小节', '1.2 小节'],
  );
});

/** 真实两级结构：章是 .firstLayer 分组标题，小节在 .posCatalog_level 里，编号写在 em 里 */
const NESTED_CATALOG_HTML = `
<div id="coursetree">
  <div class="posCatalog_list">
    <ul>
      <li>
        <div class="posCatalog_select firstLayer" id="g1">
          <span class="posCatalog_title" title="担当复兴大任 成就时代新人"><em class="posCatalog_sbar">1</em> 担当复兴大任 成就时代新人</span>
        </div>
        <div class="posCatalog_level">
          <ul>
            <li><div class="posCatalog_select" id="cur1"><span class="posCatalog_name" title="如何认识新方位" onclick="getTeacherAjax('a');"><em class="posCatalog_sbar">1.1</em> 如何认识新方位</span></div></li>
            <li><div class="posCatalog_select" id="cur2"><span class="posCatalog_name" title="如何认识新使命" onclick="getTeacherAjax('b');"><em class="posCatalog_sbar">1.2</em> 如何认识新使命</span></div></li>
          </ul>
        </div>
      </li>
      <li>
        <div class="posCatalog_select firstLayer" id="g2">
          <span class="posCatalog_title" title="领悟人生真谛 把握人生方向"><em class="posCatalog_sbar">2</em> 领悟人生真谛 把握人生方向</span>
        </div>
        <div class="posCatalog_level">
          <ul>
            <li><div class="posCatalog_select" id="cur3"><span class="posCatalog_name" title="为什么要思考人生问题" onclick="getTeacherAjax('c');"><em class="posCatalog_sbar">2.1</em> 为什么要思考人生问题</span></div></li>
          </ul>
        </div>
      </li>
    </ul>
  </div>
</div>`;

test('两级目录：小节挂到所属章上，名字与编号分别读 title 属性与页面 em', () => {
  const chapters = new ChapterLocator(createDocument(NESTED_CATALOG_HTML)).list();

  assert.deepEqual(
    chapters.map((chapter) => chapter.title),
    ['如何认识新方位', '如何认识新使命', '为什么要思考人生问题'],
  );
  assert.deepEqual(
    chapters.map((chapter) => chapter.label),
    ['1.1', '1.2', '2.1'],
  );
  assert.deepEqual(
    chapters.map((chapter) => chapter.group?.id),
    ['g1', 'g1', 'g2'],
  );
  // 章名取分组标题的 title 属性，章号取它内部的 em
  assert.deepEqual(chapters[0].group, {
    id: 'g1',
    label: '1',
    title: '担当复兴大任 成就时代新人',
  });
  // 面板 / 导出用的显示名：章用「第 N 章 章名」，小节用页面的真编号
  assert.equal(sectionHeadingText(chapters[0]), '1.1 如何认识新方位');
  assert.equal(groupHeadingText(chapters[2].group, 1), '第 2 章 领悟人生真谛 把握人生方向');
});

test('扁平目录：分组标题与小节平铺成兄弟节点时，仍按文档顺序归属', () => {
  const root = createDocument(`
    <div class="posCatalog_list">
      <div class="posCatalog_select firstLayer" id="g1"><span class="posCatalog_title" title="第一章"><em class="posCatalog_sbar">1</em> 第一章</span></div>
      <div class="posCatalog_select" id="c1"><span class="posCatalog_name" title="小节 A" onclick="go('a');"><em class="posCatalog_sbar">1.1</em> 小节 A</span></div>
      <div class="posCatalog_select firstLayer" id="g2"><span class="posCatalog_title" title="第二章"><em class="posCatalog_sbar">2</em> 第二章</span></div>
      <div class="posCatalog_select" id="c2"><span class="posCatalog_name" title="小节 B" onclick="go('b');"><em class="posCatalog_sbar">2.1</em> 小节 B</span></div>
    </div>`);

  const chapters = new ChapterLocator(root).list();
  assert.deepEqual(
    chapters.map((chapter) => chapter.group?.id),
    ['g1', 'g2'],
  );
  assert.deepEqual(
    chapters.map((chapter) => chapter.label),
    ['1.1', '2.1'],
  );
});

test('老结构：名字里自带编号（title="1.1 小节"）时不再叠加编号', () => {
  const root = createDocument(`
    <div class="chapter-list">
      <div class="posCatalog_select" id="c1"><span class="posCatalog_name" title="1.1 小节" onclick="go('a');">1.1 小节</span></div>
      <div class="posCatalog_select" id="c2"><span class="posCatalog_name" title="1.2 小节" onclick="go('b');">1.2 小节</span></div>
    </div>`);

  const chapters = new ChapterLocator(root).list();
  // 扁平目录：没有章，只有节
  assert.deepEqual(
    chapters.map((chapter) => chapter.group),
    [null, null],
  );
  // 编号从名字前缀识别，显示名仍是原来的「1.1 小节」，不会变成「1.1 1.1 小节」
  assert.deepEqual(
    chapters.map((chapter) => chapter.label),
    ['1.1', '1.2'],
  );
  assert.deepEqual(chapters.map(sectionHeadingText), ['1.1 小节', '1.2 小节']);
});
