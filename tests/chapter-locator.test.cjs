// 章节定位器用例：覆盖学习通目录树的真实结构（span.posCatalog_name 才是可点章节，
// .firstLayer / .posCatalog_title 是分组标题）。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument } = require('./dom-setup.cjs');
const { ChapterLocator } = require('../.tmp/test/src/application/chapter-locator.js');

/** 与 html/章节测验.html 一致的目录树：一个分组标题 + 两个章节 */
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
