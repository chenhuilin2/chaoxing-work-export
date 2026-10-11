// 章节切换检测用例：学习通切章走 getTeacherAjax()（重建 #mainid、移动激活类），
// URL 不变、页面不刷新，只能靠轮询「当前激活章节的标识」来发现。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument, useGlobalDocument } = require('./dom-setup.cjs');
const { ChapterLocator } = require('../.tmp/test/src/application/chapter-locator.js');
const { ChapterChangeWatcher } = require('../.tmp/test/src/application/chapter-change-watcher.js');
const { ChapterExtractionService } = require('../.tmp/test/src/application/chapter-extraction-service.js');

/** 与 html/章节练习完整页.html 一致的目录树；容器 id 的真实形态是 `cur<chapterId>` */
const CATALOG_HTML = `
<div id="coursetree">
  <div class="posCatalog_list">
    <ul>
      <li>
        <div class="posCatalog_select firstLayer" id="1200668675">
          <span class="posCatalog_title posCatalog_rotate" title="导论"><em class="posCatalog_sbar">1</em> 导论</span>
        </div>
        <div class="posCatalog_level" style="display: block;">
          <ul>
            <li>
              <div class="posCatalog_select posCatalog_active" id="cur1200668730">
                <span class="posCatalog_name" title="导论1" onclick="getTeacherAjax('265608653','151187867','1200668730');"><em class="posCatalog_sbar">1.1</em> 导论1</span>
              </div>
            </li>
            <li>
              <div class="posCatalog_select" id="cur1200668731">
                <span class="posCatalog_name" title="导论2" onclick="getTeacherAjax('265608653','151187867','1200668731');"><em class="posCatalog_sbar">1.2</em> 导论2</span>
              </div>
            </li>
          </ul>
        </div>
      </li>
    </ul>
  </div>
</div>`;

/** 模拟 jQuery 里 `$('.posCatalog_select').removeClass(...)` + `$('#cur'+id).addClass(...)` */
function activateChapter(root, chapterId) {
  for (const node of root.querySelectorAll('.posCatalog_select')) {
    node.classList.remove('posCatalog_active');
  }
  root.getElementById(`cur${chapterId}`).classList.add('posCatalog_active');
}

/** 造一个「目录树 + 可切换的 watcher」组合 */
function createScenario() {
  const root = createDocument(CATALOG_HTML);
  useGlobalDocument(root);
  const locator = new ChapterLocator(root);
  const watcher = new ChapterChangeWatcher(() => locator.activeKey());
  watcher.reset();
  return { root, locator, watcher };
}

test('读取当前章节标识：容器 id 即 `cur<chapterId>`', () => {
  const root = createDocument(CATALOG_HTML);
  assert.equal(new ChapterLocator(root).activeKey(), 'cur1200668730');
});

test('目录树没有激活章节时返回 null，不用顺序编号硬凑一个', () => {
  const root = createDocument(CATALOG_HTML.replace('posCatalog_active', ''));
  assert.equal(new ChapterLocator(root).activeKey(), null);
});

test('首次同步只建立基线，不把「目录树刚渲染出来」当成切章', () => {
  const watcher = new ChapterChangeWatcher(() => 'cur1200668730');
  assert.equal(watcher.poll(), false, '第一次读到标识只记基线');
  assert.equal(watcher.poll(), false, '同一章节反复轮询不上报');
});

test('切到别的章节后上报一次，之后继续轮询不再重复上报', () => {
  const { root, watcher } = createScenario();

  assert.equal(watcher.poll(), false);
  activateChapter(root, '1200668731');
  assert.equal(watcher.poll(), true, '激活类移到新章节 → 上报一次');
  assert.equal(watcher.poll(), false, '停留在同一章节不再上报');
});

test('读不到章节标识时既不上报也不污染基线（目录重建的瞬间）', () => {
  let key = 'cur1200668730';
  const watcher = new ChapterChangeWatcher(() => key);
  watcher.reset();

  key = null;
  assert.equal(watcher.poll(), false, '读不到就什么都不做');
  key = 'cur1200668730';
  assert.equal(watcher.poll(), false, '短暂读空后回到原章节，不算切章');

  key = 'cur1200668731';
  assert.equal(watcher.poll(), true, '真正切走之后仍然能上报');
});

test('读取章节标识抛异常时不崩溃、不上报', () => {
  const watcher = new ChapterChangeWatcher(() => {
    throw new Error('宿主页面结构异常');
  });
  watcher.reset();
  assert.equal(watcher.poll(), false);
  assert.equal(watcher.currentKey(), null);
});

test('reset 后同一章节不再上报（批量提取逐章遍历后重新对准基线）', () => {
  const { root, watcher } = createScenario();

  activateChapter(root, '1200668731');
  assert.equal(watcher.poll(), true);
  watcher.reset();
  assert.equal(watcher.poll(), false, 'reset 把当前章节记成基线');
});

test('章节提取服务透出当前章节标识，供轮询使用', () => {
  const root = createDocument(CATALOG_HTML);
  useGlobalDocument(root);
  const service = new ChapterExtractionService(null, new ChapterLocator(root));
  assert.equal(service.activeChapterKey(), 'cur1200668730');
});
