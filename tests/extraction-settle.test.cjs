// 提取编排用例：
// 1. 学生学习页面的章节名在外层壳（.prev_title），题目在嵌套 iframe 里，标题必须由外层壳提供；
// 2. 采样等待要在题目就绪时立即返回，不再固定等满整段超时；
// 3. 章节切换后卡片页会重建并回到「视频」卡，采样循环要自动补切任务卡并尽快拿到题目。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument, renderQuestion, useGlobalDocument, EXPECTED_STEM } = require('./dom-setup.cjs');
const { ExtractionService } = require('../.tmp/test/src/application/extraction-service.js');
const {
  ChapterExtractionService,
} = require('../.tmp/test/src/application/chapter-extraction-service.js');
const { ChapterLocator } = require('../.tmp/test/src/application/chapter-locator.js');
const { TaskTabLocator } = require('../.tmp/test/src/application/task-tab-locator.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

/** 与 html/章节测验.html 一致的目录树：一个分组标题 + 两个章节 */
const CATALOG_HTML = `
<div id="coursetree">
  <div class="posCatalog_list">
    <ul>
      <li>
        <div class="posCatalog_select firstLayer" id="1200668675">
          <span class="posCatalog_title titleIcon" title="导论 马克思主义中国化时代化的历史进程与理论成果"><em class="posCatalog_sbar">1</em> 导论</span>
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

/** 任务卡栏：默认停在「视频」卡（题目 iframe 尚未加载） */
const TAB_BAR_HTML = `
<div class="prev_list" id="prev_tab">
  <ul class="prev_ul clearfix">
    <li class="active" title="视频" role="option" id="dct1" onclick="changeDisplayContent(1,2,'1200668686','265608653','151187867','');">
      <div class="prev_white" aria-hidden="true"><span class="num">1</span><span class="spanText">视频</span></div>
    </li>
    <li class="c2" title="章节测验" role="option" id="dct2" onclick="changeDisplayContent(2,2,'1200668686','265608653','151187867','');">
      <div class="prev_white" aria-hidden="true"><span class="num">2</span><span class="spanText">章节测验</span></div>
    </li>
  </ul>
</div>`;

/** 第二章的答题页内容（题干不同 → 题目指纹不同） */
const SECOND_CHAPTER_HTML = renderQuestion().replace(
  '第一个明确提出了“马克思主义中国化”',
  '第一个明确提出“农村包围城市”',
);

const PAGE_REBUILD_MS = 200;

/** 用户侧看到的章节页：外层壳（标题 + 目录树 + 任务卡栏）+ 答题 iframe */
function buildStudyPage() {
  const firstAnswer = createDocument(renderQuestion());
  const secondAnswer = createDocument(SECOND_CHAPTER_HTML);
  // 卡片页重建后默认展示视频卡：此时答题页尚未加载，页面里没有题目
  const videoCard = createDocument('<div class="video-card">视频播放中</div>');

  const outer = createDocument(`
    <div class="prev_title_pos"><div class="prev_title" title="导论1">导论1</div></div>
    ${CATALOG_HTML}
    ${TAB_BAR_HTML}
    <iframe id="frame"></iframe>`);

  let current = firstAnswer;
  const frame = outer.getElementById('frame');
  Object.defineProperty(frame, 'contentDocument', { get: () => current, configurable: true });

  const quizTab = outer.getElementById('dct2');
  const videoTab = outer.getElementById('dct1');
  let quizClicks = 0;
  quizTab.addEventListener('click', () => {
    quizClicks += 1;
    quizTab.classList.add('active');
    videoTab.classList.remove('active');
    // 页面自己的 changeDisplayContent()：只有卡片页已重建（视频占位页就位）时才会加载答题页
    if (current === videoCard) current = secondAnswer;
  });

  // 切换章节：卡片页重建，任务卡栏回到默认的「视频」卡
  outer.getElementById('1200670000').addEventListener('click', () => {
    window.setTimeout(() => {
      current = videoCard;
      quizTab.classList.remove('active');
      videoTab.classList.add('active');
    }, PAGE_REBUILD_MS);
  });

  return { outer, quizClicks: () => quizClicks };
}

function buildChapterService(outer) {
  return new ChapterExtractionService(
    new ExtractionService(),
    new ChapterLocator(outer),
    new TaskTabLocator(outer),
  );
}

test('题目在嵌套 iframe 里时，标题取外层壳 .prev_title 的章节名', () => {
  const answer = createDocument(renderQuestion());
  const outer = createDocument(`
    <div class="prev_title_pos"><div class="prev_title" title="毛泽东思想的主要内容">毛泽东思想的主要内容<span class="markTag"></span></div></div>
    <iframe id="frame"></iframe>`);
  Object.defineProperty(outer.getElementById('frame'), 'contentDocument', {
    get: () => answer,
    configurable: true,
  });
  useGlobalDocument(outer);

  const result = new ExtractionService().extract();

  assert.notEqual(result, null);
  assert.equal(result.questions.length, 1);
  assert.equal(richContentToText(result.questions[0].stem), EXPECTED_STEM);
  // 答题页自身只能解析出「做作业」，章节名必须来自外层壳
  assert.equal(result.title, '毛泽东思想的主要内容');
});

test('外层壳没有 .prev_title 时，标题仍取题目所在文档自己的标题', () => {
  const answer = createDocument(renderQuestion());
  const outer = createDocument('<iframe id="frame"></iframe>');
  Object.defineProperty(outer.getElementById('frame'), 'contentDocument', {
    get: () => answer,
    configurable: true,
  });
  useGlobalDocument(outer);

  const result = new ExtractionService().extract();
  assert.equal(result.title, '做作业');
});

test('题目已就绪时采样立即返回，不等满超时', async () => {
  useGlobalDocument(createDocument(renderQuestion()));
  const startedAt = Date.now();
  const result = await new ExtractionService().settleQuestions({ timeoutMs: 5_000 });
  const elapsed = Date.now() - startedAt;

  assert.notEqual(result, null);
  assert.equal(result.questions.length, 1);
  assert.ok(elapsed < 300, `应立即返回，实际耗时 ${elapsed}ms`);
});

test('页面没有题目时按超时上限结束，不拖长尾', async () => {
  useGlobalDocument(createDocument('<div class="video-card">视频播放中</div>'));
  const startedAt = Date.now();
  const result = await new ExtractionService().settleQuestions({ timeoutMs: 200 });
  const elapsed = Date.now() - startedAt;

  assert.equal(result, null);
  assert.ok(elapsed >= 180 && elapsed < 900, `应在超时附近结束，实际耗时 ${elapsed}ms`);
});

test('批量提取：章节切换后自动补切任务卡并尽快拿到题目', async () => {
  const page = buildStudyPage();
  useGlobalDocument(page.outer);

  const startedAt = Date.now();
  const result = await buildChapterService(page.outer).extractSelected([1]);
  const elapsed = Date.now() - startedAt;

  assert.equal(result.chapters.length, 1);
  assert.equal(result.chapters[0].title, '导论2');
  assert.equal(result.chapters[0].questions.length, 1);
  assert.match(richContentToText(result.questions[0].stem), /农村包围城市/);
  // 首次点击落在尚未重建的任务卡栏上（无效），页面重建后必须再补切一次
  assert.ok(page.quizClicks() >= 2, `补切次数 ${page.quizClicks()}`);
  assert.ok(page.quizClicks() <= 3, `补切次数 ${page.quizClicks()}`);
  // 旧实现只切一次卡再等待，此时会一路空等到 20s 超时
  assert.ok(elapsed < 2_000, `单章应秒级完成，实际耗时 ${elapsed}ms`);
});

test('批量提取：章节已激活时不空等，直接采用当前题目', async () => {
  const page = buildStudyPage();
  useGlobalDocument(page.outer);

  const startedAt = Date.now();
  const result = await buildChapterService(page.outer).extractSelected([0]);
  const elapsed = Date.now() - startedAt;

  assert.equal(result.chapters[0].title, '导论1');
  assert.equal(result.chapters[0].questions.length, 1);
  assert.ok(elapsed < 1_000, `应立即返回，实际耗时 ${elapsed}ms`);
});
