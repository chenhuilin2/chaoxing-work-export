const test = require('node:test');
const assert = require('node:assert/strict');
const { createDocument, renderQuestion, EXPECTED_STEM } = require('./dom-setup.cjs');
const { TiMuExtractor } = require('../.tmp/test/src/extractors/timu-extractor.js');
const { CompositeExtractor } = require('../.tmp/test/src/extractors/composite-extractor.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

function context(document) {
  return {
    root: document,
    title: '做作业',
    pageUrl: 'https://mooc1.chaoxing.com/mooc-ans/api/work?api=1',
  };
}

test('2026 版答题页：TiMuExtractor 能提取到题目（此前因题干为空整题被丢弃）', () => {
  const document = createDocument(renderQuestion());
  const extractor = new TiMuExtractor();

  assert.equal(extractor.supports(document), true);
  const questions = extractor.extract(context(document));

  assert.equal(questions.length, 1);
  const question = questions[0];
  assert.equal(question.type, 'single-choice');
  assert.equal(question.typeMeta, '【单选题】');
  assert.equal(richContentToText(question.stem), EXPECTED_STEM);
});

test('2026 版答题页：选项按 A/B/C/D 解析出字母与正文', () => {
  const document = createDocument(renderQuestion());
  const question = new TiMuExtractor().extract(context(document))[0];

  assert.deepEqual(
    question.options.map((option) => option.key),
    ['A', 'B', 'C', 'D'],
  );
  assert.deepEqual(
    question.options.map((option) => richContentToText(option.content)),
    ['毛泽东', '李大钊', '陈独秀', '周恩来'],
  );
});

test('2026 版答题页：CompositeExtractor 端到端不再返回 null', () => {
  const document = createDocument(renderQuestion());
  const result = new CompositeExtractor().extract(document);

  assert.notEqual(result, null);
  assert.equal(result.extractor, 'timu');
  assert.equal(result.questions.length, 1);
  assert.equal(result.statistics.total, 1);
});

test('多题页面每道题都能提取到题干', () => {
  const document = createDocument(renderQuestion() + renderQuestion() + renderQuestion());
  const questions = new TiMuExtractor().extract(context(document));

  assert.equal(questions.length, 3);
  for (const question of questions) {
    assert.equal(richContentToText(question.stem), EXPECTED_STEM);
  }
});

test('旧版 .qtContent 页面行为不变', () => {
  const document = createDocument(`
    <div class="TiMu newTiMu">
      <div class="Zy_TItle">
        <div class="newZy_TItle">【多选题】</div>
        <div class="qtContent">1. 旧版多选题题干（　）</div>
      </div>
      <ul class="Zy_ulTop">
        <li>A. 甲</li><li>B. 乙</li><li>C. 丙</li>
      </ul>
    </div>`);
  const questions = new TiMuExtractor().extract(context(document));

  assert.equal(questions.length, 1);
  assert.equal(questions[0].type, 'multiple-choice');
  assert.equal(richContentToText(questions[0].stem), '旧版多选题题干（　）');
  assert.deepEqual(
    questions[0].options.map((option) => option.key),
    ['A', 'B', 'C'],
  );
});

test('学生学习页面：CompositeExtractor 的标题取自 #prev_title 的 title 属性', () => {
  const document = createDocument(
    `<div class="prev_list"><div id="prev_title" title="第一章 章节测验">章节小测</div></div>` +
      renderQuestion(),
  );
  const result = new CompositeExtractor().extract(document);

  assert.notEqual(result, null);
  assert.equal(result.title, '第一章 章节测验');
  assert.equal(result.questions.length, 1);
});

test('学生学习页面：#prev_title 只有正文时，标题取正文而非 document.title', () => {
  const document = createDocument(`<div id="prev_title">第二章 章节练习</div>` + renderQuestion());
  const result = new CompositeExtractor().extract(document);

  assert.notEqual(result, null);
  assert.equal(result.title, '第二章 章节练习');
});

test('学生学习页面：真实类名 .prev_title 的章节名优先于题目文档自身标题', () => {
  const document = createDocument(
    `<div class="prev_title_pos"><div class="prev_title" title="毛泽东思想的主要内容">毛泽东思想的主要内容<span class="markTag"></span></div></div>` +
      renderQuestion(),
  );
  const result = new CompositeExtractor().extract(document);

  assert.notEqual(result, null);
  assert.equal(result.title, '毛泽东思想的主要内容');
});

test('题干为空且无任何线索时不产出无效题目', () => {
  const document = createDocument(`
    <div class="TiMu newTiMu" typeName="单选题">
      <div class="Zy_TItle"></div>
      <ul class="Zy_ulTop"><li>A. 甲</li></ul>
    </div>`);
  const questions = new TiMuExtractor().extract(context(document));

  assert.equal(questions.length, 0);
});
