// 跨文档（同源 iframe）提取用例。
//
// 真机现场「已提取 N 道题，其中 M 道选择题没提取到选项」的成因：
// 脚本在顶层窗口运行时会把同源 iframe 的文档一并遍历（collectAccessibleDocuments），
// 而 iframe 里节点的构造器属于它们自己的 window，`node instanceof Element` 一律判 false，
// 于是 extractRichContent 进不去任何子元素 —— 题干恰好落在容器的直接文本节点上还读得到，
// 选项文字都嵌在 <a> / <label> / <span> 里就一条都读不到。
//
// 单文档离线快照（linkedom）复现不出这种差异，因此这里用 withForeignConstructors
// 把全局构造器换成无关的类来还原跨 window 的判定结果。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument, renderQuestion, withForeignConstructors } = require('./dom-setup.cjs');
const {
  extractBackgroundImages,
  extractRichContent,
  richContentToText,
} = require('../.tmp/test/src/extractors/rich-content.js');
const { CompositeExtractor } = require('../.tmp/test/src/extractors/composite-extractor.js');

test('节点来自另一个 window 时，嵌套元素里的文本仍能读出', () => {
  const document = createDocument(renderQuestion());
  // 选项字母嵌在 <label><span> 里、选项文字嵌在 <a> 里，跨 window 时两者都必须读得到
  // （模板里标签之间的换行会保留成空白，故折叠空白后比对）
  const text = withForeignConstructors(() =>
    richContentToText(extractRichContent(document.querySelector('.Zy_ulTop > li'))),
  );

  assert.equal(text.replace(/\s+/gu, ''), 'A毛泽东');
});

test('节点来自另一个 window 时，选项仍能完整提取（真题干、真选项）', () => {
  const document = createDocument(renderQuestion());
  const result = withForeignConstructors(() => new CompositeExtractor().extract(document));

  assert.equal(result?.questions.length, 1);
  assert.deepEqual(
    result.questions[0].options.map(
      (option) => `${option.key}.${richContentToText(option.content)}`,
    ),
    ['A.毛泽东', 'B.李大钊', 'C.陈独秀', 'D.周恩来'],
  );
});

test('节点来自另一个 window 时，题干里嵌套的题型标记同样不丢', () => {
  // 真机报告里题干显示为「（）标志着…」而丢了 <span class="newZy_TItle">【单选题】</span>，
  // 正是「进不去子元素」的旁证；这里守住嵌套文本一律可见
  const document = createDocument(renderQuestion());
  const stem = document.querySelector('.fontLabel');

  assert.equal(
    withForeignConstructors(() => richContentToText(extractRichContent(stem))),
    '【单选题】在中国共产党历史上，第一个明确提出了“马克思主义中国化”的科学命题和重大任务的是（）。',
  );
});

test('节点来自另一个 window 时，背景图识别结果与同 realm 一致', () => {
  const document = createDocument(
    '<div id="x" style="background-image: url(https://example.com/a.png)">题干</div>',
  );
  const element = document.querySelector('#x');
  const native = extractBackgroundImages(element);
  const foreign = withForeignConstructors(() => extractBackgroundImages(element));

  assert.deepEqual(foreign, native);
});
