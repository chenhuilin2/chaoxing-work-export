const test = require('node:test');
const assert = require('node:assert/strict');
const { createDocument, renderQuestion, EXPECTED_STEM } = require('./dom-setup.cjs');
const { resolveQuestionStem } = require('../.tmp/test/src/extractors/question-stem.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

/** 取到某题的 (容器, 标题容器) 后调用分层定位 */
function resolveFrom(document, containerSelector = '.TiMu', titleSelector = '.Zy_TItle') {
  const container = document.querySelector(containerSelector);
  const title = titleSelector ? container.querySelector(titleSelector) : null;
  return resolveQuestionStem(container, title);
}

test('2026 版答题页：题干容器 .Zy_TItle > div.fontLabel 被正确识别并剥离题号与题型标记', () => {
  const document = createDocument(renderQuestion());
  const stem = resolveFrom(document);

  assert.equal(richContentToText(stem.content), EXPECTED_STEM);
  assert.ok(!richContentToText(stem.content).includes('【单选题】'), '题型标记不应留在题干中');
  assert.ok(!/^\d/.test(richContentToText(stem.content)), '题号不应留在题干中');
  // 原始文本保留题号与题型标记，供 parseLeadingNumber 等旧逻辑使用
  assert.equal(stem.rawText, `【单选题】${EXPECTED_STEM}`);
});

test('学习通改版：题干类名被换掉时，由「题型标记的父元素」兜底', () => {
  const document = createDocument(
    renderQuestion({ stemContainerOpen: '<div class="clearfix font-cxsecret cwe-unknown-stem">' }),
  );
  const stem = resolveFrom(document);

  assert.equal(richContentToText(stem.content), EXPECTED_STEM);
});

test('学习通改版：题干类名与题型标记同时被换掉时，剔除题号后兜底取题干', () => {
  const document = createDocument(
    renderQuestion({ stemContainerOpen: '<div class="clearfix cwe-unknown-stem">', marker: '' }),
  );
  const stem = resolveFrom(document);

  assert.equal(richContentToText(stem.content), EXPECTED_STEM);
});

test('历史版本 .qtContent 仍走第一层，行为不变且 rawText 保留题号', () => {
  const document = createDocument(`
    <div class="TiMu newTiMu">
      <div class="Zy_TItle">
        <div class="qtContent">12. 旧版题干文本</div>
      </div>
      <ul class="Zy_ulTop"><li>A. 甲</li><li>B. 乙</li></ul>
    </div>`);
  const stem = resolveFrom(document);

  assert.equal(richContentToText(stem.content), '旧版题干文本');
  assert.equal(stem.rawText, '12. 旧版题干文本');
});

test('兜底路径不会把选项文本混进题干', () => {
  const document = createDocument(`
    <div class="TiMu newTiMu" typeName="单选题">
      <i class="fl">7</i>
      <div class="cwe-unknown-stem">信息安全管理是一个怎样的过程？</div>
      <ul class="Zy_ulTop"><li>A. 选项甲</li><li>B. 选项乙</li></ul>
    </div>`);
  const stem = resolveFrom(document, '.TiMu', null);
  const text = richContentToText(stem.content);

  assert.equal(text, '信息安全管理是一个怎样的过程？');
  assert.ok(!text.includes('选项甲'), '选项不应进入题干');
  assert.ok(!text.includes('选项乙'), '选项不应进入题干');
});

test('只有图片的题干不会被判为空', () => {
  const document = createDocument(`
    <div class="TiMu newTiMu">
      <div class="Zy_TItle">
        <div class="fontLabel"><img src="https://example.com/q.png" alt="题干图"></div>
      </div>
      <ul class="Zy_ulTop"><li>A. 甲</li></ul>
    </div>`);
  const stem = resolveFrom(document);

  assert.equal(stem.content.length, 1);
  assert.equal(stem.content[0].type, 'image');
  assert.equal(stem.content[0].url, 'https://example.com/q.png');
});

test('没有任何题干线索时返回空结果而不抛错', () => {
  const document = createDocument('<div class="TiMu newTiMu"><div class="Zy_TItle"></div></div>');
  const stem = resolveFrom(document);

  assert.deepEqual(stem.content, []);
  assert.equal(stem.rawText, '');
});
