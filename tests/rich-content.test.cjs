const test = require('node:test');
const assert = require('node:assert/strict');
const {
  normalizeRichContent,
  richContentToText,
  stripAnswerLabel,
  stripOptionPrefix,
  stripQuestionPrefix,
} = require('../.tmp/test/src/extractors/rich-content.js');

test('normalizeRichContent merges adjacent text with identical styles', () => {
  const result = normalizeRichContent([
    { type: 'break' },
    { type: 'text', text: ' 题' },
    { type: 'text', text: '干 ' },
    { type: 'break' },
    { type: 'break' },
  ]);
  assert.deepEqual(result, [
    { type: 'text', text: ' 题干 ' },
  ]);
});

test('normalizeRichContent 修剪尾部被空白文本挡住的换行（Word 空行修复）', () => {
  // 学习通题干 HTML 末尾常残留：换行 + 纯空白文本，导出 Word 会在题目与选项间多出空行
  const result = normalizeRichContent([
    { type: 'text', text: '信息安全管理是一个怎样的过程？' },
    { type: 'break' },
    { type: 'text', text: '\n ' },
  ]);
  assert.deepEqual(result, [
    { type: 'text', text: '信息安全管理是一个怎样的过程？' },
  ]);
});

test('normalizeRichContent 修剪首部空白，但保留图片前的换行', () => {
  // 首部空白 + 换行应被修剪
  const lead = normalizeRichContent([
    { type: 'text', text: ' ' },
    { type: 'break' },
    { type: 'text', text: '题干' },
  ]);
  assert.deepEqual(lead, [{ type: 'text', text: '题干' }]);

  // 尾部图片不是空白，其前面的换行必须保留（图片另起一行）
  const tail = normalizeRichContent([
    { type: 'text', text: '题干' },
    { type: 'break' },
    { type: 'image', url: 'https://example.com/a.png', alt: '' },
  ]);
  assert.deepEqual(tail, [
    { type: 'text', text: '题干' },
    { type: 'break' },
    { type: 'image', url: 'https://example.com/a.png', alt: '' },
  ]);
});

test('prefix strippers preserve normalized rich content', () => {
  const question = stripQuestionPrefix([{ type: 'text', text: '12.【单选题】 TCP 属于哪一层？' }]);
  assert.equal(richContentToText(question), 'TCP 属于哪一层？');

  const option = stripOptionPrefix([{ type: 'text', text: 'A. 传输层' }]);
  assert.equal(option.key, 'A');
  assert.equal(richContentToText(option.content), '传输层');

  const answer = stripAnswerLabel([{ type: 'text', text: '正确答案：A' }]);
  assert.equal(richContentToText(answer), 'A');
});
