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

test('prefix strippers preserve normalized rich content', () => {
  const question = stripQuestionPrefix([{ type: 'text', text: '12.【单选题】 TCP 属于哪一层？' }]);
  assert.equal(richContentToText(question), 'TCP 属于哪一层？');

  const option = stripOptionPrefix([{ type: 'text', text: 'A. 传输层' }]);
  assert.equal(option.key, 'A');
  assert.equal(richContentToText(option.content), '传输层');

  const answer = stripAnswerLabel([{ type: 'text', text: '正确答案：A' }]);
  assert.equal(richContentToText(answer), 'A');
});
