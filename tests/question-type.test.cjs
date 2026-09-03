const test = require('node:test');
const assert = require('node:assert/strict');
const {
  detectQuestionType,
} = require('../.tmp/test/src/extractors/question-type.js');
const { normalizeAnswer } = require('../.tmp/test/src/utils/text.js');

test('detectQuestionType maps common Chinese labels and legacy codes', () => {
  assert.equal(detectQuestionType('单项选择题'), 'single-choice');
  assert.equal(detectQuestionType('【多选题】'), 'multiple-choice');
  assert.equal(detectQuestionType('填空'), 'fill-blank');
  assert.equal(detectQuestionType('判断题'), 'true-false');
  assert.equal(detectQuestionType('简答题'), 'short-answer');
  assert.equal(detectQuestionType('0'), 'single-choice');
});

test('normalizeAnswer handles choice separators and true-false synonyms', () => {
  assert.equal(normalizeAnswer('正确答案：A，C'), 'AC');
  assert.equal(normalizeAnswer(' 对 '), 'TRUE');
  assert.equal(normalizeAnswer('×'), 'FALSE');
});
