const test = require('node:test');
const assert = require('node:assert/strict');
const { buildStatistics, deriveTypeOrder } = require('../.tmp/test/src/domain/question.js');
const { formatText } = require('../.tmp/test/src/exporters/text-formatter.js');
const { formatMarkdown } = require('../.tmp/test/src/exporters/markdown-formatter.js');

const questions = [
  {
    id: 'q1',
    number: 1,
    type: 'single-choice',
    stem: [
      { type: 'text', text: '选择正确图片' },
      { type: 'image', url: 'https://example.com/a(1).png', alt: '示意图' },
    ],
    options: [
      { key: 'A', content: [{ type: 'text', text: '选项 A' }] },
      { key: 'B', content: [{ type: 'text', text: '选项 B' }] },
    ],
    correctAnswer: [{ type: 'text', text: 'A' }],
    userAnswer: [{ type: 'text', text: 'B' }],
    analysis: [{ type: 'text', text: '因为 A 正确' }],
    isWrong: true,
    source: { extractor: 'fixture', pageUrl: 'https://example.com' },
  },
];

const result = {
  title: '格式化测试',
  questions,
  typeOrder: deriveTypeOrder(questions),
  statistics: buildStatistics(questions),
  sourceUrl: 'https://example.com',
  extractor: 'fixture',
  extractedAt: '2026-09-03T00:00:00.000Z',
};

const options = {
  format: 'txt',
  filename: 'test',
  withAnswers: true,
  withWrong: true,
  includeAnalysis: true,
  shuffle: false,
  bankImport: false,
  splitByChapter: false,
};

test('TXT formatter keeps images, answers, analysis and wrong-answer context', () => {
  const output = formatText(result, options);
  assert.match(output, /\[图片: https:\/\/example\.com\/a\(1\)\.png\]/);
  assert.match(output, /参考答案/);
  assert.match(output, /解析：因为 A 正确/);
  assert.match(output, /我的答案：B/);
});

test('Markdown formatter escapes parentheses in image URLs', () => {
  const output = formatMarkdown(result, { ...options, format: 'md' });
  assert.match(output, /!\[示意图\]\(https:\/\/example\.com\/a%281%29\.png\)/);
  assert.match(output, /# 错题汇总/);
});
