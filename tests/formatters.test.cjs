// 格式化单元测试：断言与主脚本 chaoxing-work-export.user.js 的输出保持一致
const test = require('node:test');
const assert = require('node:assert/strict');
const {
  formatOutputWithAnswers,
  formatWrongQuestionsTXT,
} = require('../.tmp/test/src/exporters/text-formatter.js');
const {
  formatOutputWithAnswersMD,
  formatWrongQuestionsMD,
} = require('../.tmp/test/src/exporters/markdown-formatter.js');
const {
  normalizeRichContent,
  shuffleQuestions,
} = require('../.tmp/test/src/exporters/legacy-bridge.js');

// 原版结构的题目数据
const legacyQuestion = {
  stem: '选择正确图片',
  stemContent: [
    { type: 'text', text: '选择正确图片' },
    { type: 'image', url: 'https://example.com/a(1).png', alt: '示意图' },
  ],
  options: [
    { letter: 'A', content: [{ type: 'text', text: '选项 A' }] },
    { letter: 'B', content: [{ type: 'text', text: '选项 B' }] },
  ],
  correctAnswer: 'A',
  correctAnswerContent: [{ type: 'text', text: 'A' }],
  myAnswer: 'B',
  isWrong: true,
};

const results = { 单选: [legacyQuestion], 多选: [], 填空: [], 判断: [], 简答: [] };
const typeOrder = ['单选'];

test('TXT 格式化保留图片、答案汇总与错题汇总（与主脚本一致）', () => {
  const output = formatOutputWithAnswers(results, typeOrder);
  // 原版行为：题干与选项各输出两行
  assert.match(output, /1\. 选择正确图片/);
  assert.match(output, /\[图片: https:\/\/example\.com\/a\(1\)\.png\]/);
  assert.match(output, /A\. 选项 A/);
  assert.match(output, /答案汇总/);
  assert.match(output, /一、单选题/);
  assert.match(output, /1\. A/);
});

test('TXT 错题汇总输出我的答案与正确答案（与主脚本一致）', () => {
  const output = formatWrongQuestionsTXT(results, typeOrder);
  assert.match(output, /错题汇总/);
  assert.match(output, /1\. \(题目\)选择正确图片/);
  assert.match(output, /我的答案: B/);
  assert.match(output, /正确答案: A/);
});

test('Markdown 格式化转义图片 URL 圆括号（与主脚本一致）', () => {
  const output = formatOutputWithAnswersMD(results, typeOrder);
  assert.match(output, /!\[示意图\]\(https:\/\/example\.com\/a%281%29\.png\)/);
  assert.match(output, /### 一、单选题（共1题）/);
  assert.match(output, /## 答案汇总/);
});

test('Markdown 错题汇总输出我的答案与正确答案（与主脚本一致）', () => {
  const output = formatWrongQuestionsMD(results, typeOrder);
  assert.match(output, /## 错题汇总/);
  assert.match(output, /- 我的答案: B/);
  assert.match(output, /- 正确答案: A/);
});

test('shuffleQuestions 同类型打乱且保持题目数量', () => {
  const many = {
    单选: [legacyQuestion, { ...legacyQuestion, correctAnswer: 'C' }, { ...legacyQuestion, correctAnswer: 'D' }],
    多选: [],
    填空: [],
    判断: [],
    简答: [],
  };
  const shuffled = shuffleQuestions(many, ['单选']);
  assert.equal(shuffled.单选.length, 3);
  // 原始数据不被修改
  assert.equal(many.单选[0].correctAnswer, 'A');
  // 打乱后仍包含全部答案
  const answers = shuffled.单选.map((q) => q.correctAnswer).sort();
  assert.deepEqual(answers, ['A', 'C', 'D']);
});

test('legacy normalizeRichContent 修剪尾部空白挡住的换行（与主脚本修复一致）', () => {
  const result = normalizeRichContent([
    { type: 'text', text: '题干' },
    { type: 'break' },
    { type: 'text', text: '\n ' },
  ]);
  // legacy 版本 pushText 会显式带上 bold/italic 等字段，故逐字段断言
  assert.equal(result.length, 1);
  assert.equal(result[0].type, 'text');
  assert.equal(result[0].text, '题干');
});
