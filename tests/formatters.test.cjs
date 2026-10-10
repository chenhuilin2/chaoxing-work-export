// 格式化单元测试：断言与主脚本 chaoxing-work-export.user.js 的输出保持一致
const test = require('node:test');
const assert = require('node:assert/strict');
const {
  formatAnswersTXT,
  formatOutput,
  formatOutputWithAnswers,
  formatSectionsOutput,
  formatSectionsText,
  formatWrongQuestionsTXT,
} = require('../.tmp/test/src/exporters/text-formatter.js');
const {
  formatAnswersMD,
  formatOutputMD,
  formatOutputWithAnswersMD,
  formatSectionsMD,
  formatSectionsOutputMD,
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
  assert.match(output, /1\. 选择正确图片/);
  assert.match(output, /\[图片: https:\/\/example\.com\/a\(1\)\.png\]/);
  assert.match(output, /A\. 选项 A/);
  assert.match(output, /答案汇总/);
  assert.match(output, /一、单选题/);
  assert.match(output, /1\. A/);
});

/** 统计子串出现次数 */
function countOccurrences(text, fragment) {
  return text.split(fragment).length - 1;
}

test('题干与每个选项各只输出一次（不再连着重两遍）', () => {
  // 旧主脚本把题干行与选项行各 `output +=` 了两次，移植时按「保持输出一致」照搬，
  // 结果是导出的每道题都重复两遍；这里锁死「各出现一次」
  const text = formatOutputWithAnswers(results, typeOrder);
  assert.equal(countOccurrences(text, '1. 选择正确图片'), 1);
  assert.equal(countOccurrences(text, 'A. 选项 A'), 1);
  assert.equal(countOccurrences(text, 'B. 选项 B'), 1);

  const markdown = formatOutputWithAnswersMD(results, typeOrder);
  assert.equal(countOccurrences(markdown, '**1.** 选择正确图片'), 1);
  assert.equal(countOccurrences(markdown, '- A. 选项 A'), 1);

  const wrong = formatWrongQuestionsTXT(results, typeOrder);
  assert.equal(countOccurrences(wrong, '1. (题目)选择正确图片'), 1);
  assert.equal(countOccurrences(wrong, '我的答案: B'), 1);

  const wrongMarkdown = formatWrongQuestionsMD(results, typeOrder);
  assert.equal(countOccurrences(wrongMarkdown, '**1.** 选择正确图片'), 1);
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
    单选: [
      legacyQuestion,
      { ...legacyQuestion, correctAnswer: 'C' },
      { ...legacyQuestion, correctAnswer: 'D' },
    ],
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

// ==================== 多章节分段（与 Word 的章节分节对齐） ====================

/** 造一道原版结构的题目 */
function legacyQ(stem, { answer = 'A', myAnswer = 'B', isWrong = false } = {}) {
  return {
    stem,
    stemContent: [{ type: 'text', text: stem }],
    options: [
      { letter: 'A', content: [{ type: 'text', text: '选项 A' }] },
      { letter: 'B', content: [{ type: 'text', text: '选项 B' }] },
    ],
    correctAnswer: answer,
    correctAnswerContent: [{ type: 'text', text: answer }],
    myAnswer,
    isWrong,
  };
}

/** 两个章节：第一章 单选 + 多选都没做错，第二章只有单选且答错（两级标题：章 + 小节） */
const SECTION_ONE = {
  chapterHeading: '第 1 章 担当复兴大任 成就时代新人',
  heading: '1.1 如何认识新方位',
  results: {
    单选: [legacyQ('第一章第一题')],
    多选: [legacyQ('第一章多选')],
    填空: [],
    判断: [],
    简答: [],
  },
  typeOrder: ['单选', '多选'],
};
const SECTION_TWO = {
  chapterHeading: '第 2 章 领悟人生真谛 把握人生方向',
  heading: '2.1 如何认识新使命',
  results: {
    单选: [legacyQ('第二章第一题', { isWrong: true })],
    多选: [],
    填空: [],
    判断: [],
    简答: [],
  },
  typeOrder: ['单选'],
};

/** 老数据 / 扁平目录：没有章信息，只有一级标题（改动前的写法） */
const FLAT_SECTIONS = [
  {
    chapterHeading: '第 1 章 如何认识新方位',
    heading: null,
    results: SECTION_ONE.results,
    typeOrder: SECTION_ONE.typeOrder,
  },
  {
    chapterHeading: '第 2 章 如何认识新使命',
    heading: null,
    results: SECTION_TWO.results,
    typeOrder: SECTION_TWO.typeOrder,
  },
];

test('TXT 多章节：章标题与小节标题两级，换章多空两行、换小节多空一行', () => {
  const body = formatSectionsText([SECTION_ONE, SECTION_TWO]);

  // 逐字锁死整段正文：这是 TXT 导出的最终形态。
  // 章标题上方空 3 行（比原来多 2 行），小节标题上方空 2 行（比原来多 1 行）。
  assert.equal(
    body,
    [
      '第 1 章 担当复兴大任 成就时代新人',
      '',
      '1.1 如何认识新方位',
      '',
      '一. 单选题（共1题）',
      '1. 第一章第一题',
      'A. 选项 A',
      'B. 选项 B',
      '',
      '二. 多选题（共1题）',
      '2. 第一章多选',
      'A. 选项 A',
      'B. 选项 B',
      '',
      '',
      '',
      '第 2 章 领悟人生真谛 把握人生方向',
      '',
      '2.1 如何认识新使命',
      '',
      '一. 单选题（共1题）',
      '1. 第二章第一题',
      'A. 选项 A',
      'B. 选项 B',
    ].join('\n'),
  );
});

test('TXT 多章节：同一章下的多个小节，小节之间空 2 行、章标题不重复', () => {
  // 把第二小节也挂到第一章下：整篇只有一个章，章标题不插，两个小节标题各出现一次
  const siblings = [
    SECTION_ONE,
    { ...SECTION_TWO, chapterHeading: SECTION_ONE.chapterHeading, heading: '1.2 如何认识新使命' },
  ];
  const body = formatSectionsText(siblings);

  assert.equal(countOccurrences(body, '第 1 章 担当复兴大任 成就时代新人'), 0);
  assert.equal(countOccurrences(body, '1.1 如何认识新方位'), 1);
  assert.equal(countOccurrences(body, '1.2 如何认识新使命'), 1);
  // 上一个题型的最后一个选项 → 空 2 行 → 下一个小节标题
  assert.match(
    body,
    /^1\.1 如何认识新方位\n\n一\. 单选题[\s\S]*B\. 选项 B\n\n\n1\.2 如何认识新使命\n\n一\. 单选题/,
  );
});

test('TXT 多章节：答案与错题紧跟在所属章之后，不再堆在文件末尾', () => {
  const output = formatSectionsOutput([SECTION_ONE, SECTION_TWO], {
    withAnswers: true,
    withWrong: true,
  });

  // 章标题只在各自章的开头出现一次：答案与错题都在章内，末页不会再来一遍
  assert.equal(countOccurrences(output, '第 1 章 担当复兴大任 成就时代新人'), 1);
  assert.equal(countOccurrences(output, '第 2 章 领悟人生真谛 把握人生方向'), 1);

  // 顺序：章 1 题目 → 章 1 答案 → 章 2 题目 → 章 2 错题
  const anchors = [
    '第 1 章 担当复兴大任 成就时代新人',
    '答案汇总',
    '第 2 章 领悟人生真谛 把握人生方向',
    '错题汇总',
  ].map((fragment) => output.indexOf(fragment));
  assert.ok(
    anchors.every((index) => index >= 0),
    `片段缺失：${anchors.join(',')}`,
  );
  assert.deepEqual(
    anchors,
    [...anchors].sort((a, b) => a - b),
  );

  // 第一章一道错题都没有：连「错题汇总」页头一起省掉
  assert.equal(countOccurrences(output, '错题汇总'), 1);

  // 答案编号与正文题号一一对应：第一章 1、2 连排，第二章重新从 1 开始
  assert.match(output, /1\.1 如何认识新方位\n\n一、单选题\n\n1\. A\n\n二、多选题\n\n2\. A/);
  assert.match(output, /2\.1 如何认识新使命\n\n一、单选题\n\n1\. A/);
  assert.match(output, /2\.1 如何认识新使命\n\n1\. \(题目\)第二章第一题/);
  assert.match(output, /我的答案: B/);
  assert.match(output, /正确答案: A/);
});

test('老数据 / 扁平目录：没有章信息时退回单级「第 N 章 名字」，只多空两行', () => {
  const body = formatSectionsText(FLAT_SECTIONS);

  assert.equal(
    body,
    [
      '第 1 章 如何认识新方位',
      '',
      '一. 单选题（共1题）',
      '1. 第一章第一题',
      'A. 选项 A',
      'B. 选项 B',
      '',
      '二. 多选题（共1题）',
      '2. 第一章多选',
      'A. 选项 A',
      'B. 选项 B',
      '',
      '',
      '',
      '第 2 章 如何认识新使命',
      '',
      '一. 单选题（共1题）',
      '1. 第二章第一题',
      'A. 选项 A',
      'B. 选项 B',
    ].join('\n'),
  );

  // 没有小节标题时 Markdown 的题型头仍是 ###，不降级
  assert.match(
    formatSectionsMD(FLAT_SECTIONS),
    /^## 第 1 章 如何认识新方位\n\n### 一、单选题（共1题）/,
  );
});

test('Markdown 多章节：章用 ## 、小节用 ###，题型头降为 ####，答案与错题跟着章走', () => {
  const body = formatSectionsMD([SECTION_ONE, SECTION_TWO]);
  assert.match(
    body,
    /^## 第 1 章 担当复兴大任 成就时代新人\n\n### 1\.1 如何认识新方位\n\n#### 一、单选题（共1题）/,
  );
  assert.match(
    body,
    /## 第 2 章 领悟人生真谛 把握人生方向\n\n### 2\.1 如何认识新使命\n\n#### 一、单选题（共1题）\n\n\*\*1\.\*\* 第二章第一题/,
  );
  // 换章多空两行：章标题前共 3 个空行
  assert.match(body, /- B\. 选项 B\n\n\n\n## 第 2 章 领悟人生真谛 把握人生方向/);

  const output = formatSectionsOutputMD([SECTION_ONE, SECTION_TWO], {
    withAnswers: true,
    withWrong: true,
  });
  // 汇总页头落在章内，与章标题同级（##），后面跟本章的小节标题
  assert.match(output, /\n\n---\n\n## 答案汇总\n\n### 1\.1 如何认识新方位/);
  assert.match(output, /\n\n---\n\n## 错题汇总\n\n### 2\.1 如何认识新使命/);
  // 只有第二章有错题，因此「错题汇总」只出现一次
  assert.equal(countOccurrences(output, '## 错题汇总'), 1);
});

test('只有一个章节但带页面编号时，小节标题不丢（拆出的「1.3 本章测试」这类章）', () => {
  const one = [SECTION_TWO];

  // 单章：章标题不插（文件名已经是章名），小节标题必须保留
  const body = formatSectionsText(one);
  assert.equal(body.startsWith('2.1 如何认识新使命\n\n一. 单选题（共1题）'), true);
  assert.equal(countOccurrences(body, '第 2 章 领悟人生真谛 把握人生方向'), 0);

  const output = formatSectionsOutput(one, { withAnswers: true, withWrong: true });
  assert.match(output, /^2\.1 如何认识新使命\n\n一\. 单选题/);
  // 答案与错题区同样带小节标签，与多小节文件一致
  assert.match(output, /答案汇总[\s\S]*2\.1 如何认识新使命\n\n一、单选题\n\n1\. A/);
  assert.match(output, /错题汇总[\s\S]*2\.1 如何认识新使命\n\n1\. \(题目\)第二章第一题/);

  assert.match(formatSectionsMD(one), /^### 2\.1 如何认识新使命\n\n#### 一、单选题（共1题）/);
});

test('单章节：分段版输出与原函数逐字一致（改造不影响单章节结果）', () => {
  const one = [{ chapterHeading: null, heading: null, results, typeOrder }];

  assert.equal(formatSectionsText(one), formatOutput(results, typeOrder));

  // 与 export-service 的组合规则一致：正文（可附答案）+ 错题
  assert.equal(
    formatSectionsOutput(one, { withAnswers: true, withWrong: false }),
    formatOutputWithAnswers(results, typeOrder),
  );
  assert.equal(
    formatSectionsOutput(one, { withAnswers: false, withWrong: true }),
    formatOutput(results, typeOrder) + formatWrongQuestionsTXT(results, typeOrder),
  );
  assert.equal(
    formatSectionsOutput(one, { withAnswers: true, withWrong: true }),
    formatOutputWithAnswers(results, typeOrder) + formatWrongQuestionsTXT(results, typeOrder),
  );

  // Markdown 侧同样逐字一致
  assert.equal(formatSectionsMD(one), formatOutputMD(results, typeOrder));
  assert.equal(
    formatSectionsOutputMD(one, { withAnswers: true, withWrong: true }),
    formatOutputWithAnswersMD(results, typeOrder) + formatWrongQuestionsMD(results, typeOrder),
  );

  // 没有章节信息时正文退化成空串，不会抛错
  assert.equal(formatSectionsText([]), '');
  assert.equal(formatSectionsOutput([], { withAnswers: true, withWrong: true }), '');
  assert.equal(formatSectionsMD([]), '');
  assert.equal(formatSectionsOutputMD([], { withAnswers: true, withWrong: true }), '');
});
