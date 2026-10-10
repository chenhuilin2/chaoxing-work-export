// Word 导出相关测试的共用夹具：造 ExtractionResult / 章节 / 题目 / 导出选项。
// 文件名不以 .test.cjs 结尾，因此不会被 tools/test.mjs 当成用例收集。
const { buildStatistics, deriveTypeOrder } = require('../.tmp/test/src/domain/question.js');

/** 造一道题（只填规划与导出链路会用到的字段） */
function question(id, type) {
  return {
    id,
    type,
    stem: [{ type: 'text', text: `题干${id}` }],
    options:
      type === 'single-choice'
        ? [{ key: 'A', content: [{ type: 'text', text: `选项${id}` }] }]
        : [],
    correctAnswer: [{ type: 'text', text: 'A' }],
    userAnswer: [],
    analysis: [],
    isWrong: false,
    source: { extractor: 'timu', pageUrl: 'https://mooc1.chaoxing.com/x' },
  };
}

/** 造一个「章」（目录第一级的分组标题，页面上编号是「1」「2」） */
function group(id, label, title) {
  return { id, label, title };
}

/**
 * 造一个小节。
 * 不传 extra 时是「扁平目录 / 老数据」：没有章信息也没有小节编号，导出会退回「第 N 章 名字」。
 */
function chapter(id, title, questions, extra = {}) {
  return {
    id,
    title,
    label: extra.label,
    group: extra.group ?? null,
    questions,
    typeOrder: deriveTypeOrder(questions),
    statistics: buildStatistics(questions),
    sourceUrl: 'https://mooc1.chaoxing.com/x',
    extractor: 'timu',
  };
}

function resultOf(chapters, title) {
  const questions = chapters.flatMap((item) => item.questions);
  return {
    title,
    questions,
    typeOrder: deriveTypeOrder(questions),
    statistics: buildStatistics(questions),
    sourceUrl: 'https://mooc1.chaoxing.com/x',
    extractor: 'timu',
    extractedAt: '2026-09-21T00:00:00.000Z',
    chapters,
  };
}

/** 面板要求完整导出选项，这里给一份「全不选」的基线 */
function exportOptions(overrides = {}) {
  return {
    format: 'word',
    filename: '卷子',
    withAnswers: false,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
    ...overrides,
  };
}

// 两个章节、两种题型：足够暴露「各章同题型被合并到一份」的回归。
// 没有章信息，走的是「扁平目录 / 老数据」那条兜底路径。
const CHAPTERS = [
  chapter('c1', '导论1', [question('q1', 'single-choice'), question('q2', 'multiple-choice')]),
  chapter('c2', '第一章 马克思主义', [question('q3', 'single-choice')]),
];

// 两级目录（学习通真实结构）：第 1 章下两个小节、第 2 章下一个小节，编号取自页面
const GROUP_A = group('g1', '1', '担当复兴大任 成就时代新人');
const GROUP_B = group('g2', '2', '领悟人生真谛 把握人生方向');
const NESTED_CHAPTERS = [
  chapter('c1', '如何认识新方位', [question('q1', 'single-choice')], {
    label: '1.1',
    group: GROUP_A,
  }),
  chapter('c2', '如何认识新使命', [question('q2', 'multiple-choice')], {
    label: '1.2',
    group: GROUP_A,
  }),
  chapter('c3', '为什么要思考人生问题', [question('q3', 'single-choice')], {
    label: '2.1',
    group: GROUP_B,
  }),
];

module.exports = {
  question,
  group,
  chapter,
  resultOf,
  exportOptions,
  CHAPTERS,
  GROUP_A,
  GROUP_B,
  NESTED_CHAPTERS,
};
