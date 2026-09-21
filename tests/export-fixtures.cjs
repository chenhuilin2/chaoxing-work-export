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

function chapter(id, title, questions) {
  return {
    id,
    title,
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

// 两个章节、两种题型：足够暴露「各章同题型被合并到一份」的回归
const CHAPTERS = [
  chapter('c1', '导论1', [question('q1', 'single-choice'), question('q2', 'multiple-choice')]),
  chapter('c2', '第一章 马克思主义', [question('q3', 'single-choice')]),
];

module.exports = { question, chapter, resultOf, exportOptions, CHAPTERS };
