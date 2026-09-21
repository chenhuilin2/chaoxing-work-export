// Word 渲染用例：把 CDN 注入的全局 docx 换成可检查的打桩实现，
// 直接断言「生成出来的段落序列」。本轮把试卷正文改成了按章节渲染，这段代码此前没有覆盖。
const test = require('node:test');
const assert = require('node:assert/strict');

/** docx.js 的最小打桩：只记录构造参数，便于断言段落序列 */
class TextRun {
  constructor(options) {
    Object.assign(this, options);
  }
}
class Paragraph {
  constructor(options) {
    Object.assign(this, options);
  }
}
class ImageRun {
  constructor(options) {
    Object.assign(this, options);
  }
}
class PageBreak {}
class Document {
  constructor(options) {
    Object.assign(this, options);
  }
}

globalThis.window = {};
globalThis.docx = {
  TextRun,
  Paragraph,
  ImageRun,
  PageBreak,
  Document,
  AlignmentType: { CENTER: 'center', LEFT: 'left' },
  HeadingLevel: { HEADING_1: 'Heading1', HEADING_2: 'Heading2' },
  convertMillimetersToTwip: (mm) => Math.round((mm / 25.4) * 1440),
  // 真实实现返回 Blob；这里直接把 Document 交回去，方便检查段落
  Packer: { toBlob: async (document) => document },
};

const { generateWordBlob } = require('../.tmp/test/src/exporters/word-exporter.js');
const { planWordDocuments } = require('../.tmp/test/src/exporters/word-plan.js');
const { CHAPTERS, chapter, question, resultOf } = require('./export-fixtures.cjs');

/** 渲染一份文档，直接返回它的段落数组 */
async function renderDocument(result, options) {
  const [plan] = planWordDocuments(result, options);
  const document = await generateWordBlob(plan.title, plan.sections, {
    withAnswers: options.bankImport || options.withAnswers,
    withWrong: !options.bankImport && options.withWrong,
    bankImport: options.bankImport,
  });
  return document.sections[0].children;
}

function textOf(paragraph) {
  return (paragraph.children ?? [])
    .map((child) => (child instanceof TextRun ? (child.text ?? '') : ''))
    .join('');
}

function paragraphTexts(children) {
  return children.filter((child) => child instanceof Paragraph).map(textOf);
}

/** 正文里的章节大标题（HEADING_2） */
function chapterHeadings(children) {
  return children
    .filter((child) => child instanceof Paragraph && child.heading === 'Heading2')
    .map(textOf);
}

/** 答案页 / 错题页里的章节小标题（26 号字，不带 heading） */
function chapterLabels(children) {
  return children
    .filter(
      (child) =>
        child instanceof Paragraph &&
        !child.heading &&
        (child.children ?? []).some((run) => run instanceof TextRun && run.size === 26),
    )
    .map(textOf);
}

/** 取正文里带题号的段落前缀，用来验证编号连续 */
function numberPrefixes(children) {
  return paragraphTexts(children)
    .map((text) => /^\d+\.\s/.exec(text)?.[0])
    .filter(Boolean);
}

test('单章节文档：不插章节标题，标题与题型分组与原先一致', async () => {
  const result = resultOf([CHAPTERS[0]], '导论1');

  const children = await renderDocument(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: true,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
  });
  const texts = paragraphTexts(children);

  assert.deepEqual(chapterHeadings(children), []);
  assert.equal(texts[0], '导论1');
  assert.ok(texts.some((text) => text.startsWith('一、单项选择题（本大题共1小题）')));
  assert.ok(texts.some((text) => text.startsWith('二、多项选择题（本大题共1小题）')));
  assert.deepEqual(numberPrefixes(children), ['1. ', '2. ', '1. ', '2. ']);
  assert.ok(texts.includes('正确答案'));
});

test('多章节单文件：正文按章节分节，各章同题型不再混在同一大题里', async () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const children = await renderDocument(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: false,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
  });
  const texts = paragraphTexts(children);

  assert.equal(texts[0], '学生学习页面 - 2 个章节');
  assert.deepEqual(chapterHeadings(children), ['导论1', '第一章 马克思主义']);
  // 每章各自出现一次题型大题头（此前两章的单选会被合并成一个大题）
  assert.equal(texts.filter((text) => text.startsWith('一、单项选择题')).length, 2);
  // 两个章节的题干落在各自的分节区间里
  const firstHeading = texts.indexOf('导论1');
  const secondHeading = texts.indexOf('第一章 马克思主义');
  const firstStem = texts.findIndex((text) => text.includes('题干q1'));
  const thirdStem = texts.findIndex((text) => text.includes('题干q3'));
  assert.ok(firstHeading < firstStem && firstStem < secondHeading && secondHeading < thirdStem);
  // 题号按章重新从 1 数起（此前是跨章连续，第二章的第一题会从 3 开始）
  assert.deepEqual(numberPrefixes(children), ['1. ', '2. ', '1. ']);
});

test('多章节单文件：题型序号按章重排，大题头字号小于章节标题', async () => {
  // 第二章没有单选题：序号若按全局固定位次取，多选就会顶着「二、多项选择题」的跳号
  const chapters = [
    chapter('c1', '第一章', [question('q1', 'single-choice'), question('q2', 'multiple-choice')]),
    chapter('c2', '第二章', [question('q3', 'multiple-choice'), question('q4', 'true-false')]),
  ];
  const result = resultOf(chapters, '课程 - 2 个章节');

  const children = await renderDocument(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: false,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
  });
  const texts = paragraphTexts(children);

  // 第一章：一、单项选择题 / 二、多项选择题；第二章：一、多项选择题 / 二、判断题
  assert.deepEqual(
    texts.filter((text) => /^[一二三四五六七八九十]、/.test(text)),
    [
      '一、单项选择题（本大题共1小题）',
      '二、多项选择题（本大题共1小题）',
      '一、多项选择题（本大题共1小题）',
      '二、判断题（本大题共1小题）',
    ],
  );

  // 题型大题头（24 半磅 = 12pt）比章节大标题（28 半磅 = 14pt）小一号
  const chapterHeading = children.find(
    (child) => child instanceof Paragraph && child.heading === 'Heading2',
  );
  assert.equal(chapterHeading.children[0].size, 28);
  const typeHeader = children.find(
    (child) => child instanceof Paragraph && textOf(child).startsWith('一、单项选择题'),
  );
  assert.equal(typeHeader.children[0].size, 24);
  assert.ok(typeHeader.children[0].size < chapterHeading.children[0].size);
});

test('按章节拆分：每个文件只装一章，标题即章节名', async () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const plans = planWordDocuments(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: false,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: true,
  });
  assert.equal(plans.length, 2);

  for (const plan of plans) {
    const document = await generateWordBlob(plan.title, plan.sections, {
      withAnswers: false,
      withWrong: false,
      bankImport: false,
    });
    const children = document.sections[0].children;
    const texts = paragraphTexts(children);

    // 一个文件只有一章，因此不再插章节标题
    assert.deepEqual(chapterHeadings(children), []);
    assert.equal(texts[0], plan.title);
    // 题号从 1 重新开始
    assert.equal(numberPrefixes(children)[0], '1. ');
  }

  const firstTexts = paragraphTexts(
    (
      await generateWordBlob(plans[0].title, plans[0].sections, {
        withAnswers: false,
        withWrong: false,
        bankImport: false,
      })
    ).sections[0].children,
  );
  assert.ok(firstTexts.some((text) => text.includes('题干q1')));
  assert.ok(!firstTexts.some((text) => text.includes('题干q3')));
});

test('答案页与错题页也按章节分节', async () => {
  // 两章各留一道错题，错题汇总页才会生成
  const wrongChapters = [
    chapter('c1', '导论1', [
      { ...question('q1', 'single-choice'), isWrong: true },
      question('q2', 'multiple-choice'),
    ]),
    chapter('c2', '第一章 马克思主义', [{ ...question('q3', 'single-choice'), isWrong: true }]),
  ];
  const result = resultOf(wrongChapters, '学生学习页面 - 2 个章节');

  const children = await renderDocument(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: true,
    withWrong: true,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
  });

  // 答案页两章各有小标题，错题页同样；题目页那两处是 HEADING_2 大标题
  assert.deepEqual(chapterLabels(children), [
    '导论1',
    '第一章 马克思主义',
    '导论1',
    '第一章 马克思主义',
  ]);
  assert.deepEqual(chapterHeadings(children), ['导论1', '第一章 马克思主义']);
  assert.equal(paragraphTexts(children).filter((text) => text === '错题汇总').length, 1);
  // 题号在正文 / 答案页 / 错题页三处都按章重新从 1 数起
  assert.deepEqual(numberPrefixes(children), [
    '1. ',
    '2. ',
    '1. ',
    '1. ',
    '2. ',
    '1. ',
    '1. ',
    '1. ',
  ]);
});

test('题库导入格式：不认章节标题，题目按题型连续排列', async () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const children = await renderDocument(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: false,
    withWrong: false,
    shuffle: false,
    bankImport: true,
    splitByChapter: false,
  });
  const texts = paragraphTexts(children);

  assert.deepEqual(chapterHeadings(children), []);
  assert.equal(texts[0], '学生学习页面 - 2 个章节');
  assert.equal(texts.filter((text) => text.includes('【单选题】')).length, 2);
  assert.equal(texts.filter((text) => text.includes('【多选题】')).length, 1);
  // 章节名不会作为独立段落混进导入格式
  assert.ok(!texts.includes('导论1'));
  assert.ok(!texts.includes('第一章 马克思主义'));
  // 两章的单选先连在一起，多选排在其后
  const lastSingle = texts.findLastIndex((text) => text.includes('【单选题】'));
  const firstMultiple = texts.findIndex((text) => text.includes('【多选题】'));
  assert.ok(lastSingle < firstMultiple);
});
