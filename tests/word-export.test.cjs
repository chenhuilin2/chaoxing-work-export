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
  HeadingLevel: { HEADING_1: 'Heading1', HEADING_2: 'Heading2', HEADING_3: 'Heading3' },
  convertMillimetersToTwip: (mm) => Math.round((mm / 25.4) * 1440),
  // 真实实现返回 Blob；这里直接把 Document 交回去，方便检查段落
  Packer: { toBlob: async (document) => document },
};

const { generateWordBlob } = require('../.tmp/test/src/exporters/word-exporter.js');
const { planExportDocuments } = require('../.tmp/test/src/exporters/export-plan.js');
const { CHAPTERS, NESTED_CHAPTERS, chapter, exportOptions, question, resultOf } = require('./export-fixtures.cjs');

/** 渲染一份文档，直接返回它的段落数组 */
async function renderDocument(result, options) {
  const [plan] = planExportDocuments(result, options, '.docx').documents;
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

/** 正文里的小节标题（HEADING_3） */
function sectionHeadings(children) {
  return children
    .filter((child) => child instanceof Paragraph && child.heading === 'Heading3')
    .map(textOf);
}

/** 正文里带分页符的段落（用来验证「只有换章才另起一页」） */
function pageBreakCount(children) {
  return children.filter(
    (child) =>
      child instanceof Paragraph && (child.children ?? []).some((run) => run instanceof PageBreak),
  ).length;
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
  // 扁平目录（没有章信息）：退回「第 N 章 名字」单级标题
  assert.deepEqual(chapterHeadings(children), ['第 1 章 导论1', '第一章 马克思主义']);
  // 每章各自出现一次题型大题头（此前两章的单选会被合并成一个大题）
  assert.equal(texts.filter((text) => text.startsWith('一、单项选择题')).length, 2);
  // 两个章节的题干落在各自的分节区间里
  const firstHeading = texts.indexOf('第 1 章 导论1');
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

  const plans = planExportDocuments(
    result,
    {
      format: 'word',
      filename: '卷子',
      withAnswers: false,
      withWrong: false,
      shuffle: false,
      bankImport: false,
      splitByChapter: true,
    },
    '.docx',
  ).documents;
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

test('答案页与错题页跟在各自的章之后（章 1 题目 → 章 1 答案 → 章 1 错题 → 章 2 …）', async () => {
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

  // 答案页 / 错题页里的章标签按「章 1 答案、章 1 错题、章 2 答案、章 2 错题」排列
  assert.deepEqual(chapterLabels(children), [
    '第 1 章 导论1',
    '第 1 章 导论1',
    '第一章 马克思主义',
    '第一章 马克思主义',
  ]);
  // 题目页的章标题仍是 HEADING_2 大标题
  assert.deepEqual(chapterHeadings(children), ['第 1 章 导论1', '第一章 马克思主义']);
  // 两章各有一页错题汇总，而不是全篇只有末尾一页
  assert.equal(paragraphTexts(children).filter((text) => text === '错题汇总').length, 2);
  // 题号在正文 / 答案 / 错题三处都按章重新从 1 数起
  assert.deepEqual(numberPrefixes(children), [
    '1. ',
    '2. ',
    '1. ',
    '2. ',
    '1. ',
    '1. ',
    '1. ',
    '1. ',
  ]);
  // 每章：题目 →（分页）答案 →（分页）错题；换章再分页。共 1 + 1 + 1 + 1 + 1 = 5 次
  assert.equal(pageBreakCount(children), 5);
});

test('两级目录：章标题（Heading2）、小节标题（Heading3）、题型头（24）三层字号递减', async () => {
  const result = resultOf(NESTED_CHAPTERS, '思想道德与法治 - 3 个章节');

  const children = await renderDocument(result, {
    format: 'word',
    filename: '卷子',
    withAnswers: true,
    withWrong: false,
    shuffle: false,
    bankImport: false,
    splitByChapter: false,
  });

  // 章标题只用页面的章号拼，小节标题用页面的小节编号
  assert.deepEqual(chapterHeadings(children), [
    '第 1 章 担当复兴大任 成就时代新人',
    '第 2 章 领悟人生真谛 把握人生方向',
  ]);
  assert.deepEqual(sectionHeadings(children), [
    '1.1 如何认识新方位',
    '1.2 如何认识新使命',
    '2.1 为什么要思考人生问题',
  ]);
  // 题号按小节重新从 1 数起
  assert.deepEqual(numberPrefixes(children), ['1. ', '1. ', '1. ', '1. ', '1. ', '1. ']);

  const sizeOf = (heading) =>
    children.find((child) => child instanceof Paragraph && child.heading === heading).children[0]
      .size;
  const typeRun = children.find(
    (child) => child instanceof Paragraph && textOf(child).startsWith('一、单项选择题'),
  ).children[0];
  assert.equal(sizeOf('Heading2'), 28);
  assert.equal(sizeOf('Heading3'), 26);
  assert.equal(typeRun.size, 24);

  // 换章另起一页（第 2 章正文前）、每章的答案页再各分一次页；小节之间不换页，共 3 次
  assert.equal(pageBreakCount(children), 3);
});

test('两级目录：按章拆出的文件里只插小节标题，不重复章名', async () => {
  const result = resultOf(NESTED_CHAPTERS, '思想道德与法治 - 3 个章节');

  const plans = planExportDocuments(
    result,
    {
      format: 'word',
      filename: '卷子',
      withAnswers: false,
      withWrong: false,
      shuffle: false,
      bankImport: false,
      splitByChapter: true,
    },
    '.docx',
  ).documents;

  assert.equal(plans.length, 2);
  const first = await generateWordBlob(plans[0].title, plans[0].sections, {
    withAnswers: false,
    withWrong: false,
    bankImport: false,
  });
  const children = first.sections[0].children;

  // 文件大标题已经是章名，正文里不再插章标题；小节标题照旧
  assert.deepEqual(chapterHeadings(children), []);
  assert.deepEqual(sectionHeadings(children), ['1.1 如何认识新方位', '1.2 如何认识新使命']);
  assert.equal(pageBreakCount(children), 0);
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

test('答案页：有解析的题目在答案下方插一段「解析：」，没有解析的不插', async () => {
  const withAnalysis = {
    ...question('q1', 'single-choice'),
    analysis: [{ type: 'text', text: '因为它确实是 A' }],
  };
  const result = resultOf(
    [chapter('c1', '导论1', [withAnalysis, question('q2', 'single-choice')])],
    '导论1',
  );

  const children = await renderDocument(result, exportOptions({ withAnswers: true }));
  const paragraphs = paragraphTexts(children);
  const analysisParagraphs = paragraphs.filter((text) => text.startsWith('解析：'));

  // 只有第一题有解析：整份文档只多一段，且紧跟在它的答案后面
  assert.deepEqual(analysisParagraphs, ['解析：因为它确实是 A']);
  const answerIndex = paragraphs.findIndex((text) => text === '1. A');
  assert.equal(paragraphs[answerIndex + 1], '解析：因为它确实是 A');

  // 解析比答案小一号、灰蓝色、浅底，让它在答案页里不抢答案的视觉层级
  const analysisNode = children.find(
    (child) => child instanceof Paragraph && textOf(child).startsWith('解析：'),
  );
  assert.equal(analysisNode.children[0].size, 20);
  assert.equal(analysisNode.children[0].color, '6B7280');
  assert.equal(analysisNode.shading.fill, 'F7F9FC');
});

test('答案页：未勾选「附加答案」时不出现解析', async () => {
  const withAnalysis = {
    ...question('q1', 'single-choice'),
    analysis: [{ type: 'text', text: '因为它确实是 A' }],
  };
  const result = resultOf([chapter('c1', '导论1', [withAnalysis])], '导论1');

  const children = await renderDocument(result, exportOptions({ withAnswers: false }));
  assert.equal(
    paragraphTexts(children).some((text) => text.startsWith('解析：')),
    false,
  );
});

test('题库导入：判断题答案按页面写法归一化成「对 / 错」，不再一律输出「错」', async () => {
  // 学习通各页面写法不统一：作业页给「对 / 错」，章节测验给「√ / ×」，还有「正确 / 错误」「A / B」。
  // 此前只认 √/✓/T/正确，页面给「对」时全部落到 else 分支 —— 整份判断题答案都是「错」。
  const cases = [
    ['对', '对'],
    ['正确', '对'],
    ['√', '对'],
    ['A', '对'],
    ['错', '错'],
    ['错误', '错'],
    ['×', '错'],
    ['B', '错'],
  ];
  const questions = cases.map(([text], index) => ({
    ...question(`tf${index}`, 'true-false'),
    correctAnswer: [{ type: 'text', text }],
  }));
  const result = resultOf([chapter('c1', '判断题集', questions)], '判断题集');

  const children = await renderDocument(result, exportOptions({ bankImport: true }));
  const answers = paragraphTexts(children).filter((text) => text.startsWith('答案：'));

  assert.deepEqual(
    answers,
    cases.map(([, expected]) => `答案：${expected}`),
  );
});

test('题库导入：判断题答案认不出来时保留页面原文，不猜「对 / 错」', async () => {
  const weird = {
    ...question('tf-weird', 'true-false'),
    correctAnswer: [{ type: 'text', text: '视情况而定' }],
  };
  const result = resultOf([chapter('c1', '判断题集', [weird])], '判断题集');

  const children = await renderDocument(result, exportOptions({ bankImport: true }));
  assert.ok(paragraphTexts(children).includes('答案：视情况而定'));
});

/** 给「抓图 → ImageRun」这条链路搭最小桩：fetch 返回图片 Blob，FileReader 给 dataURL，Image 给尺寸 */
function stubImagePipeline() {
  const saved = {
    fetch: globalThis.fetch,
    FileReader: globalThis.FileReader,
    Image: globalThis.Image,
  };
  globalThis.fetch = async () => ({ ok: true, blob: async () => ({ type: 'image/png' }) });
  globalThis.FileReader = class {
    readAsDataURL() {
      this.result = 'data:image/png;base64,iVBORw0KGgo=';
      this.onloadend?.();
    }
  };
  globalThis.Image = class {
    set src(value) {
      this.naturalWidth = 120;
      this.naturalHeight = 80;
      this.onload?.();
    }
  };
  return () => Object.assign(globalThis, saved);
}

test('带图片的题目：ImageRun 的 data 必须是 dataURL 字符串（issue #21）', async () => {
  // 此前把 fetchImageAsset 返回的 { data, type, width, height } 整个对象塞进 ImageRun.data，
  // docx 只认 string / Blob / ArrayBuffer，打包时 JSZip 报
  // 「Can't read the data of 'word/media/xxx.png'」，题目一带图片整个 Word 导出就失败。
  const imageQuestion = {
    ...question('q-img', 'single-choice'),
    stem: [
      { type: 'text', text: '看图作答' },
      { type: 'break' },
      { type: 'image', url: 'https://example.com/a.png', alt: '图' },
    ],
  };
  const result = resultOf([chapter('c-img', '图片题', [imageQuestion])], '图片题');

  const restore = stubImagePipeline();
  try {
    const children = await renderDocument(result, exportOptions({ withAnswers: false }));
    const imageRun = children
      .flatMap((node) => node.children ?? [])
      .find((child) => child instanceof ImageRun);

    assert.ok(imageRun, '应当生成一个 ImageRun');
    assert.equal(typeof imageRun.data, 'string');
    assert.ok(imageRun.data.startsWith('data:image/png;base64,'));
  } finally {
    restore();
  }
});
