// 导出文件的「文件划分 + 章节分组」规划用例（Word / TXT / Markdown 共用同一份计划）。
//
// 只测不碰 docx 的纯逻辑（拆几个文件、每个文件装哪些章节、文件名怎么起、什么时候打包成 zip）——
// word-exporter 依赖 CDN 注入的全局 docx，Node 单测里加载不了。
const test = require('node:test');
const assert = require('node:assert/strict');
const {
  CHAPTER_ARCHIVE_THRESHOLD,
  chapterHeadingText,
  planExportDocuments,
} = require('../.tmp/test/src/exporters/export-plan.js');
const {
  CHAPTERS,
  NESTED_CHAPTERS,
  chapter,
  exportOptions,
  group,
  question,
  resultOf,
} = require('./export-fixtures.cjs');

/** 造 n 个各带一道单选题的章节，用来验证拆分阈值 */
function manyChapters(count) {
  return Array.from({ length: count }, (_, index) =>
    chapter(`c${index + 1}`, `导论${index + 1}`, [question(`q${index + 1}`, 'single-choice')]),
  );
}

test('单章节结果：只产出一个文件，正文不分节', () => {
  const single = resultOf([chapter('c1', '导论1', [question('q1', 'single-choice')])], '导论1');

  // 连「按章节拆分」勾上也不拆——只有一章没有拆的意义
  const plan = planExportDocuments(single, { ...exportOptions(), splitByChapter: true }, '.docx');

  assert.equal(plan.documents.length, 1);
  assert.equal(plan.documents[0].filename, '卷子.docx');
  assert.equal(plan.documents[0].sections.length, 1);
  // 只有一节：两级标题都不插，正文直接就是内容
  assert.equal(plan.documents[0].sections[0].heading, null);
  assert.equal(plan.documents[0].sections[0].chapterHeading, null);
  assert.equal(plan.archive, null);
});

test('多章节但不拆分：一个文件里按章节分成多节，各章同题型不再混在一起', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const plan = planExportDocuments(result, exportOptions(), '.docx');

  assert.equal(plan.documents.length, 1);
  assert.equal(plan.documents[0].filename, '卷子.docx');
  // 文档大标题取聚合标题，章节名留给正文分节标题
  assert.equal(plan.documents[0].title, '学生学习页面 - 2 个章节');
  // 扁平目录：没有章信息，退回「第 N 章 名字」单级标题、没有小节标题
  assert.deepEqual(
    plan.documents[0].sections.map((section) => section.chapterHeading),
    ['第 1 章 导论1', '第一章 马克思主义'],
  );
  assert.deepEqual(
    plan.documents[0].sections.map((section) => section.heading),
    [null, null],
  );
  // 每节只装自己那章的题：单选题没有被合并成一份
  assert.deepEqual(
    plan.documents[0].sections.map((section) => section.results['单选'].length),
    [1, 1],
  );
  assert.equal(plan.documents[0].sections[0].results['单选'][0].stemContent[0].text, '题干q1');
  assert.equal(plan.documents[0].sections[1].results['单选'][0].stemContent[0].text, '题干q3');
  assert.equal(plan.documents[0].sections[0].results['多选'].length, 1);
  assert.equal(plan.documents[0].sections[1].results['多选'].length, 0);
  assert.equal(plan.archive, null);
});

test('勾选「按章节拆分文件」：一章一个文件，文件名就是章名，扩展名随格式变化', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const word = planExportDocuments(result, { ...exportOptions(), splitByChapter: true }, '.docx');
  // 文件名不含课程名与序号前缀：解压后的文件夹里只有「第 N 章 …」，一眼能挑出要看的章
  assert.deepEqual(
    word.documents.map((plan) => plan.filename),
    ['导论1.docx', '第一章 马克思主义.docx'],
  );
  // 每个文件只装一章（扁平目录下这一「章」就是该小节），文件大标题即它的名字，
  // 正文里两级标题都不插
  word.documents.forEach((plan) => {
    assert.equal(plan.sections.length, 1);
    assert.equal(plan.sections[0].heading, null);
  });
  assert.deepEqual(
    word.documents.map((plan) => plan.title),
    ['导论1', '第一章 马克思主义'],
  );

  // TXT / Markdown 用的是同一份划分，只有扩展名不同
  const txt = planExportDocuments(result, { ...exportOptions(), splitByChapter: true }, '.txt');
  assert.deepEqual(
    txt.documents.map((plan) => plan.filename),
    ['导论1.txt', '第一章 马克思主义.txt'],
  );
  const md = planExportDocuments(result, { ...exportOptions(), splitByChapter: true }, '.md');
  assert.deepEqual(
    md.documents.map((plan) => plan.filename),
    ['导论1.md', '第一章 马克思主义.md'],
  );
});

test('章节名里的文件名非法字符会被替换，空章节名回退到「章节N」', () => {
  const dirty = resultOf(
    [
      chapter('c1', '第2章：马克思/恩格斯', [question('q1', 'single-choice')]),
      chapter('c2', '', [question('q2', 'fill-blank')]),
    ],
    '学生学习页面 - 2 个章节',
  );

  const plan = planExportDocuments(
    dirty,
    { ...exportOptions({ filename: '卷子.docx' }), splitByChapter: true },
    '.docx',
  );

  assert.deepEqual(
    plan.documents.map((document) => document.filename),
    ['第2章：马克思_恩格斯.docx', '章节2.docx'],
  );
  // 文件大标题保留原始章节名（只有文件名做替换），空章节名回退到「章节N」
  assert.equal(plan.documents[0].title, '第2章：马克思/恩格斯');
  assert.equal(plan.documents[1].title, '章节2');
});

test('章节名清洗后撞车时补 `_2` 后缀，不让两个文件互相覆盖', () => {
  const clash = resultOf(
    [
      chapter('c1', '第1章 A/B', [question('q1', 'single-choice')]),
      chapter('c2', '第1章 A?B', [question('q2', 'single-choice')]),
    ],
    '卷子',
  );

  const plan = planExportDocuments(clash, { ...exportOptions(), splitByChapter: true }, '.txt');

  assert.deepEqual(
    plan.documents.map((document) => document.filename),
    ['第1章 A_B.txt', '第1章 A_B_2.txt'],
  );
});

test('两级目录：正文标题按「第 N 章 章名」+「1.1 小节名」算好', () => {
  const result = resultOf(NESTED_CHAPTERS, '思想道德与法治 - 3 个章节');

  const plan = planExportDocuments(result, exportOptions(), '.txt');

  assert.deepEqual(
    plan.documents[0].sections.map((section) => section.chapterHeading),
    [
      '第 1 章 担当复兴大任 成就时代新人',
      '第 1 章 担当复兴大任 成就时代新人',
      '第 2 章 领悟人生真谛 把握人生方向',
    ],
  );
  assert.deepEqual(
    plan.documents[0].sections.map((section) => section.heading),
    ['1.1 如何认识新方位', '1.2 如何认识新使命', '2.1 为什么要思考人生问题'],
  );
});

test('两级目录：按章节拆分是「一章一个文件」，文件内含该章的所有小节', () => {
  const result = resultOf(NESTED_CHAPTERS, '思想道德与法治 - 3 个章节');

  const plan = planExportDocuments(result, { ...exportOptions(), splitByChapter: true }, '.docx');

  // 3 个小节属于 2 章：只拆出 2 个文件，而不是 3 个；文件名就是章名
  assert.deepEqual(
    plan.documents.map((document) => document.filename),
    ['第 1 章 担当复兴大任 成就时代新人.docx', '第 2 章 领悟人生真谛 把握人生方向.docx'],
  );
  // 文件大标题是章名，正文里只插小节标题（章名已经在标题上，不必重复）
  assert.equal(plan.documents[0].title, '第 1 章 担当复兴大任 成就时代新人');
  assert.deepEqual(
    plan.documents[0].sections.map((section) => section.heading),
    ['1.1 如何认识新方位', '1.2 如何认识新使命'],
  );
  assert.equal(plan.documents[1].sections.length, 1);
  assert.equal(plan.archive, null);
});

test('两级目录：压缩包阈值按「章数」算，而不是小节数', () => {
  const chapters = [];
  for (let index = 1; index <= 6; index += 1) {
    const info = group(`g${index}`, String(index), `章名${index}`);
    chapters.push(
      chapter(`c${index}a`, `小节${index}A`, [question(`q${index}a`, 'single-choice')], {
        label: `${index}.1`,
        group: info,
      }),
      chapter(`c${index}b`, `小节${index}B`, [question(`q${index}b`, 'single-choice')], {
        label: `${index}.2`,
        group: info,
      }),
    );
  }

  const plan = planExportDocuments(
    resultOf(chapters, '卷子'),
    { ...exportOptions(), splitByChapter: true },
    '.txt',
  );

  // 12 个小节但只有 6 章：6 个文件、超过阈值 → 打包（阈值数的是章）
  assert.equal(plan.documents.length, 6);
  assert.deepEqual(plan.archive, { filename: '卷子.zip', documentCount: 6 });
});

test('题库导入：多章节也不插章节节，只按需要拆文件', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const merged = planExportDocuments(result, exportOptions({ bankImport: true }), '.docx');
  assert.equal(merged.documents.length, 1);
  // 智能导入格式不认章节标题，整篇保持题目按题型连续排列
  assert.equal(merged.documents[0].sections.length, 1);

  const split = planExportDocuments(
    result,
    { ...exportOptions({ bankImport: true }), splitByChapter: true },
    '.docx',
  );
  assert.equal(split.documents.length, 2);
  split.documents.forEach((document) => assert.equal(document.sections.length, 1));
});

test('未指定文件名时回退到结果标题', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const plan = planExportDocuments(result, exportOptions({ filename: '' }), '.txt');

  assert.equal(plan.documents[0].filename, '学生学习页面 - 2 个章节.txt');
});

test(`逐章文件数不超过 ${CHAPTER_ARCHIVE_THRESHOLD} 个时逐个下载，超过就打包成一个 zip`, () => {
  assert.equal(CHAPTER_ARCHIVE_THRESHOLD, 5);

  // 正好到阈值：仍然逐个下载
  const atLimit = planExportDocuments(
    resultOf(manyChapters(5), '卷子'),
    { ...exportOptions(), splitByChapter: true },
    '.txt',
  );
  assert.equal(atLimit.documents.length, 5);
  assert.equal(atLimit.archive, null);

  // 超过阈值：不再逐个下载，改成「卷子.zip」里装 6 个文件
  const over = planExportDocuments(
    resultOf(manyChapters(6), '卷子'),
    { ...exportOptions(), splitByChapter: true },
    '.txt',
  );
  assert.equal(over.documents.length, 6);
  assert.deepEqual(over.archive, { filename: '卷子.zip', documentCount: 6 });

  // 不拆分时再多章节也不会打包：只有一个合并文件
  const merged = planExportDocuments(resultOf(manyChapters(20), '卷子'), exportOptions(), '.txt');
  assert.equal(merged.documents.length, 1);
  assert.equal(merged.archive, null);
});

test('章节标题：没有编号时补「第 N 章」，自带编号时原样使用', () => {
  assert.equal(chapterHeadingText('如何认识新方位', 0), '第 1 章 如何认识新方位');
  assert.equal(chapterHeadingText('如何认识新使命', 4), '第 5 章 如何认识新使命');
  // 章节名自带编号：不要叠成「第 1 章 第一章 马克思主义」
  assert.equal(chapterHeadingText('第一章 马克思主义', 0), '第一章 马克思主义');
  assert.equal(chapterHeadingText('第2章：马克思/恩格斯', 1), '第2章：马克思/恩格斯');
  assert.equal(chapterHeadingText('第 3 讲 绪论', 2), '第 3 讲 绪论');
  // 空章节名回退
  assert.equal(chapterHeadingText('', 2), '第 3 章 章节3');
});
