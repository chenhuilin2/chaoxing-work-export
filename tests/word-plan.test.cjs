// Word 导出的「文件划分 + 章节分组」规划用例。
//
// 只测不碰 docx 的纯逻辑（拆几个文件、每个文件装哪些章节、文件名怎么起）——
// word-exporter 依赖 CDN 注入的全局 docx，Node 单测里加载不了。
const test = require('node:test');
const assert = require('node:assert/strict');
const { planWordDocuments } = require('../.tmp/test/src/exporters/word-plan.js');
const { CHAPTERS, chapter, exportOptions, question, resultOf } = require('./export-fixtures.cjs');

test('单章节结果：只产出一个文件，正文不分节', () => {
  const single = resultOf([chapter('c1', '导论1', [question('q1', 'single-choice')])], '导论1');

  // 连「按章节拆分」勾上也不拆——只有一章没有拆的意义
  const plans = planWordDocuments(single, { ...exportOptions(), splitByChapter: true });

  assert.equal(plans.length, 1);
  assert.equal(plans[0].filename, '卷子.docx');
  assert.equal(plans[0].sections.length, 1);
  assert.equal(plans[0].sections[0].title, '导论1');
});

test('多章节但不拆分：一个文件里按章节分成多节，各章同题型不再混在一起', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const plans = planWordDocuments(result, exportOptions());

  assert.equal(plans.length, 1);
  assert.equal(plans[0].filename, '卷子.docx');
  // 文档大标题取聚合标题，章节名留给正文分节标题
  assert.equal(plans[0].title, '学生学习页面 - 2 个章节');
  assert.deepEqual(
    plans[0].sections.map((section) => section.title),
    ['导论1', '第一章 马克思主义'],
  );
  // 每节只装自己那章的题：单选题没有被合并成一份
  assert.deepEqual(
    plans[0].sections.map((section) => section.results['单选'].length),
    [1, 1],
  );
  assert.equal(plans[0].sections[0].results['单选'][0].stemContent[0].text, '题干q1');
  assert.equal(plans[0].sections[1].results['单选'][0].stemContent[0].text, '题干q3');
  assert.equal(plans[0].sections[0].results['多选'].length, 1);
  assert.equal(plans[0].sections[1].results['多选'].length, 0);
});

test('勾选「按章节拆分文件」：一章一个 docx，文件名带序号与章节名', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const plans = planWordDocuments(result, { ...exportOptions(), splitByChapter: true });

  assert.equal(plans.length, 2);
  assert.deepEqual(
    plans.map((plan) => plan.filename),
    ['卷子_1_导论1.docx', '卷子_2_第一章 马克思主义.docx'],
  );
  // 每个文件只装一章，文档标题即章节名（正文里不再插章节标题）
  plans.forEach((plan) => {
    assert.equal(plan.sections.length, 1);
    assert.equal(plan.title, plan.sections[0].title);
  });
});

test('章节名里的文件名非法字符会被替换，文件名扩展名会被去掉', () => {
  const dirty = resultOf(
    [
      chapter('c1', '第2章：马克思/恩格斯', [question('q1', 'single-choice')]),
      chapter('c2', '', [question('q2', 'fill-blank')]),
    ],
    '学生学习页面 - 2 个章节',
  );

  const plans = planWordDocuments(dirty, {
    ...exportOptions({ filename: '卷子.docx' }),
    splitByChapter: true,
  });

  assert.deepEqual(
    plans.map((plan) => plan.filename),
    ['卷子_1_第2章：马克思_恩格斯.docx', '卷子_2_章节2.docx'],
  );
  // 文档标题保留原始章节名（只有文件名做替换），空章节名回退到「章节N」
  assert.equal(plans[0].title, '第2章：马克思/恩格斯');
  assert.equal(plans[1].title, '章节2');
});

test('题库导入：多章节也不插章节节，只按需要拆文件', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const merged = planWordDocuments(result, exportOptions({ bankImport: true }));
  assert.equal(merged.length, 1);
  // 智能导入格式不认章节标题，整篇保持题目按题型连续排列
  assert.equal(merged[0].sections.length, 1);

  const split = planWordDocuments(result, {
    ...exportOptions({ bankImport: true }),
    splitByChapter: true,
  });
  assert.equal(split.length, 2);
  split.forEach((plan) => assert.equal(plan.sections.length, 1));
});

test('未指定文件名时回退到结果标题', () => {
  const result = resultOf(CHAPTERS, '学生学习页面 - 2 个章节');

  const plans = planWordDocuments(result, exportOptions({ filename: '' }));

  assert.equal(plans[0].filename, '学生学习页面 - 2 个章节.docx');
});
