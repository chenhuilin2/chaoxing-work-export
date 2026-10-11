// 试卷批阅结果页 / 章节测验「已批阅」视图：页面只渲染「我的答案」+ 一个对错图标，
// **不输出「正确答案」文本**时，如何拿到答案。
//
// 规则（宁可缺，不可错）：
// 1. 批阅为全对（.marking_dui 等）→ 我的答案与标准答案一致，回填我的答案；
// 2. 批阅错误 / 部分正确 → 页面确实没给正确答案，留空，且计入错题；
// 3. 页面本身有「正确答案」文本 → 一律以页面为准，不做推断。
//
// 这条路径此前只覆盖了 .TiMu 模板（tests/yipi-view.test.cjs），
// 试卷批阅页用的 .mark_item .questionLi 模板此前漏接回填，故单独成文。
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDocument } = require('./dom-setup.cjs');
const { MarkItemExtractor } = require('../.tmp/test/src/extractors/mark-item-extractor.js');
const { QuestionLiExtractor } = require('../.tmp/test/src/extractors/question-li-extractor.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

const LETTERS = ['A', 'B', 'C', 'D'];

/** 渲染一道试卷批阅页的单选题：只有「我的答案」+ 对错图标，没有「正确答案」文本 */
function markedQuestion(options = {}) {
  const {
    number = 1,
    stem = '（）标志着“马克思主义的中国化”这一命题的正式提出。',
    optionTexts = ['党的七大', '党的五大', '党的七届二中全会', '党的六届六中全会'],
    myAnswer = 'D',
    gradingIcon = '<span class="marking_dui"></span>',
    correctAnswer = '',
  } = options;

  const optionList = `<ul class="mark_letter">${optionTexts
    .map(
      (text, index) =>
        `<li><i class="fl">${LETTERS[index]}、</i><a href="javascript:void(0)" class="fl">${text}</a></li>`,
    )
    .join('')}</ul>`;

  const correctBx = correctAnswer
    ? `<span class="colorGreen"><span class="stuAnswerContent">${correctAnswer}</span></span>`
    : '';

  return `
    <div class="questionLi" typeName="单选题">
      <div class="Zy_TItle clearfix">
        <i class="fl">${number}</i>
        <div class="fontLabel">${stem}</div>
      </div>
      ${optionList}
      <div class="mark_answer">
        <div class="mark_key">
          ${correctBx}
          <span class="colorDeep"><span class="stuAnswerContent">${myAnswer}</span></span>
        </div>
        ${gradingIcon}
      </div>
    </div>`;
}

/** 试卷批阅页：按题型分块，每块 .type_tit + 若干 .questionLi */
function markItem(questions, typeTitle = '一、单选题（共2题）') {
  return `<div class="mark_item"><div class="type_tit">${typeTitle}</div>${questions}</div>`;
}

function extractWith(ExtractorClass, bodyHtml) {
  return new ExtractorClass().extract({
    root: createDocument(bodyHtml),
    title: '第一章',
    pageUrl: 'https://mooc1.chaoxing.com/mooc-ans/work/selectWorkQuestionYiPiYue',
  });
}

const markItemExtract = (bodyHtml) => extractWith(MarkItemExtractor, bodyHtml);
const questionLiExtract = (bodyHtml) => extractWith(QuestionLiExtractor, bodyHtml);

test('试卷批阅页：批阅正确时用「我的答案」回填正确答案（此前该模板完全不回填）', () => {
  const questions = markItemExtract(`<div id="ZyBottom">${markItem(markedQuestion())}</div>`);

  assert.equal(questions.length, 1);
  assert.equal(questions[0].type, 'single-choice');
  assert.equal(richContentToText(questions[0].userAnswer), 'D');
  assert.equal(richContentToText(questions[0].correctAnswer), 'D');
  assert.equal(questions[0].isWrong, false);
  // 选项与题干仍要正常解析，避免回填掩盖了别的问题
  assert.deepEqual(
    questions[0].options.map((option) => `${option.key}.${richContentToText(option.content)}`),
    ['A.党的七大', 'B.党的五大', 'C.党的七届二中全会', 'D.党的六届六中全会'],
  );
});

test('试卷批阅页：答错 / 部分正确不推断答案，留空并计入错题', () => {
  const questions = markItemExtract(`
    <div id="ZyBottom">
      ${markItem(
        [
          markedQuestion({
            number: 1,
            myAnswer: 'A',
            gradingIcon: '<span class="marking_cuo"></span>',
          }),
          markedQuestion({
            number: 2,
            myAnswer: 'B',
            gradingIcon: '<span class="marking_bandui"></span>',
          }),
        ].join(''),
      )}
    </div>`);

  assert.equal(questions.length, 2);
  // 页面没给正确答案，就不能把错答案当标准答案
  assert.equal(richContentToText(questions[0].correctAnswer), '');
  assert.equal(richContentToText(questions[1].correctAnswer), '');
  assert.equal(richContentToText(questions[0].userAnswer), 'A');
  assert.equal(questions[0].isWrong, true);
  assert.equal(questions[1].isWrong, true);
});

test('答对图标类名各版本变体都能识别（dui_true / icon_dui / icon-dui / 类名含 dui）', () => {
  const icons = [
    '<span class="marking_dui"></span>',
    '<span class="marking_dui_true"></span>',
    '<span class="icon_dui"></span>',
    '<span class="icon-dui"></span>',
    '<span class="ans-dui-icon"></span>',
    '<i data-result="right"></i>',
    '<i data-correct="true"></i>',
  ];

  icons.forEach((gradingIcon, index) => {
    const questions = markItemExtract(
      `<div id="ZyBottom">${markItem(markedQuestion({ number: index + 1, gradingIcon }))}</div>`,
    );
    assert.equal(
      richContentToText(questions[0].correctAnswer),
      'D',
      `图标 ${gradingIcon} 未被识别为答对`,
    );
  });
});

test('同一容器里既有答对又有错题标记（多次作答记录）时不回填，错误标记优先', () => {
  const questions = markItemExtract(
    `<div id="ZyBottom">${markItem(
      markedQuestion({
        gradingIcon: '<span class="marking_dui"></span><span class="marking_cuo"></span>',
      }),
    )}</div>`,
  );

  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), '');
  assert.equal(questions[0].isWrong, true);
});

test('页面本身给了「正确答案」文本时以页面为准，不做推断', () => {
  const questions = markItemExtract(
    `<div id="ZyBottom">${markItem(
      markedQuestion({
        myAnswer: 'D',
        correctAnswer: 'C',
        gradingIcon: '<span class="marking_dui"></span>',
      }),
    )}</div>`,
  );

  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), 'C');
});

test('没有任何批阅图标（未批阅）时不回填，避免凭空造答案', () => {
  const questions = markItemExtract(
    `<div id="ZyBottom">${markItem(markedQuestion({ gradingIcon: '' }))}</div>`,
  );

  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), '');
});

test('questionLi 提取器（无 .mark_item 包裹）同样能按批阅结果回填', () => {
  const questions = questionLiExtract(
    `<div id="ZyBottom">${markedQuestion({ myAnswer: 'C' })}</div>`,
  );

  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), 'C');
  assert.equal(richContentToText(questions[0].userAnswer), 'C');
  assert.equal(questions[0].isWrong, false);
});

// ---------------------------------------------------------------------------
// 用户提供的真实页面结构原样固化（2026-10 试卷批阅结果页快照）。
//
// 与前面对照的合成夹具不同的真实类名：题干是 `h3.mark_name > span.qtContent`（题号与
// `(单选题, 0.2 分)` 是 h3 的直接文本/兄弟 span）、选项是 `ul.mark_letter.qtDetail > li`、
// 我的答案是 `.mark_key > .colorDeep.marginRight40 > span.stuAnswerContent`、
// 批阅图标是 `.mark_judge_name > span.marking_dui[aria-label="答案正确"]`。
// ---------------------------------------------------------------------------
const REAL_QUESTION = `
<div class="marBom60 questionLi scroll_884897087" id="question884897087" data="884897087">
<div class="aiArea">
<div class="aiAreaContent">
<h3 class="mark_name colorDeep">1. <span class="colorShallow">(单选题, 0.2 分)</span> <span
class="qtContent">以下各项中，所反映出的辩证法原理，与“城门失火，殃及池鱼”相同的是()。</span>
</h3>
<ul class="mark_letter colorDeep qtDetail">
<li tabindex="0" role="option">A. 揠苗助长</li>
<li tabindex="0" role="option">B. 刮目相看</li>
<li tabindex="0" role="option">C. 唇亡齿寒</li>
<li tabindex="0" role="option">D. 刻舟求剑</li>
</ul>
<div class="mark_answer">
<div class="mark_key clearfix" tabindex="0" role="option">
<span class="colorDeep marginRight40 fl"><i class="fontWeight">我的答案:</i><span
class="stuAnswerContent">C</span></span>
</div>
<div class="mark_score" tabindex="0" role="option">
<div class=" fl catalog_tishi80">
<div class="mark_judge_name">
<span class="marking_dui" aria-label="答案正确"></span>
</div>
</div>
<div class="totalScore fr"><i>0.2</i>分</div>
</div>
</div>
</div>
</div>
</div>`;

const realPage = (questionHtml) =>
  `<div id="ZyBottom"><div class="mark_item"><div class="type_tit">一. 单选题（共158题，40.0 分）</div>${questionHtml}</div></div>`;

test('真实页面结构（2026-10 快照）：题干/选项/我的答案/正确答案全链路命中', () => {
  const questions = markItemExtract(realPage(REAL_QUESTION));

  assert.equal(questions.length, 1);
  const [question] = questions;
  assert.equal(question.type, 'single-choice');
  assert.equal(question.typeMeta, '(单选题, 0.2 分)');
  // 题干取 span.qtContent，题号与题型分值都不残留
  assert.equal(
    richContentToText(question.stem),
    '以下各项中，所反映出的辩证法原理，与“城门失火，殃及池鱼”相同的是()。',
  );
  assert.deepEqual(
    question.options.map((option) => `${option.key}.${richContentToText(option.content)}`),
    ['A.揠苗助长', 'B.刮目相看', 'C.唇亡齿寒', 'D.刻舟求剑'],
  );
  // 我的答案去掉「我的答案:」标签，正确答案由 .marking_dui 回填
  assert.equal(richContentToText(question.userAnswer), 'C');
  assert.equal(richContentToText(question.correctAnswer), 'C');
  assert.equal(question.isWrong, false);
});

test('布局类名含 right 不得误判为答对（.colorDeep.marginRight40）', () => {
  // 保留真实结构，只把批阅图标换成一个正常的布局 span：我的答案仍在，
  // 但页面没给批阅结果 —— 此时绝不能凭 marginRight40 里的「Right」把 C 当正确答案。
  const ungraded = REAL_QUESTION.replace(
    '<span class="marking_dui" aria-label="答案正确"></span>',
    '<span class="fr"></span>',
  );
  const questions = markItemExtract(realPage(ungraded));

  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].userAnswer), 'C');
  assert.equal(richContentToText(questions[0].correctAnswer), '');
  assert.equal(questions[0].isWrong, false);
});

test('批阅结果只写在 aria-label 上（类名被改）时仍能识别', () => {
  const byLabel = REAL_QUESTION.replace('class="marking_dui"', 'class="mark_icon"');
  const questions = markItemExtract(realPage(byLabel));

  assert.equal(richContentToText(questions[0].correctAnswer), 'C');
});

test('真实结构的答错变体：留空且计入错题', () => {
  const wrong = REAL_QUESTION.replace(
    '<span class="marking_dui" aria-label="答案正确"></span>',
    '<span class="marking_cuo" aria-label="答案错误"></span>',
  ).replace('<span\nclass="stuAnswerContent">C</span>', '<span\nclass="stuAnswerContent">A</span>');
  const questions = markItemExtract(realPage(wrong));

  assert.equal(richContentToText(questions[0].userAnswer), 'A');
  assert.equal(richContentToText(questions[0].correctAnswer), '');
  assert.equal(questions[0].isWrong, true);
});

test('多个小题各自独立判断：对的回填、错的留空', () => {
  const questions = markItemExtract(`
    <div id="ZyBottom">
      ${markItem(
        [
          markedQuestion({
            number: 1,
            myAnswer: 'C',
            gradingIcon: '<span class="marking_dui"></span>',
          }),
          markedQuestion({
            number: 2,
            myAnswer: 'A',
            gradingIcon: '<span class="marking_cuo"></span>',
          }),
          markedQuestion({
            number: 3,
            myAnswer: 'B',
            gradingIcon: '<span class="marking_dui"></span>',
          }),
        ].join(''),
      )}
    </div>`);

  assert.deepEqual(
    questions.map((question) => richContentToText(question.correctAnswer)),
    ['C', '', 'B'],
  );
  assert.deepEqual(
    questions.map((question) => question.isWrong),
    [false, true, false],
  );
});
