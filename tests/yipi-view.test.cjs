// 章节测验「已完成 / 已批阅」视图（selectWorkQuestionYiPiYue 模板）用例。
//
// 该模板与答题页的差异（依据 html/dom.html 真实快照）：
// 1. 题干容器多一个 .qtContent 类，选项在 <form> 内的 ul.Zy_ulTop.qtDetail，li 是 li.clearfix + i/a；
// 2. 分组标题 h3.newTestType 是 .aiArea 的**前一个兄弟节点**，不是后代；
// 3. 不输出「正确答案」文本，只用 span.marking_dui / span.marking_cuo 表示批阅结果。
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDocument, withForeignConstructors } = require('./dom-setup.cjs');
const { TiMuExtractor } = require('../.tmp/test/src/extractors/timu-extractor.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

/** 渲染一道已批阅视图的题目（marker 传空串即为「题干里没有题型标记」的场景） */
function gradedQuestion(options = {}) {
  const {
    id = '405842706',
    index = '1',
    marker = '<span class="newZy_TItle">【单选题】</span>',
    stem = '（）标志着“马克思主义的中国化”这一命题的正式提出。',
    options: optionTexts = ['党的七大', '党的五大', '党的七届二中全会', '党的六届六中全会'],
    myAnswer = 'D',
    marking = 'marking_dui',
    correctAnswer = '',
    analysis = '',
    ulClass = 'Zy_ulTop qtDetail',
  } = options;
  const letters = ['A', 'B', 'C', 'D'];

  const optionList =
    optionTexts.length === 0
      ? ''
      : `<form action="javascript:;" method="post"><ul class="${ulClass}" tabindex="-1" aria-hidden="true">${optionTexts
          .map(
            (text, position) =>
              `<li class="clearfix" tabindex="0" role="option"><i class="fl">${letters[position]}、</i><a href="javascript:void(0)" class="fl">${text}</a></li>`,
          )
          .join('')}</ul></form>`;

  return `
    <div class="aiArea"><div class="aiAreaContent">
      <div class="TiMu newTiMu ans-cc singleQuesId" data="${id}" id="question${id}">
        <div class="Zy_TItle clearfix">
          <i class="fl" tabindex="0" role="option">${index}</i>
          <div class="clearfix font-cxsecret fontLabel qtContent">${marker}${stem}</div>
        </div>
        ${optionList}
        <div class="newAnswerBx">
          <div class="myAnswerBx marBot16">
            <div class="myAnswer">
              <span class="answerFont fl">我的答案：</span>
              <div class="fl answerCon">${myAnswer}</div>
            </div>
            <div class="answerScore"><div class="CorrectOrNot fl"><span class="${marking}"></span></div></div>
          </div>
          ${correctAnswer ? `<div class="correctAnswer"><div class="answerCon">${correctAnswer}</div></div>` : ''}
          ${analysis ? `<div class="answerKeyBx"><div class="answerCon">${analysis}</div></div>` : ''}
        </div>
      </div>
    </div></div>`;
}

function extract(bodyHtml) {
  return new TiMuExtractor().extract({
    root: createDocument(bodyHtml),
    title: '导论1',
    pageUrl: 'https://mooc1.chaoxing.com/mooc-ans/work/selectWorkQuestionYiPiYue',
  });
}

test('已批阅视图：题干用 .qtContent、选项在 form 内的 ul.Zy_ulTop 时仍能取到选项', () => {
  const questions = extract(`<div id="ZyBottom">${gradedQuestion()}</div>`);
  assert.equal(questions.length, 1);
  const [question] = questions;
  assert.equal(question.type, 'single-choice');
  // 题号节点是裸数字 <i class="fl">1</i>（无「.」「、」），必须以「整串就是数字」识别出来
  assert.equal(question.number, 1);
  assert.deepEqual(
    question.options.map((option) => `${option.key}.${richContentToText(option.content)}`),
    ['A.党的七大', 'B.党的五大', 'C.党的七届二中全会', 'D.党的六届六中全会'],
  );
  assert.equal(
    richContentToText(question.stem),
    '（）标志着“马克思主义的中国化”这一命题的正式提出。',
  );
});

test('已批阅视图：选项容器类名全部改版时，靠 li[role="option"] 兜底仍能取到选项', () => {
  const questions = extract(`
    <div id="ZyBottom">
      ${gradedQuestion({ id: '1', ulClass: 'opt-wrap' })}
    </div>`);
  assert.equal(questions.length, 1);
  assert.deepEqual(
    questions[0].options.map((option) => `${option.key}.${richContentToText(option.content)}`),
    ['A.党的七大', 'B.党的五大', 'C.党的七届二中全会', 'D.党的六届六中全会'],
  );
});

test('已批阅视图：marking_cuo（教师批阅错误）标记为错题，marking_dui 不算错题', () => {
  const wrong = extract(`
    <div id="ZyBottom">
      ${gradedQuestion({ id: '1', marking: 'marking_cuo' })}
      ${gradedQuestion({ id: '2', marking: 'marking_dui' })}
    </div>`);
  assert.equal(wrong.length, 2);
  assert.equal(wrong[0].isWrong, true);
  assert.equal(wrong[1].isWrong, false);
});

test('已批阅视图：分组标题 h3.newTestType 是兄弟节点时题型可继承（题干无标记也能识别）', () => {
  const questions = extract(`
    <div id="ZyBottom">
      <h3 class="newTestType">一. 单选题（共1题）</h3>
      ${gradedQuestion({ id: '1', marker: '' })}
      <h3 class="newTestType">二. 填空题（共1题）</h3>
      ${gradedQuestion({ id: '2', marker: '', stem: '马克思主义中国化的第一次历史性飞跃的理论成果是____。', options: [], myAnswer: '毛泽东思想' })}
    </div>`);
  assert.equal(questions.length, 2);
  assert.equal(questions[0].type, 'single-choice');
  assert.equal(questions[1].type, 'fill-blank');
  assert.equal(richContentToText(questions[1].userAnswer), '毛泽东思想');
});

test('已批阅视图：correctAnswer / answerKeyBx 两种答案解析容器都能识别', () => {
  const questions = extract(`
    <div id="ZyBottom">
      ${gradedQuestion({
        id: '1',
        correctAnswer: 'D',
        analysis: '党的六届六中全会上毛泽东正式提出“马克思主义中国化”命题。',
      })}
    </div>`);
  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), 'D');
  assert.equal(
    richContentToText(questions[0].analysis),
    '党的六届六中全会上毛泽东正式提出“马克思主义中国化”命题。',
  );
});

test('已批阅视图：批阅正确时用我的答案回填正确答案', () => {
  const questions = extract(`
    <div id="ZyBottom">
      ${gradedQuestion({ id: '1', myAnswer: 'D', marking: 'marking_dui' })}
    </div>`);
  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), 'D');
  assert.equal(richContentToText(questions[0].userAnswer), 'D');
  assert.equal(questions[0].isWrong, false);
});

test('已批阅视图：批阅错误或部分正确时不算正确答案，且都计入错题', () => {
  const questions = extract(`
    <div id="ZyBottom">
      ${gradedQuestion({ id: '1', myAnswer: 'A', marking: 'marking_cuo' })}
      ${gradedQuestion({ id: '2', myAnswer: 'B', marking: 'marking_bandui' })}
    </div>`);
  assert.equal(questions.length, 2);
  assert.equal(richContentToText(questions[0].correctAnswer), '');
  assert.equal(richContentToText(questions[0].userAnswer), 'A');
  assert.equal(questions[0].isWrong, true);
  assert.equal(richContentToText(questions[1].correctAnswer), '');
  assert.equal(questions[1].isWrong, true);
});

test('页面本身给了正确答案时不回填，保持页面答案', () => {
  const questions = extract(`
    <div id="ZyBottom">
      ${gradedQuestion({ id: '1', myAnswer: 'D', marking: 'marking_dui', correctAnswer: 'C' })}
    </div>`);
  assert.equal(questions.length, 1);
  assert.equal(richContentToText(questions[0].correctAnswer), 'C');
});

test('已批阅视图：顶层脚本跨层读 iframe 文档（节点来自另一个 window）时选项仍完整', () => {
  // 该模板的选项文字一律嵌在 <a> 里，一旦遍历进不去子元素就会「只有题干、没有选项」；
  // 真机现场即如此（诊断显示 li 命中 4 个、解析出 0 条），这里守住跨 window 读取
  const questions = withForeignConstructors(() =>
    extract(`<div id="ZyBottom">${gradedQuestion()}</div>`),
  );

  assert.equal(questions.length, 1);
  assert.deepEqual(
    questions[0].options.map((option) => `${option.key}.${richContentToText(option.content)}`),
    ['A.党的七大', 'B.党的五大', 'C.党的七届二中全会', 'D.党的六届六中全会'],
  );
  assert.equal(
    richContentToText(questions[0].stem),
    '（）标志着“马克思主义的中国化”这一命题的正式提出。',
  );
});
