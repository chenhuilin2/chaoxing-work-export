// 简答题（含多段文字答案）的答案提取。
//
// 学习通把答案区的「标签」与「正文」渲染成两个**同级同名**容器，简答题最典型：
//   <div class="correctAnswerBx">
//     <div class="correctAnswer">正确答案：</div>            ← 第一个：只有标签，内容为空
//     <div class="correctAnswer marTop6">五四运动是…</div>    ← 第二个：真正的答案
//   </div>
// 旧实现「取第一个匹配元素」永远读到空标签，导出的答案汇总对简答题整页是「（未找到答案）」。
// 现改为「同一选择器下取第一个非空匹配」。
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDocument } = require('./dom-setup.cjs');
const { TiMuExtractor } = require('../.tmp/test/src/extractors/timu-extractor.js');
const {
  extractCorrectAnswer,
  extractUserAnswer,
} = require('../.tmp/test/src/extractors/common-answer.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

// 用户提供的真实页面结构（2026-10 章节测验「第四章」，fixture 见 tmp.html 第 57 题），
// 类名原样保留：题干 .fontLabel + span.newZy_TItle、【我的答案】双容器、正确答案双容器。
const ESSAY_QUESTION = `
<div class="TiMu newTiMu ans-cc singleQuesId" data="214729266">
  <div class="Zy_TItle clearfix">
    <i class="fl" tabindex="0" role="option">57</i>
    <div class="clearfix font-cxsecret fontLabel"
      style="line-height: 1.5; font-size: 14px;padding-right:15px;" tabindex="0"
      role="option"><span class="newZy_TItle">【简答题】</span>
      <p>请简述五四运动的历史意义。</p>
    </div>
  </div>
  <div class="newAnswerBx">
    <div class="myAnswerBx marBot16">
      <div class="myAnswer marBot16">
        <span class="answerFont fl">我的答案：</span>
        <p class="clear"></p>
      </div>
    </div>
    <div class="myAllAnswerBx">
      <div class="myAnswerBx marBot16">
        <div class="myAnswer marTop16">
          <p>五四爱国运动，是一次彻底的反对帝国主义封建主义的爱国运动，是中国新民主主义革命的开始。</p>
        </div>
      </div>
    </div>
    <div class="correctAnswerBx marBot16">
      <div class="correctAnswer">
        <span class="answerFont fl">正确答案：</span>
        <p class="clear"></p>
      </div>
      <div class="correctAnswer marTop6">
        <span style="word-break: break-all;display: inline-block;">
          <p>五四运动是中国旧民主主义革命走向新民主主义革命的转折点，促进了马克思主义在中国的传播。
            <br>五四运动孕育了以爱国、进步、民主、科学为主要内容的伟大五四精神，其核心是爱国主义。
            <br>五四运动以全民族的行动激发了追求真理、追求进步的伟大觉醒。</p>
        </span>
      </div>
    </div>
  </div>
</div>`;

/** 选择题：答案区只有一个容器，标签与内容在同一节点里（改造不能影响这条路） */
function singleChoiceQuestion() {
  return `
<div class="TiMu newTiMu ans-cc singleQuesId" data="214729128">
  <div class="Zy_TItle clearfix">
    <i class="fl" tabindex="0" role="option">1</i>
    <div class="clearfix font-cxsecret fontLabel" tabindex="0" role="option">
      <span class="newZy_TItle">【单选题】</span>题干一
    </div>
  </div>
  <ul class="Zy_ulTop" tabindex="-1" aria-hidden="true">
    <li class="clearfix" role="option"><i class="fl">A、</i><a href="javascript:void(0);" class="fl font-cxsecret">甲</a></li>
    <li class="clearfix" role="option"><i class="fl">B、</i><a href="javascript:void(0);" class="fl font-cxsecret">乙</a></li>
  </ul>
  <div class="newAnswerBx">
    <div class="myAnswerBx marBot16">
      <div class="myAnswer"><span class="answerFont fl">我的答案：</span><div class="fl answerCon">C</div></div>
      <div class="answerScore"><div class="CorrectOrNot fl"><span class="marking_dui"></span></div></div>
    </div>
    <div class="correctAnswerBx marBot16">
      <div class="correctAnswer"><span class="answerFont fl">正确答案：</span><div class="fl answerCon">C</div></div>
    </div>
  </div>
</div>`;
}

const page = (inner) => `<div id="ZyBottom">${inner}</div>`;

function extractTimu(html) {
  return new TiMuExtractor().extract({
    root: createDocument(html),
    title: '第四章',
    pageUrl: 'https://mooc1.chaoxing.com/mooc-ans/work/doHomeWorkNew',
  });
}

test('简答题：正确答案取「第二个 .correctAnswer」（标签容器为空时不误取）', () => {
  const questions = extractTimu(page(ESSAY_QUESTION));

  assert.equal(questions.length, 1);
  const [question] = questions;
  assert.equal(question.type, 'short-answer');
  assert.equal(richContentToText(question.stem), '请简述五四运动的历史意义。');

  const answer = richContentToText(question.correctAnswer);
  assert.ok(answer.length > 60, `答案过短：${answer}`);
  assert.match(answer, /^五四运动是中国旧民主主义革命走向新民主主义革命的转折点/);
  assert.match(answer, /五四精神/);
  // 没有把「正确答案：」标签一起带出来
  assert.doesNotMatch(answer, /正确答案/);
  // 多个 <br> / <p> 换成多行，保留原始分段
  assert.ok(answer.split('\n').length >= 2, `未保留换行：${JSON.stringify(answer)}`);
});

test('简答题：我的答案取 myAllAnswerBx 内的作答内容', () => {
  const questions = extractTimu(page(ESSAY_QUESTION));
  const mine = richContentToText(questions[0].userAnswer);

  assert.match(mine, /^五四爱国运动/);
  assert.match(mine, /新民主主义革命的开始/);
});

test('简答题不算错题（无标准答案比对语义），也不因未批阅而误判', () => {
  const questions = extractTimu(page(ESSAY_QUESTION));

  assert.equal(questions[0].isWrong, false);
});

test('回归：选择题答案仍取第一个容器里的 .answerCon（不受「取非空」改动影响）', () => {
  const questions = extractTimu(page(singleChoiceQuestion()));

  assert.equal(questions.length, 1);
  assert.equal(questions[0].type, 'single-choice');
  assert.equal(richContentToText(questions[0].correctAnswer), 'C');
  assert.equal(richContentToText(questions[0].userAnswer), 'C');
});

test('extractCorrectAnswer：同一选择器下的多个匹配，跳过空标签取第一个非空', () => {
  const doc = createDocument(`
    <div class="TiMu newTiMu" data="1">
      <div class="newAnswerBx">
        <div class="correctAnswerBx">
          <div class="correctAnswer"><span class="answerFont fl">正确答案：</span></div>
          <div class="correctAnswer">真正的答案内容</div>
        </div>
      </div>
    </div>`);
  const container = doc.querySelector('.TiMu');

  assert.equal(richContentToText(extractCorrectAnswer(container)), '真正的答案内容');
});

test('extractCorrectAnswer：所有候选都为空时返回空，不会造出假答案', () => {
  const doc = createDocument(`
    <div class="TiMu newTiMu" data="1">
      <div class="newAnswerBx">
        <div class="correctAnswerBx">
          <div class="correctAnswer"><span class="answerFont fl">正确答案：</span></div>
        </div>
      </div>
    </div>`);
  const container = doc.querySelector('.TiMu');

  assert.equal(richContentToText(extractCorrectAnswer(container)), '');
});

test('extractUserAnswer：同样跳过空的标签容器', () => {
  const doc = createDocument(`
    <div class="TiMu newTiMu" data="1">
      <div class="newAnswerBx">
        <div class="myAnswerBx marBot16">
          <div class="myAnswer marBot16"><span class="answerFont fl">我的答案：</span><p class="clear"></p></div>
        </div>
        <div class="myAllAnswerBx">
          <div class="myAnswerBx marBot16">
            <div class="myAnswer marTop16"><p>我写了一大段</p></div>
          </div>
        </div>
      </div>
    </div>`);
  const container = doc.querySelector('.TiMu');
  const mine = richContentToText(extractUserAnswer(container));

  assert.match(mine, /我写了一大段/);
  assert.doesNotMatch(mine, /我的答案/);
});

test('多种题型混排：简答题与选择题各自取到答案，互不串题', () => {
  const questions = extractTimu(page(singleChoiceQuestion() + ESSAY_QUESTION));

  assert.equal(questions.length, 2);
  assert.deepEqual(
    questions.map((question) => question.type),
    ['single-choice', 'short-answer'],
  );
  assert.equal(richContentToText(questions[0].correctAnswer), 'C');
  assert.ok(richContentToText(questions[1].correctAnswer).length > 60);
});
