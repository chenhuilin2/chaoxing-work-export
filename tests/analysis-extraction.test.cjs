// 作业页（.mark_item .questionLi）的「答案解析」提取。
//
// 学习通作业页把解析渲染成「标签 + 正文」同层的结构：
//
//   <div class="analysisDiv optionAnalysisDisplayMarker">
//     <div class="analysis">
//       <i class="custom-style">答案解析：</i>
//       <span class="qtAnalysis workTextWrap"><p>…</p></span>
//     </div>
//   </div>
//
// 这段 HTML 有两个坑：`.analysis` 可能挂在 `.mark_answer` 里面（旧模板），也可能是它的兄弟节点
// （作业页常见），所以选择器要同时覆盖「`.mark_answer .analysis`」与「`.analysisDiv .qtAnalysis`」；
// 另外标签「答案解析：」必须整段剥掉 —— 正则里「答案」排在「答案解析」前面时会留下「解析：」，
// 导出就成了「解析：解析：…」。
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDocument } = require('./dom-setup.cjs');
const { MarkItemExtractor } = require('../.tmp/test/src/extractors/mark-item-extractor.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

const LETTERS = ['A', 'B', 'C', 'D'];

/**
 * 渲染一道作业页题目。
 * @param {object} options
 * @param {string} [options.analysis] 解析正文；空串表示这题没有解析
 * @param {boolean} [options.analysisInsideAnswer] 解析是否挂在 `.mark_answer` 里面（旧模板）
 */
function homeworkQuestion(options = {}) {
  const {
    number = 1,
    stem = '从本质上讲，信息安全管理是一个怎样的过程？',
    optionTexts = ['一次性部署防火墙', '持续性管理风险的过程', '定期漏洞扫描', '仅由 IT 部门负责'],
    correctAnswer = 'B',
    analysis = '密码学理论是信息安全的基石。',
    analysisInsideAnswer = false,
  } = options;

  const optionList = `<ul class="mark_letter colorDeep qtDetail">${optionTexts
    .map((text, index) => `<li class="workTextWrap">${LETTERS[index]}. <p>${text}</p></li>`)
    .join('')}</ul>`;

  const analysisHtml = analysis
    ? `<div class="analysisDiv optionAnalysisDisplayMarker">
         <div class="analysis"><i class="custom-style">答案解析：</i><span class="qtAnalysis workTextWrap"><p>${analysis}</p></span></div>
       </div>`
    : '';

  return `
    <div class="questionLi singleQuesId" typeName="单选题">
      <h3 class="mark_name colorDeep">${number}. <span class="colorShallow">(单选题)</span><span class="qtContent workTextWrap"><p>${stem}</p></span></h3>
      ${optionList}
      <div class="mark_answer">
        <div class="mark_key clearfix">
          <span class="colorDeep"><i class="fontWeight custom-style">我的答案:</i><span class="stuAnswerContent">${correctAnswer}</span></span>
          <span class="colorGreen"><i class="fontWeight custom-style">正确答案:</i><span class="rightAnswerContent">${correctAnswer}</span></span>
        </div>
        ${analysisInsideAnswer ? analysisHtml : ''}
      </div>
      ${analysisInsideAnswer ? '' : analysisHtml}
    </div>`;
}

function extract(bodyHtml) {
  return new MarkItemExtractor().extract({
    root: createDocument(
      `<div class="mark_item"><div class="type_tit">一. 单选题（共2题）</div>${bodyHtml}</div>`,
    ),
    title: '第1章作业',
    pageUrl: 'https://mooc1.chaoxing.com/mooc-ans/api/work?api=1',
  });
}

test('作业页：解析在 .mark_answer 之外时也能提取，且不带「答案解析：」标签', () => {
  const [question] = extract(homeworkQuestion({ analysis: '密码学理论是信息安全的基石。' }));

  assert.equal(richContentToText(question.analysis), '密码学理论是信息安全的基石。');
  assert.equal(richContentToText(question.correctAnswer), 'B');
});

test('作业页：解析挂在 .mark_answer 里面（旧模板）同样提取且剥掉标签', () => {
  const [question] = extract(
    homeworkQuestion({ analysis: '老模板里的解析。', analysisInsideAnswer: true }),
  );

  assert.equal(richContentToText(question.analysis), '老模板里的解析。');
});

test('作业页：没有解析的题目 analysis 为空数组（导出时不该凭空多一行）', () => {
  const questions = extract(
    homeworkQuestion({ number: 1, analysis: '有解析' }) +
      homeworkQuestion({ number: 2, stem: '没有解析的题', analysis: '' }),
  );

  assert.equal(questions.length, 2);
  assert.equal(richContentToText(questions[0].analysis), '有解析');
  assert.deepEqual(questions[1].analysis, []);
});
