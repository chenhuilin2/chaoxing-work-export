// 选项解析用例：学习通改版换掉选项类名时，靠「语义兜底」仍要能取到选项。
// 背景：答题页每个选项是 <li class="font-cxsecret before-after" role="radio" onclick="addChoice(this)">，
// 类名随版本变化，但 role 与 onclick 极稳定，因此 parseOptions 会在各提取器自带候选之后追加兜底选择器。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument, renderQuestion } = require('./dom-setup.cjs');
const {
  CompositeExtractor,
} = require('../.tmp/test/src/extractors/composite-extractor.js');
const { parseOptions } = require('../.tmp/test/src/extractors/option-parser.js');
const {
  incompleteChoiceCount,
} = require('../.tmp/test/src/application/extraction-service.js');
const { richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');

/** 把选项列表的类名换掉，但保留 role="radio" */
function keepRoleOnly() {
  return renderQuestion().replace('class="Zy_ulTop w-top fl"', 'class="opt-wrap"');
}

/** 连 role 也去掉，只留 onclick="addChoice(this)" */
function keepOnclickOnly() {
  return renderQuestion()
    .replace('class="Zy_ulTop w-top fl"', 'class="opt-wrap"')
    .replace(
      /class="font-cxsecret before-after" qid="405842140" qtype="0" role="radio" aria-label="[^"]*"/g,
      'class="opt-item" onclick="addChoice(this);"',
    );
}

function extractOptions(html) {
  const document = createDocument(html);
  const container = document.querySelector('.TiMu.newTiMu');
  return {
    parsed: parseOptions(container, ['.Zy_ulTop.qtDetail > li', '.Zy_ulTop > li', '.answerBg', '.option-list > li']),
    question: new CompositeExtractor().extract(document)?.questions[0] ?? null,
  };
}

test('类名改版但保留 role="radio" 时仍能取到选项', () => {
  const html = keepRoleOnly();
  const { parsed, question } = extractOptions(html);

  assert.equal(parsed.length, 4);
  assert.equal(parsed.map((option) => option.key).join(''), 'ABCD');
  assert.equal(richContentToText(parsed[0].content), '毛泽东');
  assert.equal(question?.options.length, 4);
});

test('连 role 都没有、只剩 onclick="addChoice(this)" 时仍能取到选项', () => {
  const html = keepOnclickOnly();
  const { parsed, question } = extractOptions(html);

  assert.equal(parsed.length, 4);
  assert.equal(richContentToText(parsed[3].content), '周恩来');
  assert.equal(question?.options.length, 4);
});

test('结构正常时行为不变：自带候选先生效，兜底不介入', () => {
  const { parsed, question } = extractOptions(renderQuestion());

  assert.equal(parsed.length, 4);
  assert.deepEqual(
    parsed.map((option) => option.key),
    ['A', 'B', 'C', 'D'],
  );
  assert.equal(question?.options.length, 4);
});

test('incompleteChoiceCount 能识别「只渲染了题干」的选择题', () => {
  const withOptions = new CompositeExtractor().extract(createDocument(renderQuestion()));
  assert.equal(incompleteChoiceCount(withOptions), 0);

  const stemOnly = createDocument(renderQuestion());
  stemOnly.querySelector('.Zy_ulTop').remove();
  const stripped = new CompositeExtractor().extract(stemOnly);
  assert.equal(stripped?.questions[0].options.length, 0);
  assert.equal(incompleteChoiceCount(stripped), 1);

  // 非选择题没有选项属正常，不应被判定为「未渲染完」
  const fillBlank = createDocument(`
    <div class="TiMu newTiMu">
      <div class="Zy_TItle clearfix">
        <i class="fl" role="option">1</i>
        <div class="clearfix fontLabel"><span class="newZy_TItle">【填空题】</span>第一个提出“马克思主义中国化”的人是____。</div>
      </div>
      <div class="clearfix"><input class="blankItemInp" type="text"></div>
    </div>`);
  const blankResult = new CompositeExtractor().extract(fillBlank);
  assert.equal(blankResult?.questions.length, 1);
  assert.equal(incompleteChoiceCount(blankResult), 0);
});

// 已批阅视图（章节测验「已完成」）的部分版本把选项放在题目容器之外（同级表单 / 外层单题块），
// 此时容器内所有选项候选都会落空 —— 即使选项明明显示在页面上。以下用例覆盖向外层兜底。
test('选项被移出题目容器、但仍在该题的单题块内时，仍能取到选项', () => {
  const document = createDocument(renderQuestion());
  const block = document.querySelector('.singleQuesId');
  block.appendChild(document.querySelector('.Zy_ulTop'));

  const question = new CompositeExtractor().extract(document)?.questions[0];

  assert.equal(question?.options.length, 4);
  assert.deepEqual(
    question?.options.map((option) => option.key),
    ['A', 'B', 'C', 'D'],
  );
});

test('选项被移出单题块时，靠页面标的 qid 精确找回，不串到邻题', () => {
  // 第二题换掉题干正文，避免两题题干相同被指纹去重
  const html =
    renderQuestion() +
    renderQuestion()
      .replace(/405842140/g, '405842141')
      .replace('在中国共产党历史上', '在中国国民党历史上');
  const document = createDocument(html);
  const blocks = [...document.querySelectorAll('.singleQuesId')];
  // 第一题的选项挪到页面最外层（既不在 .TiMu 内，也不在它自己的单题块内）
  document.body.appendChild(blocks[0].querySelector('.Zy_ulTop'));

  const questions = new CompositeExtractor().extract(document)?.questions ?? [];

  assert.equal(questions.length, 2);
  assert.equal(questions[0].options.length, 4);
  assert.equal(questions[1].options.length, 4);
  // 两题的选项各自独立，没有被合并成一份
  assert.equal(richContentToText(questions[0].options[0].content), '毛泽东');
});

test('容器外兜底不猜：页面没有任何 qid 线索时不会乱取别的选项', () => {
  // 去掉所有 4 位以上的数字（qid、单题块 id、答案输入框 id），使兜底无从定位
  const document = createDocument(renderQuestion().replace(/\d{4,}/gu, ''));
  document.body.appendChild(document.querySelector('.Zy_ulTop'));

  const result = new CompositeExtractor().extract(document);

  assert.equal(result?.questions.length, 1);
  assert.equal(result.questions[0].options.length, 0);
});
