// 提取诊断用例：把「提取到题干、提取不到选项」的真机现场结构带回来。
// 背景：离线快照是渲染完成后的 DOM，同一份快照离线能读全选项，复现不出真机差异；
// 因此诊断报告必须能明确区分「容器内没命中任何选项候选」与「选项在题目容器之外」。
const test = require('node:test');
const assert = require('node:assert/strict');

const { createDocument, renderQuestion } = require('./dom-setup.cjs');
const { buildExtractionDiagnostics } = require('../.tmp/test/src/extractors/dom-diagnostics.js');

test('结构正常时，报告给出候选命中数与完整的链路自检', () => {
  const document = createDocument(renderQuestion());
  // 版本号用测试专用值，避免每次发版都要改断言
  const report = buildExtractionDiagnostics({ root: document, version: 'test' });

  assert.match(report, /提取诊断 vtest /);
  assert.match(report, /\[L0\]/);
  assert.match(report, /链路自检: 提取器=timu 题数=1 选项总数=4 缺选项选择题=0/);
  // 自带候选命中，兜底不介入；候选命中数与解析数都为 4，说明选项链路正常
  assert.match(report, /\.Zy_ulTop > li=4/);
  assert.match(report, /解析出选项=4 条/);
  // 链路正常时不展开逐 li 文本，报告保持精简
  assert.doesNotMatch(report, /容器内 li 原文/);
});

test('选项节点在、但读不出文字时，报告给出「原文 vs 管道」对照（命中数>0 / 解析数=0）', () => {
  // 节点存在但无文本：命中数为 4、解析数为 0，报告必须展开并提供逐 li 对照
  const document = createDocument(`
    <div class="singleQuesId" id="question405842140" data="405842140">
      <div class="TiMu newTiMu">
        <div class="Zy_TItle clearfix">
          <div class="clearfix font-cxsecret fontLabel qtContent"><span class="newZy_TItle">【单选题】</span>题干</div>
        </div>
        <ul class="Zy_ulTop qtDetail"><li></li><li></li><li></li><li></li></ul>
      </div>
    </div>`);

  const report = buildExtractionDiagnostics({ root: document });

  assert.match(report, /\.Zy_ulTop > li=4/);
  assert.match(report, /解析出选项=0 条/);
  assert.match(report, /容器内 li 原文\(4 个\): \(空\) \| \(空\) \| \(空\) \| \(空\)/);
  assert.match(report, /容器内 li 管道: \(空\) \| \(空\) \| \(空\) \| \(空\)/);
});

test('选项被移到题目容器之外时，报告能指出「容器内为 0、外层有选项」并给出 qid 线索', () => {
  const document = createDocument(renderQuestion());
  const container = document.querySelector('.TiMu.newTiMu');
  const list = document.querySelector('.Zy_ulTop');
  // 复现「题干在 .TiMu 内、选项与 .TiMu 同级挂在 .singleQuesId 下」的结构差异
  document.querySelector('.singleQuesId').appendChild(list);

  const report = buildExtractionDiagnostics({ root: document });

  // 容器内所有候选项都为 0 —— 真机上「选项看得见却提取不到」正是这个签名
  assert.match(report, /容器内选项计数: [^\n]*\.Zy_ulTop > li=0/);
  assert.match(report, /容器内选项计数: [^\n]*li\[role="radio"\]=0/);
  assert.match(report, /解析出选项=0 条/);
  // 外层与全文档 qid 查找给出证据
  assert.match(report, /外层 \.singleQuesId: 选项计数 [^\n]*\.Zy_ulTop > li=4/);
  assert.match(report, /按 qid 全文档查找: li\[qid\]=4/);
  // 提取链路已靠外层兜底把选项找了回来
  assert.match(report, /链路自检: 提取器=timu 题数=1 选项总数=4 缺选项选择题=0/);
  assert.ok(container.ownerDocument.contains(list));
});

test('文档里没有题目时报告不抛错，且明确说明未提取到题目', () => {
  const report = buildExtractionDiagnostics({ root: createDocument('<div>空空如也</div>') });

  assert.match(report, /链路自检: 该文档未提取到题目/);
  assert.match(report, /题目容器=/);
});
