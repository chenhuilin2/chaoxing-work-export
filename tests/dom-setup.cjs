// 测试用 DOM 环境。
// src/extractors 下的代码使用浏览器语义（Node.TEXT_NODE、instanceof Element 等），
// 这里把 linkedom 的构造器补到全局，使这些模块可在 Node 中按真实浏览器语义运行。
const { parseHTML, Node, Element, HTMLElement, HTMLImageElement, Event } = require('linkedom');

globalThis.Node = Node;
globalThis.Element = Element;
globalThis.HTMLElement = HTMLElement;
globalThis.HTMLImageElement = HTMLImageElement;

// 章节/任务卡定位器会派发鼠标事件并滚动到可视区，linkedom 未实现，这里补最小桩
if (!globalThis.MouseEvent) {
  globalThis.MouseEvent = class MouseEvent extends Event {
    constructor(type, init = {}) {
      super(type, { bubbles: init.bubbles, cancelable: init.cancelable });
    }
  };
}
if (typeof HTMLElement.prototype.scrollIntoView !== 'function') {
  HTMLElement.prototype.scrollIntoView = function scrollIntoView() {};
}

// delay/poll 走 window.setTimeout，部分服务（ExtractionService）直接读全局 document，
// Node 环境下补最小桩
if (!globalThis.window) globalThis.window = globalThis;
if (!globalThis.location) {
  Object.defineProperty(globalThis, 'location', {
    value: { href: 'https://mooc1.chaoxing.com/mycourse/studentstudy' },
    configurable: true,
  });
}

/** 把文档装到全局，供依赖全局 document 的服务使用 */
function useGlobalDocument(document) {
  globalThis.document = document;
  return document;
}

/** 让 iframe 元素指向给定文档，模拟同源 iframe（linkedom 不加载 iframe） */
function linkFrame(iframe, innerDocument) {
  Object.defineProperty(iframe, 'contentDocument', { get: () => innerDocument, configurable: true });
  return iframe;
}

/** 用给定的 body 片段构造一个可解析的文档 */
function createDocument(bodyHtml) {
  const { document } = parseHTML(
    `<!doctype html><html><head><title>做作业</title></head><body>${bodyHtml}</body></html>`,
  );
  if (!document.location) {
    Object.defineProperty(document, 'location', {
      value: { href: 'https://mooc1.chaoxing.com/mooc-ans/api/work?api=1' },
      configurable: true,
    });
  }
  return document;
}

/** 2026 版答题页（做作业 / 章节测验）单题结构，题干容器为 .Zy_TItle > div.fontLabel */
function renderQuestion(options = {}) {
  const {
    stemContainerOpen = '<div class="clearfix font-cxsecret fontLabel">',
    marker = '<span class="newZy_TItle">【单选题】</span>',
    index = '<i class="fl" tabindex="0" role="option">1</i>',
    wrapTitle = true,
  } = options;
  const title = `
      <div class="Zy_TItle clearfix">
        ${index}
        ${stemContainerOpen}${marker}在中国共产党历史上，第一个明确提出了“马克思主义中国化”的科学命题和重大任务的是（）。</div>
      </div>`;
  return `
  <div class="singleQuesId" id="question405842140" data="405842140">
    <div class="TiMu newTiMu" data="0">
      ${wrapTitle ? title : `<div class="TiMu_body">${index}${stemContainerOpen}在中国共产党历史上，第一个明确提出了“马克思主义中国化”的科学命题和重大任务的是（）。</div>`}
      <div class="clearfix">
        <ul class="Zy_ulTop w-top fl" tabindex="-1" aria-hidden="true">
          <li class="font-cxsecret before-after" qid="405842140" qtype="0" role="radio" aria-label="A 毛泽东选择">
            <label class="fl before"><span class="num_option choice405842140" data="A">A</span></label>
            <a class="fl after" href="javascript:void(0);">毛泽东</a>
          </li>
          <li class="font-cxsecret before-after" qid="405842140" qtype="0" role="radio" aria-label="B 李大钊选择">
            <label class="fl before"><span class="num_option choice405842140" data="B">B</span></label>
            <a class="fl after" href="javascript:void(0);">李大钊</a>
          </li>
          <li class="font-cxsecret before-after" qid="405842140" qtype="0" role="radio" aria-label="C 陈独秀选择">
            <label class="fl before"><span class="num_option choice405842140" data="C">C</span></label>
            <a class="fl after" href="javascript:void(0);">陈独秀</a>
          </li>
          <li class="font-cxsecret before-after" qid="405842140" qtype="0" role="radio" aria-label="D 周恩来选择">
            <label class="fl before"><span class="num_option choice405842140" data="D">D</span></label>
            <a class="fl after" href="javascript:void(0);">周恩来</a>
          </li>
        </ul>
        <input type="hidden" name="answer405842140" id="answer405842140" value="">
      </div>
    </div>
  </div>`;
}

/** 真实题干（已剥离题号与题型标记） */
const EXPECTED_STEM =
  '在中国共产党历史上，第一个明确提出了“马克思主义中国化”的科学命题和重大任务的是（）。';

module.exports = {
  createDocument,
  renderQuestion,
  EXPECTED_STEM,
  useGlobalDocument,
  linkFrame,
};
