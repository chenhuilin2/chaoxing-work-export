const test = require('node:test');
const assert = require('node:assert/strict');
require('./dom-setup.cjs');
const { resolvePageTitle } = require('../.tmp/test/src/extractors/page-context.js');

/** 用给定的 body 片段构造文档（含默认 document.title） */
function createDocument(bodyHtml, docTitle = '做作业') {
  const { parseHTML } = require('linkedom');
  const { document } = parseHTML(
    `<!doctype html><html><head><title>${docTitle}</title></head><body>${bodyHtml}</body></html>`,
  );
  if (!document.location) {
    Object.defineProperty(document, 'location', {
      value: { href: 'https://mooc1.chaoxing.com/mooc-ans/api/work?api=1' },
      configurable: true,
    });
  }
  return document;
}

test('学生学习页面：标题优先取自 #prev_title 的 title 属性', () => {
  const document = createDocument('<div id="prev_title" title="第一章 章节测验">章节小测</div>');
  assert.equal(resolvePageTitle(document), '第一章 章节测验');
});

test('学生学习页面：#prev_title 无 title 属性时回退到其正文', () => {
  const document = createDocument('<div id="prev_title">第一章 章节测验</div>');
  assert.equal(resolvePageTitle(document), '第一章 章节测验');
});

test('学生学习页面：真实 DOM 的 .prev_title 类名同样生效', () => {
  // html/章节练习完整页.html：<div class="prev_title" title="毛泽东思想的主要内容">
  const document = createDocument(
    '<div class="prev_title_pos"><div class="prev_title" title="毛泽东思想的主要内容">毛泽东思想的主要内容<span class="markTag"></span></div></div>',
  );
  assert.equal(resolvePageTitle(document), '毛泽东思想的主要内容');
});

test('无 prev_title 时仍走原有选择器与 document.title 兜底', () => {
  const document = createDocument('<div class="mark_title">期末作业</div>');
  assert.equal(resolvePageTitle(document), '期末作业');

  const fallback = createDocument('<div class="TiMu">题干</div>');
  assert.equal(resolvePageTitle(fallback), '做作业');
});
