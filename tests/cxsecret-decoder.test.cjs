// font-cxsecret 字体反爬解码的单元测试。
// 夹具字体 tests/fixtures/cxsecret-quiz.ttf 取自真实「章节测验」页面内联的 @font-face，
// 其混淆码位与真值的对应关系已用两条独立方法（字形哈希表 / 栅格形状匹配）交叉验证过。
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { createDocument, EXPECTED_STEM } = require('./dom-setup.cjs');
const {
  sha1Prefix32,
  decodeBase64,
  getGlyphTable,
  buildGlyphDecodeMap,
  replaceByCodePointMap,
} = require('../.tmp/test/src/extractors/cxsecret-font.js');
const {
  getCxSecretDecoder,
  isCxSecretScope,
  decodeCxSecretText,
} = require('../.tmp/test/src/extractors/cxsecret-decoder.js');
const { extractRichContent, richContentToText } = require('../.tmp/test/src/extractors/rich-content.js');
const { TiMuExtractor } = require('../.tmp/test/src/extractors/timu-extractor.js');

const FONT_BASE64 = fs.readFileSync(path.join(__dirname, 'fixtures', 'cxsecret-quiz.ttf')).toString('base64');
const FONT_STYLE = `<style>@font-face { font-family: 'font-cxsecret'; src: url('data:application/font-ttf;charset=utf-8;base64,${FONT_BASE64}') format('truetype'); } .font-cxsecret, .font-cxsecret p, .font-cxsecret div, .font-cxsecret i, .font-cxsecret em { font-family: 'font-cxsecret'; }</style>`;

// 真实页面上第一题的原文与真值（见 docs/chapter-quiz-diagnosis.md）
const STEM_SCRAMBLED =
  '【单选题】在中悢共产党悤史上，第一悡明悞提出了“马克悝主义中悢化”的悧学命题和重大悥务的是（）。';
const OPTIONS_SCRAMBLED = ['愨泽悩', '悮大慳', '悱独悰', '悴恩悳'];
const OPTIONS_DECODED = ['毛泽东', '李大钊', '陈独秀', '周恩来'];
// 该字体里 U+53C3(參) 是 U+4E2D(中) 的混淆码位，用于验证「作用域外不改动」
const DECOY = '參';
const DECOY_REAL = '中';

/** 真实答题页单题结构（题干与选项均处于 font-cxsecret 作用域内） */
function renderScrambledQuestion() {
  const options = OPTIONS_SCRAMBLED.map(
    (text, index) => `
      <li class="font-cxsecret before-after" qid="405842140" qtype="0" tabindex="0" role="radio" aria-label="${'ABCD'[index]} ${text}选择">
        <label class="fl before"><span class="num_option choice405842140" data="${'ABCD'[index]}">${'ABCD'[index]}</span></label>
        <a tabindex="-1" href="javascript:void(0);" class="fl after">${text}</a>
      </li>`,
  ).join('');
  return `
    <div class="singleQuesId" id="question405842140">
      <div class="TiMu newTiMu" data="0">
        <div class="Zy_TItle clearfix">
          <i class="fl" tabindex="0" role="option">1</i>
          <div class="clearfix font-cxsecret fontLabel" tabindex="0" role="option">
            <span class="newZy_TItle">【单选题】</span>在中悢共产党悤史上，第一悡明悞提出了“马克悝主义中悢化”的悧学命题和重大悥务的是（）。
          </div>
        </div>
        <ul class="Zy_ulTop w-top fl" tabindex="-1" aria-hidden="true">${options}</ul>
      </div>
    </div>`;
}

function context(document) {
  return {
    root: document,
    title: '章节测验',
    pageUrl: 'https://mooc1.chaoxing.com/mooc-ans/api/work?api=1',
  };
}

test('SHA-1 首 4 字节与 node crypto 一致（含各种分块边界）', () => {
  for (const size of [0, 1, 54, 55, 56, 63, 64, 65, 119, 120, 128, 1000]) {
    const buffer = Buffer.alloc(size);
    for (let i = 0; i < size; i += 1) buffer[i] = (i * 37 + size) & 0xff;
    const expected = crypto.createHash('sha1').update(buffer).digest().readUInt32BE(0);
    assert.equal(sha1Prefix32(new Uint8Array(buffer)), expected, `长度 ${size} 不一致`);
  }
});

test('内置解码表可解析，条目数与生成器一致', () => {
  const table = getGlyphTable();
  assert.notEqual(table, null);
  assert.equal(table.length / 6, 20902);
});

test('base64 解码与 node Buffer 一致（含换行与填充）', () => {
  const raw = fs.readFileSync(path.join(__dirname, 'fixtures', 'cxsecret-quiz.ttf'));
  assert.deepEqual(Buffer.from(decodeBase64(FONT_BASE64)), raw);
  assert.deepEqual(Buffer.from(decodeBase64(`  ${FONT_BASE64.slice(0, 1000)}\n`)), raw.subarray(0, 750));
});

test('从真实页面字体内解出混淆码位映射，且映射随页面变化', () => {
  const bytes = new Uint8Array(fs.readFileSync(path.join(__dirname, 'fixtures', 'cxsecret-quiz.ttf')));
  const map = buildGlyphDecodeMap(bytes);

  assert.equal(map.size, 83);
  assert.equal(map.get(0x53c3), '中');
  assert.equal(map.get(0x51c4), '误');
  // 混淆码位对应的真字不等于码位本身，这正是「反爬」的表现
  assert.notEqual(map.get(0x53c3), String.fromCharCode(0x53c3));
});

test('解码文档：题干与选项的混淆文字被还原，作用域外文字不动', () => {
  const document = createDocument(
    `${FONT_STYLE}
     <div id="stem" class="clearfix font-cxsecret fontLabel">${STEM_SCRAMBLED}</div>
     <div id="plain" class="plain">${DECOY}</div>
     <div id="scoped" class="font-cxsecret">${DECOY}</div>`,
  );
  const decoder = getCxSecretDecoder(document);
  assert.notEqual(decoder, null);

  const stem = document.getElementById('stem');
  assert.equal(isCxSecretScope(stem), true);
  assert.equal(decodeCxSecretText(STEM_SCRAMBLED, stem), '【单选题】' + EXPECTED_STEM);

  // 同一码位：作用域内还原成真字，作用域外保持原样（避免改坏正常内容）
  assert.equal(decodeCxSecretText(DECOY, document.getElementById('scoped')), DECOY_REAL);
  assert.equal(decodeCxSecretText(DECOY, document.getElementById('plain')), DECOY);
});

test('页面无内联反爬字体时解码器为空，文本原样返回', () => {
  const document = createDocument(`<div class="font-cxsecret fontLabel">${STEM_SCRAMBLED}</div>`);

  assert.equal(getCxSecretDecoder(document), null);
  assert.equal(decodeCxSecretText(STEM_SCRAMBLED, document.querySelector('.fontLabel')), STEM_SCRAMBLED);
  assert.equal(decodeCxSecretText(STEM_SCRAMBLED, null), STEM_SCRAMBLED);
});

test('内联字体数据损坏时不抛错，退化为不解码', () => {
  const document = createDocument(
    `<style>@font-face { font-family: 'font-cxsecret'; src: url('data:application/font-ttf;base64,AAAAAAA=') format('truetype'); }</style>
     <div class="font-cxsecret">${STEM_SCRAMBLED}</div>`,
  );

  assert.equal(getCxSecretDecoder(document), null);
  assert.equal(
    richContentToText(extractRichContent(document.querySelector('.font-cxsecret'))),
    STEM_SCRAMBLED,
  );
});

test('页面找不到 cxsecret 类名时退化为整页还原（对方改名也不至于输出乱码）', () => {
  // 该文档带反爬字体，但没有任何 cxsecret 类名标记
  const document = createDocument(`${FONT_STYLE}<div class="renamed">${STEM_SCRAMBLED}</div>`);
  const decoder = getCxSecretDecoder(document);

  assert.notEqual(decoder, null);
  assert.equal(decoder.scoped, false);
  assert.equal(decodeCxSecretText(DECOY, document.querySelector('.renamed')), DECOY_REAL);
});

test('extractRichContent 还原题干与选项正文', () => {
  const document = createDocument(`${FONT_STYLE}${renderScrambledQuestion()}`);

  assert.equal(
    richContentToText(extractRichContent(document.querySelector('.fontLabel'))),
    '【单选题】' + EXPECTED_STEM,
  );
  const optionTexts = Array.from(document.querySelectorAll('.Zy_ulTop li .after')).map((element) =>
    richContentToText(extractRichContent(element)),
  );
  assert.deepEqual(optionTexts, OPTIONS_DECODED);
});

test('端到端：真实混淆页面经 TiMuExtractor 提取后题干与选项可读', () => {
  const document = createDocument(`${FONT_STYLE}${renderScrambledQuestion()}`);
  const questions = new TiMuExtractor().extract(context(document));

  assert.equal(questions.length, 1);
  const question = questions[0];
  assert.equal(question.type, 'single-choice');
  // 题型标记本身未被混淆，且与外层同处作用域，应原样保留
  assert.equal(question.typeMeta, '【单选题】');
  assert.equal(richContentToText(question.stem), EXPECTED_STEM);
  assert.deepEqual(
    question.options.map((option) => option.key),
    ['A', 'B', 'C', 'D'],
  );
  assert.deepEqual(
    question.options.map((option) => richContentToText(option.content)),
    OPTIONS_DECODED,
  );
});

test('replaceByCodePointMap 未命中时原样返回同一字符串引用', () => {
  const map = new Map([[0x53c3, '中']]);
  const untouched = '毫无混淆的文本';

  assert.equal(replaceByCodePointMap(untouched, map), untouched);
  assert.equal(replaceByCodePointMap('參', map), '中');
  assert.equal(replaceByCodePointMap('', map), '');
});
