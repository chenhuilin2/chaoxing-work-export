// 导出服务用例：文件划分（含「章节多了打成一个 zip」）、压缩包内容，以及 TXT/MD 的章节化文本。
//
// 只覆盖 TXT / Markdown 分支：Node 原生有 Blob，不需要浏览器环境；export-service 只在 Word
// 分支里碰 window 与 docx（那条链路由 tests/word-export.test.cjs 打桩覆盖）。
const test = require('node:test');
const assert = require('node:assert/strict');
const { ExportService } = require('../.tmp/test/src/exporters/export-service.js');
const { chapter, exportOptions, question, resultOf } = require('./export-fixtures.cjs');

const service = new ExportService();

/** 造 n 个各带一道单选题的章节 */
function manyChapters(count) {
  return Array.from({ length: count }, (_, index) =>
    chapter(`c${index + 1}`, `导论${index + 1}`, [question(`q${index + 1}`, 'single-choice')]),
  );
}

const textOf = (artifact) => artifact.blob.text();

/** 从 zip 字节里按条目名取出内容：本地文件头的「压缩后大小」固定在第 18 字节 */
function readEntry(bytes, name) {
  const needle = Buffer.from(name, 'utf8');
  const at = Buffer.from(bytes).indexOf(needle);
  assert.notEqual(at, -1, `压缩包里应有条目 ${name}`);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const size = view.getUint32(at - 12, true);
  return Buffer.from(bytes.slice(at + needle.length, at + needle.length + size)).toString('utf8');
}

test('TXT 多章节不拆分：一个文件里按章节分段，题号与题型序号按章重排', async () => {
  const result = resultOf(manyChapters(3), '卷子');

  const artifacts = await service.createArtifacts(result, exportOptions({ format: 'txt' }));

  assert.equal(artifacts.length, 1);
  assert.equal(artifacts[0].filename, '卷子.txt');
  assert.equal(artifacts[0].mimeType, 'text/plain;charset=utf-8');
  const text = await textOf(artifacts[0]);
  // 换章时章标题上方空 3 行（比改动前多 2 行）
  assert.equal(
    text,
    [
      '第 1 章 导论1',
      '',
      '一. 单选题（共1题）',
      '1. 题干q1',
      'A. 选项q1',
      '',
      '',
      '',
      '第 2 章 导论2',
      '',
      '一. 单选题（共1题）',
      '1. 题干q2',
      'A. 选项q2',
      '',
      '',
      '',
      '第 3 章 导论3',
      '',
      '一. 单选题（共1题）',
      '1. 题干q3',
      'A. 选项q3',
    ].join('\n'),
  );
});

test('TXT 按章节拆分且不超过 5 章：逐个下载，一章一个文件', async () => {
  const result = resultOf(manyChapters(5), '卷子');

  const artifacts = await service.createArtifacts(
    result,
    exportOptions({ format: 'txt', splitByChapter: true }),
  );

  assert.equal(artifacts.length, 5);
  // 文件名只有章名，没有「卷子_1_」这类前缀
  assert.deepEqual(
    artifacts.map((artifact) => artifact.filename),
    ['导论1.txt', '导论2.txt', '导论3.txt', '导论4.txt', '导论5.txt'],
  );
  // 单章文件不再插章节标题
  assert.equal(await textOf(artifacts[0]), '一. 单选题（共1题）\n1. 题干q1\nA. 选项q1');
});

test('TXT 按章节拆分超过 5 章：打成一个 zip，包内是「一个装着所有章节文件的文件夹」', async () => {
  const result = resultOf(manyChapters(6), '卷子');

  const artifacts = await service.createArtifacts(
    result,
    exportOptions({ format: 'txt', splitByChapter: true }),
  );

  assert.equal(artifacts.length, 1);
  assert.equal(artifacts[0].filename, '卷子.zip');
  assert.equal(artifacts[0].mimeType, 'application/zip');
  assert.equal(artifacts[0].archiveEntries, 6);

  const bytes = new Uint8Array(await artifacts[0].blob.arrayBuffer());
  const view = new DataView(bytes.buffer);
  // 结束记录里的条目数 = 1 个目录条目 + 6 个章节文件
  const endAt = bytes.length - 22;
  assert.equal(view.getUint32(endAt, true), 0x06054b50);
  assert.equal(view.getUint16(endAt + 10, true), 7);

  assert.equal(readEntry(bytes, '卷子/'), '');
  assert.equal(readEntry(bytes, '卷子/导论6.txt'), '一. 单选题（共1题）\n1. 题干q6\nA. 选项q6');
});

test('Markdown 拆分超过 5 章同样打包，且包里的文件是 .md', async () => {
  const result = resultOf(manyChapters(7), '卷子 - 7 个章节');

  const artifacts = await service.createArtifacts(
    result,
    exportOptions({ format: 'md', splitByChapter: true, filename: '卷子 - 7 个章节' }),
  );

  assert.equal(artifacts.length, 1);
  assert.equal(artifacts[0].filename, '卷子 - 7 个章节.zip');
  const bytes = new Uint8Array(await artifacts[0].blob.arrayBuffer());
  assert.equal(
    readEntry(bytes, '卷子 - 7 个章节/导论1.md'),
    '### 一、单选题（共1题）\n\n**1.** 题干q1\n\n- A. 选项q1',
  );
});

test('多章节带答案：答案紧跟在所属章之后，不再出现在文件末尾', async () => {
  const result = resultOf(manyChapters(2), '卷子');

  const artifacts = await service.createArtifacts(
    result,
    exportOptions({ format: 'txt', withAnswers: true }),
  );

  const text = await textOf(artifacts[0]);
  // 章 1 题目 → 章 1 答案 → 章 2 题目 → 章 2 答案
  const chapter1 = text.indexOf('第 1 章 导论1');
  const answer1 = text.indexOf('答案汇总');
  const chapter2 = text.indexOf('第 2 章 导论2');
  const answer2 = text.indexOf('答案汇总', answer1 + 1);
  assert.ok(chapter1 >= 0 && answer1 >= 0 && chapter2 >= 0 && answer2 >= 0, '四个锚点都应存在');
  assert.ok(
    chapter1 < answer1 && answer1 < chapter2 && chapter2 < answer2,
    `顺序应为 章1 < 答案1 < 章2 < 答案2，实际 ${[chapter1, answer1, chapter2, answer2].join(',')}`,
  );
  // 末尾是第 2 章的答案，而不是一份全文汇总
  assert.match(text, /1\. A$/);
  assert.equal(text.includes('错题汇总'), false);
});

test('复制到剪贴板的文本与下载的合并文件一致（多章节同样按章分段）', async () => {
  const result = resultOf(manyChapters(2), '卷子');
  const options = exportOptions({ format: 'txt', withAnswers: true });

  // 勾了「按章节拆分」时复制仍然给整份合并文本
  const preview = service.previewText(result, { ...options, splitByChapter: true });
  const [artifact] = await service.createArtifacts(result, options);

  assert.equal(preview, await textOf(artifact));
  assert.match(preview, /^第 1 章 导论1\n\n一\. 单选题（共1题）/);
  assert.match(preview, /\n\n\n\n第 2 章 导论2\n\n/);

  // Word 格式仍然不支持复制
  assert.equal(service.previewText(result, exportOptions({ format: 'word' })), '');
});

test('单章节结果：TXT 内容与文件名不受多章节改造影响', async () => {
  const result = resultOf([chapter('c1', '导论1', [question('q1', 'single-choice')])], '导论1');

  // 文件名留空时回退到结果标题（多章节分支才用得到章节序号与压缩包名）
  const artifacts = await service.createArtifacts(
    result,
    exportOptions({ format: 'txt', filename: '' }),
  );

  assert.equal(artifacts[0].filename, '导论1.txt');
  assert.equal(await textOf(artifacts[0]), '一. 单选题（共1题）\n1. 题干q1\nA. 选项q1');
});
