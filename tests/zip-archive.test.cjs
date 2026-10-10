// 内置 ZIP 打包器的字节级用例。
//
// 打包器是手写的（不引入 jszip），所以结构必须逐字段验证：
// 本地文件头 / 中央目录项 / 中央目录结束记录三段的签名、偏移、CRC-32 与 UTF-8 文件名标志。
// 真实解压校验（Python zipfile 读回内容并验 CRC）在 .workbuddy/verify-zip.cjs 里。
const test = require('node:test');
const assert = require('node:assert/strict');
const { createZipBlob } = require('../.tmp/test/src/exporters/zip-archive.js');

const encoder = new TextEncoder();
const decoder = new TextDecoder();

async function bytesOf(blob) {
  return new Uint8Array(await blob.arrayBuffer());
}

function viewOf(bytes) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

const u16 = (view, offset) => view.getUint16(offset, true);
const u32 = (view, offset) => view.getUint32(offset, true);

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
/** 通用标志位 11：文件名按 UTF-8 编码 */
const UTF8_FLAG = 0x0800;
/** 本地文件头固定部分 30 字节，中央目录项固定部分 46 字节，结束记录 22 字节 */
const LOCAL_SIZE = 30;
const CENTRAL_SIZE = 46;
const END_SIZE = 22;

/** 中央目录项里各字段的相对偏移 */
const CENTRAL_EXTERNAL_ATTR = 38;
const CENTRAL_LOCAL_OFFSET = 42;

test('ZIP：本地文件头字段与 CRC-32 正确（已知值 abc → 0x352441C2）', async () => {
  const bytes = await bytesOf(createZipBlob([{ name: 'a.txt', data: encoder.encode('abc') }]));
  const view = viewOf(bytes);

  assert.equal(u32(view, 0), LOCAL_SIGNATURE);
  assert.equal(u16(view, 4), 20, '解压所需版本');
  assert.equal(u16(view, 6), UTF8_FLAG, '文件名 UTF-8 标志');
  assert.equal(u16(view, 8), 0, '存储方式：不压缩');
  assert.equal(u32(view, 14), 0x352441c2, 'CRC-32');
  assert.equal(u32(view, 18), 3, '压缩后大小');
  assert.equal(u32(view, 22), 3, '原始大小');
  assert.equal(u16(view, 26), 5, '文件名长度 a.txt');
  assert.equal(u16(view, 28), 0, '无扩展字段');
  // 文件名与数据紧跟文件头
  assert.equal(decoder.decode(bytes.slice(30, 35)), 'a.txt');
  assert.equal(decoder.decode(bytes.slice(35, 38)), 'abc');

  // 中央目录项与结束记录
  const centralAt = LOCAL_SIZE + 5 + 3;
  assert.equal(u32(view, centralAt), CENTRAL_SIGNATURE);
  assert.equal(u32(view, centralAt + CENTRAL_LOCAL_OFFSET), 0, '条目指向第一个本地头');
  const endAt = centralAt + CENTRAL_SIZE + 5;
  assert.equal(endAt, bytes.length - END_SIZE);
  assert.equal(u32(view, endAt), END_SIGNATURE);
  assert.equal(u16(view, endAt + 8), 1, '本磁盘条目数');
  assert.equal(u16(view, endAt + 10), 1, '总条目数');
  assert.equal(u32(view, endAt + 12), CENTRAL_SIZE + 5, '中央目录大小');
  assert.equal(u32(view, endAt + 16), centralAt, '中央目录起始偏移');
});

test('ZIP：多个条目时中央目录记录的本地头偏移各自正确', async () => {
  const entries = [
    { name: '卷子/卷子_1_导论1.txt', data: encoder.encode('第一章内容') },
    { name: '卷子/卷子_2_导论2.txt', data: encoder.encode('第二章内容') },
  ];
  const names = entries.map((entry) => encoder.encode(entry.name));
  const bytes = await bytesOf(createZipBlob(entries));
  const view = viewOf(bytes);

  const secondLocalAt = LOCAL_SIZE + names[0].length + entries[0].data.length;
  assert.equal(u32(view, secondLocalAt), LOCAL_SIGNATURE, '第二个条目的本地头紧跟第一个');

  const centralAt = secondLocalAt + LOCAL_SIZE + names[1].length + entries[1].data.length;
  assert.equal(u32(view, centralAt), CENTRAL_SIGNATURE);
  assert.equal(u32(view, centralAt + CENTRAL_LOCAL_OFFSET), 0);
  const secondCentralAt = centralAt + CENTRAL_SIZE + names[0].length;
  assert.equal(u32(view, secondCentralAt), CENTRAL_SIGNATURE);
  assert.equal(u32(view, secondCentralAt + CENTRAL_LOCAL_OFFSET), secondLocalAt);

  const endAt = secondCentralAt + CENTRAL_SIZE + names[1].length;
  assert.equal(u32(view, endAt), END_SIGNATURE);
  assert.equal(u16(view, endAt + 10), 2);
  assert.equal(u32(view, endAt + 16), centralAt, '中央目录起始偏移');
});

test('ZIP：中文章节名按 UTF-8 写入，目录条目标记为目录', async () => {
  const folder = '如何认识新方位';
  const folderName = `${folder}/`;
  const fileName = `${folder}/卷子_1_如何认识新方位.txt`;
  const bytes = await bytesOf(
    createZipBlob([
      { name: folderName, data: new Uint8Array(0) },
      { name: fileName, data: encoder.encode('题干') },
    ]),
  );
  const view = viewOf(bytes);
  const folderBytes = encoder.encode(folderName);

  assert.equal(u16(view, 6), UTF8_FLAG);
  assert.equal(u16(view, 26), folderBytes.length);
  assert.equal(decoder.decode(bytes.slice(30, 30 + folderBytes.length)), folderName);
  // 目录条目：内容为空
  assert.equal(u32(view, 18), 0, '压缩后大小');
  assert.equal(u32(view, 22), 0, '原始大小');

  // 外部属性只在中央目录项里：第一条目的目录位应为 1
  const centralAt =
    LOCAL_SIZE +
    folderBytes.length +
    LOCAL_SIZE +
    encoder.encode(fileName).length +
    encoder.encode('题干').length;
  assert.equal(u32(view, centralAt + CENTRAL_EXTERNAL_ATTR), 0x10);

  // 中文字节数 = 字符数 × 3，可被独立 UTF-8 解码还原
  assert.equal(folderBytes.length, folder.length * 3 + 1);
});

test('ZIP：没有条目时只产出 22 字节的结束记录', async () => {
  const bytes = await bytesOf(createZipBlob([]));
  const view = viewOf(bytes);
  assert.equal(bytes.length, END_SIZE);
  assert.equal(u32(view, 0), END_SIGNATURE);
  assert.equal(u16(view, 10), 0);
});

test('ZIP：时间戳用 MS-DOS 格式写入', async () => {
  const stamp = new Date(2026, 9, 10, 21, 32, 14);
  const bytes = await bytesOf(createZipBlob([{ name: 'a', data: new Uint8Array([1]) }], stamp));
  const view = viewOf(bytes);

  const expectedTime = (21 << 11) | (32 << 5) | (14 >> 1);
  const expectedDate = ((2026 - 1980) << 9) | (10 << 5) | 10;
  assert.equal(u16(view, 10), expectedTime);
  assert.equal(u16(view, 12), expectedDate);
});
