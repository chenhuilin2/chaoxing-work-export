/**
 * font-cxsecret 字体反爬解码 —— 纯逻辑层（不依赖 DOM）
 *
 * 学习通把题干/选项里的部分汉字换成「备用码位」，再用内联字体把该码位画成真字的轮廓，
 * 于是人眼看正常、读 textContent 得到乱码（如「中悢共产党悤史上」）。
 *
 * 破解思路：子集字体的 glyf 表逐字节保留了源字体的字形数据，因此
 *
 *     sha1(glyf[loca[gid] : loca[gid + 1]])[:4]  →  真实汉字码位
 *
 * 只要在运行时对页面字体做同样的计算并查表，即可逐字还原。
 * 映射每个页面随机打乱，所以必须按页面实时解码，静态「码位→汉字」表无效。
 * 4 字节键在 20902 个汉字上实测零冲突（见 tools/gen-cxsecret-table.py）。
 */

import { CXSECRET_GLYPH_TABLE_B64 } from './cxsecret-table';

/** 表内每条记录：4 字节哈希键 + 2 字节 BMP 码位 */
const ENTRY_SIZE = 6;
/** 认定一个内嵌字体是反爬字体所需的最少命中数，避免误伤普通内嵌字体 */
const MIN_DECODE_HITS = 4;

/* ------------------------------------------------------------------ base64 */

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const BASE64_INDEX = (() => {
  const index = new Int16Array(128).fill(-1);
  for (let i = 0; i < BASE64_ALPHABET.length; i += 1) index[BASE64_ALPHABET.charCodeAt(i)] = i;
  return index;
})();

/** 自带 base64 解码，避免依赖宿主环境的 atob（油猴沙箱/Node 表现不一致） */
export function decodeBase64(input: string): Uint8Array {
  const clean = input.replace(/[^A-Za-z0-9+/]/g, '');
  const length = clean.length;
  const output = new Uint8Array(((length * 3) >> 2) + 3);
  let buffer = 0;
  let bits = 0;
  let offset = 0;
  for (let i = 0; i < length; i += 1) {
    const code = clean.charCodeAt(i);
    const value = code < 128 ? (BASE64_INDEX[code] ?? -1) : -1;
    if (value < 0) continue;
    buffer = (buffer << 6) | value;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      output[offset] = (buffer >> bits) & 0xff;
      offset += 1;
    }
  }
  return output.subarray(0, offset);
}

/* -------------------------------------------------------------------- SHA-1 */

/**
 * 计算 SHA-1 并返回摘要首 4 字节（大端无符号整数）。
 * 只取首 4 字节即可：既省一次输出序列化，也把键压到 4 字节。
 */
export function sha1Prefix32(bytes: Uint8Array): number {
  const length = bytes.length;
  // 补齐：数据 + 1 字节 0x80 + 8 字节长度，整体对齐到 64 字节
  const padded = new Uint8Array((((length + 8) >> 6) + 1) << 6);
  padded.set(bytes);
  padded[length] = 0x80;
  const view = new DataView(padded.buffer);
  const bitLength = length * 8;
  view.setUint32(padded.length - 8, Math.floor(bitLength / 0x100000000));
  view.setUint32(padded.length - 4, bitLength >>> 0);

  let h0 = 0x67452301;
  let h1 = 0xefcdab89;
  let h2 = 0x98badcfe;
  let h3 = 0x10325476;
  let h4 = 0xc3d2e1f0;
  const words = new Uint32Array(80);

  for (let block = 0; block < padded.length; block += 64) {
    for (let i = 0; i < 16; i += 1) words[i] = view.getUint32(block + i * 4);
    for (let i = 16; i < 80; i += 1) {
      const mixed =
        (words[i - 3] ?? 0) ^ (words[i - 8] ?? 0) ^ (words[i - 14] ?? 0) ^ (words[i - 16] ?? 0);
      words[i] = (mixed << 1) | (mixed >>> 31);
    }
    let a = h0;
    let b = h1;
    let c = h2;
    let d = h3;
    let e = h4;
    for (let i = 0; i < 80; i += 1) {
      let f: number;
      let k: number;
      if (i < 20) {
        f = (b & c) | (~b & d);
        k = 0x5a827999;
      } else if (i < 40) {
        f = b ^ c ^ d;
        k = 0x6ed9eba1;
      } else if (i < 60) {
        f = (b & c) | (b & d) | (c & d);
        k = 0x8f1bbcdc;
      } else {
        f = b ^ c ^ d;
        k = 0xca62c1d6;
      }
      const temp = ((((a << 5) | (a >>> 27)) + f + e + k + (words[i] ?? 0)) | 0) >>> 0;
      e = d;
      d = c;
      c = ((b << 30) | (b >>> 2)) >>> 0;
      b = a;
      a = temp;
    }
    h0 = (h0 + a) >>> 0;
    h1 = (h1 + b) >>> 0;
    h2 = (h2 + c) >>> 0;
    h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0;
  }
  return h0;
}

/* --------------------------------------------------------------- 解码表查询 */

let cachedTable: Uint8Array | null | undefined;

/** 懒解析内置解码表（163 KB base64 → 122 KB 二进制），只做一次 */
export function getGlyphTable(): Uint8Array | null {
  if (cachedTable !== undefined) return cachedTable;
  try {
    const compact = CXSECRET_GLYPH_TABLE_B64.replace(/\s+/g, '');
    cachedTable = compact ? decodeBase64(compact) : null;
  } catch {
    cachedTable = null;
  }
  return cachedTable;
}

function readKey(table: Uint8Array, offset: number): number {
  const a = table[offset] ?? 0;
  const b = table[offset + 1] ?? 0;
  const c = table[offset + 2] ?? 0;
  const d = table[offset + 3] ?? 0;
  return ((a << 24) | (b << 16) | (c << 8) | d) >>> 0;
}

/** 表按哈希键升序排列，二分查找；返回真实汉字码位，未命中返回 0 */
export function lookupGlyphCodePoint(table: Uint8Array, key: number): number {
  let low = 0;
  let high = Math.floor(table.length / ENTRY_SIZE) - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    const base = mid * ENTRY_SIZE;
    const current = readKey(table, base);
    if (current === key) return ((table[base + 4] ?? 0) << 8) | (table[base + 5] ?? 0);
    if (current < key) low = mid + 1;
    else high = mid - 1;
  }
  return 0;
}

/* ------------------------------------------------------------------ sfnt 解析 */

function readTag(view: DataView, offset: number): string {
  return String.fromCharCode(
    view.getUint8(offset),
    view.getUint8(offset + 1),
    view.getUint8(offset + 2),
    view.getUint8(offset + 3),
  );
}

interface SfntDirectory {
  readonly view: DataView;
  readonly bytes: Uint8Array;
  readonly tables: ReadonlyMap<string, readonly [number, number]>;
}

function readSfntDirectory(font: ArrayBuffer | Uint8Array): SfntDirectory | null {
  const bytes = font instanceof Uint8Array ? font : new Uint8Array(font);
  if (bytes.byteLength < 12) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // TTC（字体集合）取其第一个字体；页面内嵌字体一般是单个 TTF
  let base = 0;
  if (readTag(view, 0) === 'ttcf') {
    if (bytes.byteLength < 16) return null;
    base = view.getUint32(12);
  }
  if (base + 12 > bytes.byteLength) return null;
  const count = view.getUint16(base + 4);
  const tables = new Map<string, readonly [number, number]>();
  for (let i = 0; i < count; i += 1) {
    const entry = base + 12 + i * 16;
    if (entry + 16 > bytes.byteLength) break;
    tables.set(readTag(view, entry), [view.getUint32(entry + 8), view.getUint32(entry + 12)]);
  }
  return { view, bytes, tables };
}

function readLoca(view: DataView, offset: number, numGlyphs: number, format: number): number[] | null {
  const count = numGlyphs + 1;
  const size = format === 0 ? 2 : 4;
  if (offset + count * size > view.byteLength) return null;
  const offsets: number[] = new Array(count);
  for (let i = 0; i < count; i += 1) {
    offsets[i] = format === 0 ? view.getUint16(offset + i * 2) * 2 : view.getUint32(offset + i * 4);
  }
  return offsets;
}

/** 读取 cmap → { 码位: 字形索引 }；支持 format 4（BMP）与 format 12 */
function readCmap(view: DataView, offset: number): Map<number, number> {
  const result = new Map<number, number>();
  if (offset + 4 > view.byteLength) return result;
  const count = view.getUint16(offset + 2);
  const subtables: Array<{ rank: number; offset: number }> = [];
  for (let i = 0; i < count; i += 1) {
    const entry = offset + 4 + i * 8;
    if (entry + 8 > view.byteLength) break;
    const platform = view.getUint16(entry);
    const encoding = view.getUint16(entry + 2);
    const rank = platform === 3 && encoding === 1 ? 0 : platform === 3 && encoding === 10 ? 1 : 2;
    subtables.push({ rank, offset: offset + view.getUint32(entry + 4) });
  }
  subtables.sort((left, right) => left.rank - right.rank);

  for (const subtable of subtables) {
    const start = subtable.offset;
    if (start + 2 > view.byteLength) continue;
    const format = view.getUint16(start);
    if (format === 4) {
      const segmentCount = view.getUint16(start + 6) >> 1;
      const endBase = start + 14;
      const startBase = endBase + segmentCount * 2 + 2;
      const deltaBase = startBase + segmentCount * 2;
      const rangeBase = deltaBase + segmentCount * 2;
      if (rangeBase + segmentCount * 2 > view.byteLength) continue;
      for (let i = 0; i < segmentCount; i += 1) {
        const end = view.getUint16(endBase + i * 2);
        const first = view.getUint16(startBase + i * 2);
        if (first === 0xffff) continue;
        const delta = view.getInt16(deltaBase + i * 2);
        const rangeOffset = view.getUint16(rangeBase + i * 2);
        for (let code = first; code <= end; code += 1) {
          if (code === 0xffff) break;
          if (result.has(code)) continue;
          let glyph: number;
          if (rangeOffset === 0) {
            glyph = (code + delta) & 0xffff;
          } else {
            const pointer = rangeBase + i * 2 + rangeOffset + (code - first) * 2;
            if (pointer + 2 > view.byteLength) continue;
            glyph = view.getUint16(pointer);
            if (glyph !== 0) glyph = (glyph + delta) & 0xffff;
          }
          if (glyph !== 0) result.set(code, glyph);
        }
      }
      continue;
    }
    if (format === 12) {
      const groupCount = view.getUint32(start + 12);
      for (let i = 0; i < groupCount; i += 1) {
        const group = start + 16 + i * 12;
        if (group + 12 > view.byteLength) break;
        const first = view.getUint32(group);
        const last = view.getUint32(group + 4);
        const glyph = view.getUint32(group + 8);
        for (let code = first; code <= last; code += 1) {
          if (result.has(code)) continue;
          result.set(code, glyph + (code - first));
        }
      }
    }
  }
  return result;
}

/**
 * 解析内嵌字体，得到「混淆码位 → 真实汉字」映射。
 * 映射为空表示这不是反爬字体（例如普通内嵌字体，其字形还原后码位不变）。
 */
export function buildGlyphDecodeMap(font: ArrayBuffer | Uint8Array): Map<number, string> {
  const result = new Map<number, string>();
  const table = getGlyphTable();
  if (!table) return result;
  const sfnt = readSfntDirectory(font);
  if (!sfnt) return result;
  const { view, bytes, tables } = sfnt;
  const glyf = tables.get('glyf');
  const loca = tables.get('loca');
  const head = tables.get('head');
  const maxp = tables.get('maxp');
  const cmapTable = tables.get('cmap');
  // 无 glyf 说明是 CFF 轮廓，本机制不适用
  if (!glyf || !loca || !head || !maxp || !cmapTable) return result;
  if (head[0] + 52 > bytes.byteLength) return result;

  const numGlyphs = view.getUint16(maxp[0] + 4);
  const offsets = readLoca(view, loca[0], numGlyphs, view.getInt16(head[0] + 50));
  if (!offsets) return result;
  const cmap = readCmap(view, cmapTable[0]);
  if (cmap.size === 0) return result;

  const glyfOffset = glyf[0];
  let hits = 0;
  for (const [codePoint, glyph] of cmap) {
    if (glyph + 1 >= offsets.length) continue;
    const start = offsets[glyph] ?? 0;
    const end = offsets[glyph + 1] ?? 0;
    if (end <= start) continue;
    const decoded = lookupGlyphCodePoint(table, sha1Prefix32(bytes.subarray(glyfOffset + start, glyfOffset + end)));
    // 解码结果与码位相同说明该字形本就是它自己，不属于混淆，跳过
    if (decoded === 0 || decoded === codePoint) continue;
    result.set(codePoint, String.fromCharCode(decoded));
    hits += 1;
  }
  // 命中过少则视为误判，避免把普通内嵌字体当反爬字体处理
  return hits >= MIN_DECODE_HITS ? result : new Map<number, string>();
}

/** 按映射逐字替换；无任何替换时原样返回，避免多余字符串拼接 */
export function replaceByCodePointMap(text: string, map: ReadonlyMap<number, string>): string {
  if (!text || map.size === 0) return text;
  let output = '';
  let changed = false;
  for (const char of text) {
    const decoded = map.get(char.codePointAt(0) ?? 0);
    if (decoded === undefined) {
      output += char;
    } else {
      output += decoded;
      changed = true;
    }
  }
  return changed ? output : text;
}
