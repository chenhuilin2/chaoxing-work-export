// 内置 ZIP 打包器：只做「存储」（不压缩），不引入任何外部依赖。
//
// 用途：按章节拆分的文件超过阈值时，把所有章节文件装进一个压缩包 —— 浏览器没法在一次下载里
// 直接建目录，压缩包解压后就是「一个装着所有章节文件的文件夹」。
//
// 两个必须小心的细节：
// 1. 章节名是中文，条目名要按 UTF-8 写入并置位通用标志位 11（语言编码标志），
//    否则 Windows 自带解压会按本地代码页解码成乱码；
// 2. ZIP 的所有多字节字段都是小端，CRC-32 用的是反射多项式 0xEDB88320。

/** 压缩包里的一个条目 */
export interface ZipEntry {
  /** 包内路径，用 `/` 分隔；以 `/` 结尾表示目录条目 */
  readonly name: string;
  /** 文件内容（目录条目传空数组） */
  readonly data: Uint8Array;
}

/** CRC-32 查表（构造时生成一次） */
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let index = 0; index < data.length; index += 1) {
    crc = (CRC_TABLE[(crc ^ (data[index] ?? 0)) & 0xff] ?? 0) ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** 小端字节写入器 */
class ByteWriter {
  private readonly chunks: Uint8Array[] = [];
  private size = 0;

  u16(value: number): void {
    const bytes = new Uint8Array(2);
    new DataView(bytes.buffer).setUint16(0, value & 0xffff, true);
    this.chunks.push(bytes);
    this.size += 2;
  }

  u32(value: number): void {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value >>> 0, true);
    this.chunks.push(bytes);
    this.size += 4;
  }

  raw(bytes: Uint8Array): void {
    this.chunks.push(bytes);
    this.size += bytes.length;
  }

  /** 已写入的字节数（中央目录里记录各条目的偏移要用） */
  get length(): number {
    return this.size;
  }

  get parts(): Uint8Array[] {
    return this.chunks;
  }
}

/** ZIP 头里的时间戳是 MS-DOS 格式（1980 年起算，秒只有 2 秒精度） */
function dosTimestamp(date: Date): { time: number; date: number } {
  const time =
    ((date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1)) & 0xffff;
  const day =
    (((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()) & 0xffff;
  return { time, date: day };
}

/** 通用标志位 11：条目名按 UTF-8 编码 */
const UTF8_FLAG = 0x0800;
/** 单个条目 4GB 上限（不做 ZIP64，章节文件不可能到这个量级） */
const ZIP64_LIMIT = 0xffffffff;
/** 外部属性里的「目录」位 */
const DIRECTORY_ATTRIBUTE = 0x10;

/**
 * 打成 ZIP。
 *
 * 条目名原样写入，所以调用方可以传「文件夹/章节文件.txt」这样的相对路径，
 * 解压后自然落进一个同名文件夹。
 */
export function createZipBlob(entries: readonly ZipEntry[], now: Date = new Date()): Blob {
  const encoder = new TextEncoder();
  const stamp = dosTimestamp(now);
  const local = new ByteWriter();
  const central = new ByteWriter();
  let offset = 0;

  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const size = entry.data.length;
    if (size > ZIP64_LIMIT) {
      throw new Error(`ZIP 打包失败：条目「${entry.name}」超过 4GB`);
    }
    const crc = crc32(entry.data);

    // 本地文件头
    local.u32(0x04034b50);
    local.u16(20);
    local.u16(UTF8_FLAG);
    local.u16(0); // 存储方式：0 = 不压缩
    local.u16(stamp.time);
    local.u16(stamp.date);
    local.u32(crc);
    local.u32(size);
    local.u32(size);
    local.u16(name.length);
    local.u16(0); // 无扩展字段
    local.raw(name);
    local.raw(entry.data);

    // 中央目录项
    central.u32(0x02014b50);
    central.u16(20); // 创建版本
    central.u16(20); // 解压所需版本
    central.u16(UTF8_FLAG);
    central.u16(0);
    central.u16(stamp.time);
    central.u16(stamp.date);
    central.u32(crc);
    central.u32(size);
    central.u32(size);
    central.u16(name.length);
    central.u16(0); // 无扩展字段
    central.u16(0); // 无注释
    central.u16(0); // 起始磁盘号
    central.u16(0); // 内部属性
    central.u32(entry.name.endsWith('/') ? DIRECTORY_ATTRIBUTE : 0);
    central.u32(offset);
    central.raw(name);

    offset += 30 + name.length + size;
  }

  // 中央目录结束记录
  const end = new ByteWriter();
  end.u32(0x06054b50);
  end.u16(0); // 本磁盘号
  end.u16(0); // 中央目录所在磁盘号
  end.u16(entries.length);
  end.u16(entries.length);
  end.u32(central.length);
  end.u32(local.length);
  end.u16(0); // 无注释

  return new Blob([...local.parts, ...central.parts, ...end.parts], { type: 'application/zip' });
}
