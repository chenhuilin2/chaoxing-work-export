#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 font-cxsecret 字形哈希解码表（src/extractors/cxsecret-table.ts）

原理
----
学习通的 font-cxsecret 反爬把真实汉字换成「备用码位」，再让内联字体把该码位的
字形画成真字的轮廓。人眼看着正常，读 textContent 得到乱码。
映射每个页面随机打乱，因此静态「码位→汉字」表无效。

但子集字体逐字节保留了源字体的 glyf 字形数据，于是可以：

    key = sha1( glyf[loca[gid] : loca[gid + 1]] )[:4]   # 4 字节
    value = 真实汉字的 BMP 码位

运行时（油猴脚本内）对页面字体做同样的计算并查表即可还原。
4 字节键在 20902 个汉字上实测 0 冲突。

输入（二选一）
--------------
--font PATH   字节兼容的源字体（TTF，必须有 glyf/loca/cmap）。推荐路径：
              只依赖 OFL 授权的字体文件，产出的表可完全自行复现。
              注意必须与服务端同批转换产物逐字节一致，否则哈希不匹配。
--pkl DIR     参考实现预生成的哈希表目录（HanSansCN_glyfHashedTables.pkl +
              HanSansCN_CmapTables.pkl）。便利路径，用于无源字体时重建。
              该目录需自备，本仓库不分发。

用法
----
    python tools/gen-cxsecret-table.py --pkl .workbuddy
    python tools/gen-cxsecret-table.py --font path/to/SourceHanSansCN-Normal.ttf

产出
----
    src/extractors/cxsecret-table.ts
"""

from __future__ import annotations

import argparse
import base64
import os
import pickle
import struct
import sys

from hashlib import sha1

# 覆盖范围：CJK 统一汉字基本区（与参考实现一致）
RANGE_LO, RANGE_HI = 0x4E00, 0x9FA5
LINE_WIDTH = 100


def tables(raw: bytes) -> dict:
    n = struct.unpack_from('>H', raw, 4)[0]
    out = {}
    for i in range(n):
        o = 12 + i * 16
        out[raw[o:o + 4].decode('latin1')] = struct.unpack_from('>II', raw, o + 8)
    return out


def cmap_mapping(raw: bytes) -> dict:
    """读取 cmap（format 4 / 12）→ {码位: 字形索引}"""
    t = tables(raw)
    co = t['cmap'][0]
    cnt = struct.unpack_from('>H', raw, co + 2)[0]
    subs = []
    for i in range(cnt):
        pid, eid, off = struct.unpack_from('>HHI', raw, co + 4 + i * 8)
        subs.append((pid, eid, co + off))
    # 优先 Windows BMP
    subs.sort(key=lambda x: (0 if (x[0], x[1]) == (3, 1) else 1))
    out = {}
    for _pid, _eid, sub in subs:
        fmt = struct.unpack_from('>H', raw, sub)[0]
        if fmt == 4:
            seg_x2 = struct.unpack_from('>H', raw, sub + 6)[0]
            seg = seg_x2 // 2
            endo, starto = sub + 14, sub + 16 + seg_x2
            deltao, rangeo = starto + seg_x2, starto + seg_x2 * 2
            ends = struct.unpack_from('>%dH' % seg, raw, endo)
            starts = struct.unpack_from('>%dH' % seg, raw, starto)
            deltas = struct.unpack_from('>%dh' % seg, raw, deltao)
            ranges = struct.unpack_from('>%dH' % seg, raw, rangeo)
            for i in range(seg):
                if starts[i] == 0xFFFF:
                    continue
                for c in range(starts[i], ends[i] + 1):
                    if ranges[i] == 0:
                        gid = (c + deltas[i]) & 0xFFFF
                    else:
                        p = rangeo + i * 2 + ranges[i] + (c - starts[i]) * 2
                        if p + 2 > len(raw):
                            continue
                        gid = struct.unpack_from('>H', raw, p)[0]
                        if gid:
                            gid = (gid + deltas[i]) & 0xFFFF
                    if gid:
                        out.setdefault(c, gid)
        elif fmt == 12:
            ngroups = struct.unpack_from('>I', raw, sub + 12)[0]
            for i in range(ngroups):
                s, e, g = struct.unpack_from('>III', raw, sub + 16 + i * 12)
                for c in range(s, e + 1):
                    out.setdefault(c, g + (c - s))
    return out


def from_font(path: str) -> list:
    """从字节兼容的源字体直接算哈希（推荐路径）"""
    raw = open(path, 'rb').read()
    t = tables(raw)
    for need in ('glyf', 'loca', 'head', 'maxp', 'cmap'):
        if need not in t:
            raise SystemExit('字体缺少 %s 表，无法使用：%s' % (need, path))
    loc_fmt = struct.unpack_from('>h', raw, t['head'][0] + 50)[0]
    nglyphs = struct.unpack_from('>H', raw, t['maxp'][0] + 4)[0]
    lo = t['loca'][0]
    if loc_fmt == 0:
        loca = [v * 2 for v in struct.unpack_from('>%dH' % (nglyphs + 1), raw, lo)]
    else:
        loca = list(struct.unpack_from('>%dI' % (nglyphs + 1), raw, lo))
    goff = t['glyf'][0]
    cm = cmap_mapping(raw)
    out = []
    for cp in range(RANGE_LO, RANGE_HI + 1):
        gid = cm.get(cp)
        if gid is None or gid + 1 >= len(loca):
            continue
        b = raw[goff + loca[gid]: goff + loca[gid + 1]]
        if not b:
            continue
        out.append((sha1(b).digest()[:4], cp))
    return out


def from_pkl(d: str) -> list:
    """用参考实现的预生成哈希表重建（便利路径）"""
    with open(os.path.join(d, 'HanSansCN_glyfHashedTables.pkl'), 'rb') as f:
        name2hash = {g: h for g, h in pickle.load(f)[0]}
    with open(os.path.join(d, 'HanSansCN_CmapTables.pkl'), 'rb') as f:
        raw_ct = pickle.load(f)
    out = []
    for cp, g in raw_ct:
        if cp is None or not (RANGE_LO <= cp <= RANGE_HI):
            continue
        h = name2hash.get(g)
        if h is None:
            continue
        out.append((h[0][:4], cp))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description='生成 font-cxsecret 解码表')
    ap.add_argument('--font', help='字节兼容的源字体 TTF（推荐）')
    ap.add_argument('--pkl', help='参考实现 pkl 所在目录')
    ap.add_argument('--out', default=None, help='输出 .ts 路径')
    args = ap.parse_args()

    if bool(args.font) == bool(args.pkl):
        ap.error('必须且只能指定 --font 或 --pkl 之一')

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_path = args.out or os.path.join(root, 'src', 'extractors', 'cxsecret-table.ts')

    entries = from_font(args.font) if args.font else from_pkl(args.pkl)
    entries.sort()

    # 冲突检测
    seen, coll = {}, []
    for k, cp in entries:
        if k in seen and seen[k] != cp:
            coll.append((k, seen[k], cp))
        seen[k] = cp
    if coll:
        # 冲突不可接受：运行时无法区分，会解出错误汉字
        for k, a, b in coll[:10]:
            print('  冲突 %s: U+%04X vs U+%04X' % (k.hex(), a, b), file=sys.stderr)
        raise SystemExit('4 字节键存在 %d 处冲突，需加长键长' % len(coll))

    blob = bytearray()
    for k, cp in entries:
        blob += k + struct.pack('>H', cp)
    b64 = base64.b64encode(bytes(blob)).decode('ascii')

    lines = [b64[i:i + LINE_WIDTH] for i in range(0, len(b64), LINE_WIDTH)]
    src = args.font or '(pkl 便利路径)'
    header = """/**
 * font-cxsecret 字形哈希解码表（自动生成，请勿手工编辑）
 *
 * 生成命令：python tools/gen-cxsecret-table.py --pkl <dir>
 *           python tools/gen-cxsecret-table.py --font <byte-compatible.ttf>
 * 数据源：%s
 *
 * 布局：每条 6 字节 = [sha1(glyph bytes)[:4]][真实汉字 BMP 码位 2 字节大端]，按哈希键升序。
 * 覆盖：U+4E00–U+9FA5 共 %d 字，4 字节键经生成器校验无冲突。
 *
 * 说明：本表是对字体字形数据计算得到的「哈希 → 码位」事实性索引，不含任何字体轮廓数据
 * 或第三方源代码。上游字形数据来自 Source Han Sans（SIL OFL 1.1，Adobe）。
 */
export const CXSECRET_GLYPH_TABLE_B64 = `
%s`;
""" % (src, len(entries), '\n'.join(lines))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(header)

    print('已写入 %s' % out_path)
    print('  条目 %d  裸数据 %.1f KB  base64 %.1f KB  键长 4 字节  冲突 0'
          % (len(entries), len(blob) / 1024, len(b64) / 1024))


if __name__ == '__main__':
    main()
