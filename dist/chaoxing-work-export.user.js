// ==UserScript==
// @name         超星学习通作业/考试一键提取导出 Word 文档
// @namespace    https://github.com/chenhuilin2/chaoxing-work-export
// @version      3.0.18
// @description  TypeScript 重构版：提取学习通作业与考试题目，支持富文本、答案解析、错题、章节批量提取及 Word/TXT/Markdown 导出
// @author       huilin and contributors
// @license      GPL-3.0-only
// @homepageURL  https://github.com/chenhuilin2/chaoxing-work-export
// @supportURL   https://github.com/chenhuilin2/chaoxing-work-export/issues
// @updateURL    https://raw.githubusercontent.com/chenhuilin2/chaoxing-work-export/main/dist/chaoxing-work-export.meta.js
// @downloadURL  https://raw.githubusercontent.com/chenhuilin2/chaoxing-work-export/main/dist/chaoxing-work-export.user.js
// @icon         https://pan-yz.chaoxing.com/favicon.ico
// @match        *://chaoxing.com/*
// @match        *://*.chaoxing.com/*
// @require      https://unpkg.com/docx@8.5.0/build/index.umd.js
// @run-at       document-end
// @grant        none
// ==/UserScript==

(function () {
  'use strict';
  const modules = new Map();

  function normalizePath(value) {
    const output = [];
    for (const segment of value.split('/')) {
      if (!segment || segment === '.') continue;
      if (segment === '..') output.pop();
      else output.push(segment);
    }
    return output.join('/');
  }

  function resolveRequest(parent, request) {
    if (!request.startsWith('.')) return request;
    const base = parent.includes('/') ? parent.slice(0, parent.lastIndexOf('/') + 1) : '';
    return normalizePath(base + request);
  }

  function define(name, dependencies, factory) {
    modules.set(name, {
      dependencies,
      factory,
      exports: {},
      initialized: false,
      initializing: false,
    });
  }

  function load(name) {
    const record = modules.get(name);
    if (!record) throw new Error('[Chaoxing Work Export] Missing module: ' + name);
    if (record.initialized) return record.exports;
    if (record.initializing) return record.exports;

    record.initializing = true;
    const module = { exports: record.exports };
    const localRequire = (request) => load(resolveRequest(name, request));
    const args = record.dependencies.map((dependency) => {
      if (dependency === 'require') return localRequire;
      if (dependency === 'exports') return record.exports;
      if (dependency === 'module') return module;
      return load(resolveRequest(name, dependency));
    });
    const returned = record.factory.apply(undefined, args);
    record.exports = returned !== undefined ? returned : module.exports;
    record.initialized = true;
    record.initializing = false;
    return record.exports;
  }

  define("app-config", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.UI_HOST_ID = exports.APP_VERSION = exports.APP_NAME = void 0;
    exports.APP_NAME = 'Chaoxing Work Export';
    exports.APP_VERSION = '3.0.18';
    exports.UI_HOST_ID = 'chaoxing-work-export-root';
});
define("domain/export-options", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
});
define("domain/question", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.QUESTION_TYPE_LONG_LABELS = exports.QUESTION_TYPE_LABELS = exports.QUESTION_TYPES = void 0;
    exports.emptyTypeCounts = emptyTypeCounts;
    exports.buildStatistics = buildStatistics;
    exports.deriveTypeOrder = deriveTypeOrder;
    exports.groupQuestions = groupQuestions;
    exports.QUESTION_TYPES = [
        'single-choice',
        'multiple-choice',
        'fill-blank',
        'true-false',
        'short-answer',
    ];
    exports.QUESTION_TYPE_LABELS = {
        'single-choice': '单选',
        'multiple-choice': '多选',
        'fill-blank': '填空',
        'true-false': '判断',
        'short-answer': '简答',
    };
    exports.QUESTION_TYPE_LONG_LABELS = {
        'single-choice': '单项选择题',
        'multiple-choice': '多项选择题',
        'fill-blank': '填空题',
        'true-false': '判断题',
        'short-answer': '简答题',
    };
    function emptyTypeCounts() {
        return {
            'single-choice': 0,
            'multiple-choice': 0,
            'fill-blank': 0,
            'true-false': 0,
            'short-answer': 0,
        };
    }
    function buildStatistics(questions) {
        const byType = emptyTypeCounts();
        let wrong = 0;
        let withCorrectAnswer = 0;
        let withUserAnswer = 0;
        let withAnalysis = 0;
        for (const question of questions) {
            byType[question.type] += 1;
            if (question.isWrong)
                wrong += 1;
            if (question.correctAnswer.length > 0)
                withCorrectAnswer += 1;
            if (question.userAnswer.length > 0)
                withUserAnswer += 1;
            if (question.analysis.length > 0)
                withAnalysis += 1;
        }
        return {
            total: questions.length,
            wrong,
            withCorrectAnswer,
            withUserAnswer,
            withAnalysis,
            byType,
        };
    }
    function deriveTypeOrder(questions) {
        const seen = new Set();
        const order = [];
        for (const question of questions) {
            if (!seen.has(question.type)) {
                seen.add(question.type);
                order.push(question.type);
            }
        }
        return order;
    }
    function groupQuestions(questions) {
        const groups = {
            'single-choice': [],
            'multiple-choice': [],
            'fill-blank': [],
            'true-false': [],
            'short-answer': [],
        };
        for (const question of questions)
            groups[question.type].push(question);
        return groups;
    }
});
define("domain/history", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
});
define("domain/settings", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.DEFAULT_SETTINGS = void 0;
    exports.DEFAULT_SETTINGS = {
        theme: 'auto',
        shortcut: { ctrl: true, shift: true, alt: false, key: 'e' },
        hideShortcut: { ctrl: true, shift: true, alt: false, key: 'h' },
        exportPreferences: {
            format: 'word',
            withAnswers: false,
            withWrong: false,
            shuffle: false,
            bankImport: false,
            splitByChapter: false,
        },
        enableDrag: false,
        rememberPanelPosition: true,
        panelPosition: null,
        // 默认打开：进入作业/考试/章节练习页即自动提取一次（用户可在设置里关掉）
        autoExtractOnLoad: true,
    };
});
define("utils/dom", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.textOf = textOf;
    exports.firstMatch = firstMatch;
    exports.allMatches = allMatches;
    exports.parseLeadingNumber = parseLeadingNumber;
    exports.collectAccessibleDocuments = collectAccessibleDocuments;
    exports.isEditableTarget = isEditableTarget;
    function textOf(element) {
        return (element?.textContent ?? '').replace(/\u00a0/g, ' ').trim();
    }
    function firstMatch(root, selectors) {
        for (const selector of selectors) {
            const element = root.querySelector(selector);
            if (element)
                return element;
        }
        return null;
    }
    function allMatches(root, selectors) {
        const output = [];
        const seen = new Set();
        for (const selector of selectors) {
            root.querySelectorAll(selector).forEach((element) => {
                if (!seen.has(element)) {
                    seen.add(element);
                    output.push(element);
                }
            });
        }
        return output;
    }
    function parseLeadingNumber(value) {
        // 两种写法都算题号：「1. 题干」这种带分隔符的前缀，以及 <i class="fl">1</i> 这种只含题号的节点。
        // 已批阅视图的题号节点是裸数字（无「.」「、」），只认前者会让该模板的题号整列为 undefined。
        // 不能放宽成「行首数字」：题干正文常以年份开头（「2024 年…」），那样会把年份当成题号。
        const match = value.match(/^\s*(\d+)\s*[.、．]/u) ?? value.match(/^\s*(\d{1,3})\s*$/u);
        if (!match?.[1])
            return undefined;
        const parsed = Number.parseInt(match[1], 10);
        return Number.isFinite(parsed) ? parsed : undefined;
    }
    /**
     * 深度优先收集可同步访问的文档，含同源 iframe。
     * 学习通把知识卡片、目录、答题页放在多层 iframe 里，定位 DOM 前必须先跨层收集。
     * 跨域 frame 会抛错并被跳过，交由 FrameBridge 通过 postMessage 处理。
     */
    function collectAccessibleDocuments(root = document, maxDepth = 5) {
        const output = [];
        const visited = new Set();
        const visit = (current, depth) => {
            if (visited.has(current) || depth > maxDepth)
                return;
            visited.add(current);
            output.push({ document: current, depth });
            current.querySelectorAll('iframe').forEach((frame) => {
                try {
                    if (frame.contentDocument)
                        visit(frame.contentDocument, depth + 1);
                }
                catch {
                    // 跨域 frame 无法同步访问，交由 FrameBridge 处理
                }
            });
        };
        visit(root, 0);
        return output;
    }
    function isEditableTarget(target) {
        if (!(target instanceof Element))
            return false;
        return Boolean(target.closest('input, textarea, select, [contenteditable="true"]'));
    }
});
define("utils/array", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.shuffleCopy = shuffleCopy;
    exports.uniqueBy = uniqueBy;
    function shuffleCopy(values, random = Math.random) {
        const output = [...values];
        for (let index = output.length - 1; index > 0; index -= 1) {
            const swapIndex = Math.floor(random() * (index + 1));
            const temporary = output[index];
            output[index] = output[swapIndex];
            output[swapIndex] = temporary;
        }
        return output;
    }
    function uniqueBy(values, keyOf) {
        const seen = new Set();
        const output = [];
        for (const value of values) {
            const key = keyOf(value);
            if (seen.has(key))
                continue;
            seen.add(key);
            output.push(value);
        }
        return output;
    }
});
define("extractors/contracts", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
});
define("utils/text", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.normalizeWhitespace = normalizeWhitespace;
    exports.normalizeInlineWhitespace = normalizeInlineWhitespace;
    exports.sanitizeFilename = sanitizeFilename;
    exports.normalizeAnswer = normalizeAnswer;
    exports.escapeMarkdownAlt = escapeMarkdownAlt;
    exports.escapeMarkdownUrl = escapeMarkdownUrl;
    exports.humanizeError = humanizeError;
    function normalizeWhitespace(value) {
        return value
            .replace(/\u00a0/g, ' ')
            .replace(/[\t\r\f ]+/g, ' ')
            .replace(/ *\n */g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    }
    function normalizeInlineWhitespace(value) {
        return value.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    }
    function sanitizeFilename(value, fallback = '学习通题目') {
        const sanitized = normalizeInlineWhitespace(value)
            .replace(/[\\/:*?"<>|]/g, '_')
            .replace(/[. ]+$/g, '')
            .slice(0, 120);
        return sanitized || fallback;
    }
    function normalizeAnswer(value) {
        const normalized = normalizeInlineWhitespace(value)
            .replace(/^(?:正确答案|参考答案|答案|我的答案|你的答案|学生答案)\s*[:：]?\s*/u, '')
            .replace(/[，、;；\s]+/g, '')
            .toUpperCase();
        if (/^(?:正确|对|TRUE|T|√|✓)$/u.test(normalized))
            return 'TRUE';
        if (/^(?:错误|错|FALSE|F|×|✕|X)$/u.test(normalized))
            return 'FALSE';
        return normalized;
    }
    function escapeMarkdownAlt(value) {
        return (value || '图片').replace(/[[\]\r\n]/g, ' ').trim() || '图片';
    }
    function escapeMarkdownUrl(value) {
        return value.replace(/[()\\]/g, (character) => {
            return `%${character.charCodeAt(0).toString(16).toUpperCase()}`;
        });
    }
    function humanizeError(error) {
        if (error instanceof Error && error.message)
            return error.message;
        return typeof error === 'string' ? error : '发生未知错误';
    }
});
define("extractors/cxsecret-table", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.CXSECRET_GLYPH_TABLE_B64 = void 0;
    /**
     * font-cxsecret 字形哈希解码表（自动生成，请勿手工编辑）
     *
     * 生成命令：python tools/gen-cxsecret-table.py --pkl <dir>
     *           python tools/gen-cxsecret-table.py --font <byte-compatible.ttf>
     * 数据源：(pkl 便利路径)
     *
     * 布局：每条 6 字节 = [sha1(glyph bytes)[:4]][真实汉字 BMP 码位 2 字节大端]，按哈希键升序。
     * 覆盖：U+4E00–U+9FA5 共 20902 字，4 字节键经生成器校验无冲突。
     *
     * 说明：本表是对字体字形数据计算得到的「哈希 → 码位」事实性索引，不含任何字体轮廓数据
     * 或第三方源代码。上游字形数据来自 Source Han Sans（SIL OFL 1.1，Adobe）。
     */
    exports.CXSECRET_GLYPH_TABLE_B64 = `
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
A+RNkIpoA+Rk7FhiA+c743GkA+emf3W1A+e55XXHA+4x+IGVA/HmYIIvA/J31GDpA/KD51hRA/RDyWiwA/TtyVTSA/bJolXzA/h0
fZACA/h5+p3kA/jR7pMzA/zxmofKBAJQa1p3BALGeIfOBARX35TKBAgxqJUABArDxHkXBAuu8WUxBA2Pb1ChBBDKtI0hBBG473a8
BBJKX2gRBBaEiFrkBBbfr4P4BBfIUp27BBhYF1jNBBzYoW6lBB6S9XWBBCDmY4WABCNxEG5vBCQzuoGqBC5673P0BDykFpUeBEfq
tI40BE9AOWC7BFWG0ZLfBFpr1lBLBF3KXptcBF97m4LJBGhQpnOWBG1WT22hBHSxc5FcBHacomtTBHdW3JkaBHoiV3pZBHtdapSJ
BHtfOGxbBHxD71X2BICOoV97BINbqHm+BIO57VOeBITtdIWxBIZV9GwlBIjByXADBJZv8pdBBJctBlwOBJuYAmklBJ4bJXngBKBd
SHFbBKCu9Gf1BKIG509fBKiJPE+7BKk/uYmuBKmnU5KdBKpRhlMLBLEzG5m0BLE3WoXCBLxQclaIBMNUM3mqBMdGYmkcBMd7T1Nz
BMfk3pbLBMf/RZS1BMkqP1WbBM1uPU6OBNLusmEkBNh/am+yBN6XYW8TBOSGGmDkBOjkPHEeBO+N9357BPLF75SSBPVVsIcvBPWl
T4/MBPirPV2wBP0KJpXoBQC2C3XrBQIJFH7tBQJjSH5wBQNb95hyBQPN/G84BQd9d52pBQf8V5vTBQq4tH9GBQuPR09iBQ2DzXoD
BQ69+neOBRCB+J0BBRIZOJSYBRIxSZshBRYv7pL1BRjN+YS8BRvKiXusBR9Na2uSBSCOI1bXBSCPpJx/BSMavJ49BSYsN5KwBSgM
r1dABSgd5pGoBSoP95vSBSvhWWjXBS1QM2uVBS4e+GZKBTefq5PgBThO92rqBThn9od1BTs6XIz3BT8uT1WvBUFGGIcHBUM7WWpf
BUdW5plFBUlPTVMfBUtF0ld7BU2PN2dwBU9S1VqABVAUKo/EBVEfFX6dBVKsn5M3BVk1j1A/BVlRt3K0BVnyopULBVvFp44XBV6y
vWmBBWO/1Ze5BWiXmZmHBWpkEl1SBW3jKnVOBW+M415/BXCsiZ6uBXb4m3rUBYGGt4F3BYGyhFW8BYQhiGawBYU6yFTeBYd6sJfN
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
K7ryYYIRK7x7G3wCK7/sxGZqK8SYjX1dK8a9IX5OK8pbpp5jK8tSIJaNK83yqVFdK872VZRdK9JDD5HZK9m0XYpEK9rnzF/XK+IP
tmq6K+Ktg4XOK+OmlYAmK+Y2tX0RK+wFzZ6YK/YRN5SxK/o9yFb0LABvcJ0uLALsPJy/LAn4uW7JLA5M5ohiLA851mI0LBMILnAm
LBfPv3NeLBl+mWP+LBsBimr1LBtAzWgHLB3ajGCqLCCc2YoxLCRnQ2J9LCuit1bRLDDOIlazLDDhUJafLDUImlqRLDd3LFdgLDq+
cF+fLDtryZNSLDxM8nWfLDzx1mTXLD3jLFwZLD5HA1L2LD9v/HE2LEYkbJoaLEauH3glLEcqF4ONLEeVVJWYLEl2Nlu4LE3SO59u
LFAec3NULFB2UJN7LFCgPJ1wLFDXHZ05LFNai4eYLFzVKXoLLF7fh10pLF+02HETLGKJA1KiLGUEJlkGLGZ61pScLGfK7k5JLGiw
Kn+tLGz9FE6bLHMTDXB2LHOYT3soLHc5C2zsLHfNVXEiLHmY/YXiLHmbIIjKLHpSK5edLH6EvJ+dLIc4/Jf5LIoepZbMLI9dIleq
LJC5nI9ZLJE7FVGFLJaybo3ZLJfDg5tWLJlVP4XoLJzGq4ZmLJ8c/31YLKNxC4vdLKRIdIg5LKYEB3vNLKfciHhWLKu6sn8eLLCR
DIEULLzu9GydLMRHAFBHLMYrxFYcLNEHEpruLNNWUJkxLNQYBXmILNaBDFSvLNbiU5JyLNcL0I6NLNx/BU/kLOOS5VmZLOTs+4oF
LOZmXW9nLOkjEmYbLOz/uY9mLO97uXxDLPAIG3L+LPPZi2nlLPTkBZOaLPXsn2eDLPa7r3SrLPdcO4hYLPhOSoKaLP+J1l1lLQbR
1GNRLQfSMJPbLQq9iHa/LQtlT58nLQzZOFZ8LRHNJVWtLRX1+12HLRaEV2LuLRaVin8vLRcykHQyLRkLCFBELRupXmB9LRveWFZ4
LR3G9WjjLR/IOoHyLSNLbZTHLSyk5JbYLTC5n2y5LTEObYP/LTK6iW3jLTRRK1H3LTewc3F0LTlUtIUPLTtFEpN0LUHkE5gLLUas
i5fJLUa+En12LUtW/lfYLU02Hm1aLU1kKoJuLU5zTnaqLVaOBY2+LVhel2tILVytFogHLV2UdZLsLV9kP5h1LWJF5VUVLWJvAnKz
LWQ2OJfmLWjpgIIVLYJA1FrPLYK9+1ISLYWLfpE4LYrcBmMsLZQkgJVFLZbGEnOoLZkcBJ2fLZmjY5oHLZniaF0NLZwL7FrsLZ2Q
9lkRLZ3R6mGmLaHKy5W8LaaT3pjYLah27lA2LalAdFCWLatxXGFgLazG1p9RLbEnA1BkLboxlm0QLcMcyY7DLcMyPZGXLcaIl2Gd
Lcr54mReLc3dZFGiLc36Dl2BLc7BFlWfLc9cy2/4Lc/PIV5iLdPy44OmLdRrAG2VLdTeoWuCLdWAYZIyLdXc4ZmhLdaKjpLMLdnL
TVhBLdux8WLDLd3a93kQLd/wXXdhLeFTuGysLeNltZipLeOCkWF1Leb5elFWLecZTZimLedCp3m4LeqO5WqCLex0fn8ULe3In5IW
LfWwSmO0Lfhq+JU5Lfz+zI8rLf4w+mQtLgHlUYcILgMOM4asLgj0p19iLg59PI7OLhPebFRoLhWhlpKjLhcbm16XLhjYdHrRLhsE
j4H0Lhw38pS0LiZs+omGLifOmpGMLiml+p93LiqvdnxsLituSILNLi4BQoDsLjB9CHOvLjJ4ilGoLjLqiYT4LjQSOpM2LjW52G8S
LjY223odLjnhaGLGLjq2856JLj/iqoC4Lkf0HoQ0Lk4JlJ9eLlH+/V9TLlXQjplwLlXV5lDkLllNxpdWLlmZJU6dLlySml5VLmV6
qnK+LmfRQIPnLmtelHDXLm6ekYInLm6mupbhLnVl8nEpLngfBI46LnyzN5yILn94tIjALoFYBJsSLoL6j5H0LoPJy56CLoQFyVTi
LoYjZVRGLogqPHYvLotn3FcaLpSm4WVULpYT9lmaLpzJoJP1Lpzzol8VLp113nuWLp2LYFzcLp/p5ICJLqGp2Y1lLqM3SE5RLqr+
uoFaLq/UQYHdLrP4NHIbLrlQv3SXLr+lvXgqLsQztI32Lsd7DJYXLsfFslocLsjeoowtLswtelYOLs2S/40QLtERM5oVLtLXh33i
LtRaxFJ9LtXC1ZHzLtmgK4jXLuAWO3H6LuGPDpjeLuHYxFgoLuY6e5+lLugTc40CLukxp5YjLukzcXQoLuttlX7ALu5TeofXLvA5
44hvLvMju1zyLvO3xU71LvRgYVNkLvhDdo1+Lvhs4IFGLvlcemsVLvqtdVqCLv69XYuELv/RzZQ5LwSNXY8gLwiKd55uLwi1EHxU
Lws0ZFQuLxZx+GIMLxbATHdZLxl3E1e1LxsDg1paLyQq/GVNLyeuElvrLyfYvoSbLyo71mjpLys3zpPKLzD4yVYkLzGiBFE7LzgU
7V0lLzly+lXcLzwJcoVQLz0JCk/KLz9IFFrbL0G39F5EL0HkuG3AL0MxCISwL0Q49IiPL0ToT4W1L0bLQ5mBL0xWO2YVL00LYpBy
L02Q3ZZ7L1COhWmWL1UZWGKNL1eO81ZRL1ejtp8HL1hECnstL18neJT4L2AJ825rL2HhD1nAL2hJnn1/L2rKNZC/L2zKGFE2L3BT
LG2aL3U0rXqzL39xoZ9rL4EPbn/xL4Fx640NL4Wr/4+lL4X1Clq2L4ZOQnq8L4awR3VIL5C+wIN0L5WvRU+ML5ZY0ngwL5auWpVZ
L5nTQYp4L5nwvGrTL5+ZpXeHL6PT/mjSL6b/hnchL63bdn7CL63slGXFL7BTJ2g3L7gKxpI3L79UFVhnL8RWLX+dL8WUOo5CL8r2
ho3dL80eDnKCL8/nFFbEL9NbMpLpL9hY54NpL9sQ84w1L913uWhmL99/QZhRL9+qsInBL+jxqXlDL+uMOFUlL/Jbr5GeL/WZv3sO
L/tBRVKCL//rs3PZMAOu7F4pMAidtVt/MA/sRJerMBIwqXjpMBRLxGW9MBYjcW3kMBuVmlurMCA391k5MCSKb2qZMCfjGGXPMCmO
K2zvMCnueofBMDZe947vMDfxcJZiMDjseJIEMD9E01x2MEAHNJ2aMEPsHk8qMEQwqHyOMESarHMXMEXvKF07MEhkFFyAMEkP+V85
MEuCUFBqME3Nm1LfMFARBnUdMFfIWpB2MFlVdJ50MGNQNn03MGi1dHhLMGqrypARMGsaYFu2MGvH+opdMG7ztWLJMHLEvFu+MHaw
umsGMHk15WaCMHoX7Gw/MHvDuo2wMIHDJIhIMINWypABMIwac2akMJMmJZzbMJU1bo0yMJVxDFl5MJYEipfkMJf+4XvOMJiJOJlN
MJlHWGdoMJtcPY/qMKJVmF5oMKRWLmSxMKR2SZURMKyY5GgsMK6OUnNoMK9Om13lMK/eG4jZMLAQxVq1MLIa0m8QMLIpuYqHMLKC
bXgfMLX7lXTsMLehWJoTMLoldFxyMLv+KlqvML25R1FQML6Aw3QaML7ssFz9ML7uYWfxML8TBW6HMMBpWmZ6MMLkDImDMMt7Vmd3
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
OAjiYoCWOAomCFLyOAp7ZX/ROAzCznUHOBPjN1DGOBSZA1PsOBblT5WROBwEpmhTOCMUgG9JOC6neVI5ODA45WATODF7RYFeODGe
kIqIODIWQ3P6ODdVlZS8ODjGQVjbODmxl1D4ODnbyn/XODuz/Z10OD6uBFQpOELJ0HeaOEQVdmyCOEYA33iPOEmmtlWBOEtEsoci
OE1kxIBdOE7knHZiOE/JKH8HOFOggY9jOFgr82ChOGSkP5ZoOGXKxJ9cOGhiNokrOGjCcIePOGnB1JKfOGolZmtMOGyEnGjJOIMN
2Y72OIMnenyyOINoBnu2OIqv7WPGOJRbdIXUOJwN6ZtkOKMQO1m+OKQU0F5GOKiYc3VDOKpRylugOLIYo3qwOLeNcozzOLh8V5Hj
OL9CPlVfOMSmb5UPOMbSCoJtOMrrf4V4OMsCaplSOMwj1mJBONiI12P3ONixnnAqONrqtFLZONysEY9xON8IC2BEOOMPT1uaOOcQ
p3zJOOdnG1JCOOefd1S0OOrcnFTGOPHjJYcKOPNHpWDZOPjA8WZROPkH+VQ/OP0aZJaxOP06H4xAOQBYkE+oOQOa+HlZOQTqY1qY
OQlHhJCkOQwRBXOSORBmFV04ORHgvmUqORRBMo2GORrUzHaMOR0GlHvYOSEvMGHuOSLJJ3ulOSN2fGxEOSZhzochOSzMf3VgOS0C
lJkVOTTYNmF3OTc2kmuYOTdRa5o/OTe8NH7gOTnRTW7mOTpyhmi2OTze+m+TOT4w2Gt0OT+aI57WOUZK+5QKOUsbLIbAOU48VW9s
OU6ulpGQOU7j/2IVOVfIV1VCOVgRxVBCOVliupFqOVmUs2N/OVrFBIf1OVujAXMLOVx9L1AKOV3INWw8OV4lJYaiOWXgd4pbOWet
MJiaOWhZDHx5OWtCm1YYOWtl15VQOWulyXGBOXKX4In9OXNY3IjcOXW9Vmo9OXgjb3lwOXiOAZrTOXj7o1qfOXn/B1ghOX279pKy
OX3BKJ0EOYOmJpkjOYYyr4juOYZmLICCOYmpZnFAOYsx3pewOYtkzm1MOYx8SWQdOZJOaIfFOZ+D8ZkiOaHu8VwVOaJm3FJaOapN
z2UDOarOZ3UhOa25X5xlOa3cYWbCObXxgk79ObqwNXyRObtQClExOcKxXk91OcV9iYFZOdM/mFwxOdaKgm8yOdo4S4LhOeBTZmnL
Oe/0EJANOfCEF3eqOfG2VZM0OfNiIlXvOfOTAWzMOfcnDGVwOfhUpXZMOfpcsnmzOfy0QomiOf5a15XJOf9ov5d5OgYdsGNKOga7
yZ8ZOgpmMJ6rOgyk8o31Og9DPF6tOhNpUZuHOhPRJHThOhVAbnv2OhjXzH+JOh2zaVQKOh9JqZd6OiB5/1TxOiXa2GUZOiYgM1sD
OismTlPGOjLytoTAOjONGV00OjecNn52OjxCwIxlOkLu/pgcOkXjRo20OkaMVHMdOknVOX8VOkp/pFx/OlT8WFmVOlV6ulffOlao
lmejOlbXuV3EOlqBt2VQOlzZEYH8Ol2Uj2aJOl66apOGOl7hQZf7Ol8MhlwFOmBot2vGOmU7KWDsOmZqUF4UOmtrZY7lOm2sAHO7
Om5DJYkLOm/vOltIOnRpclC9OnY+k5hIOngLI5ssOniBnYVwOn25FFU2OoEnbITwOoMlFIvtOoM23GN3OoaDAV72OoccklSGOo06
1YGGOo1GuYP9Oo7ftILWOpRVSmnxOpfJPWn1Opo/PneEOqe85GR4OqmXUZsJOquMeJyLOq0h3J2qOrV1ZnzoOrrEO3u+OrtkJJbQ
Or48wGDlOr5yaHOdOsF1h2whOsN2lFxuOsXylVO2OscTzWkBOssGkpOxOswWZm29Oszhh41VOtCTlZxTOtOBH5LROuRIrGPYOuUZ
cV5kOuXsE4STOufrM5mNOul2s3VZOuqm0lJfOvJG32GEOvtoQp1qOvuoaHo5OwC5j2/UOwPDNG2rOwgoEn6lOwqvG2Q2Ow3/4miW
OxTEnoBMOxToYIu2Oxeg/VZKOxw8DmeLOyEFSYYxOyMbWpzIOyMlz1grOyaZx18gOyzqYI3IOy5tHYm0OzG4bpPNOzWi0I9vOz4O
pWN0Oz8BJ3A7O0JLbk67O0L13oUtO0nKqWfCO0wJ8mI9O080W2ElO09IxnpYO1mupXsUO1yFr5TpO1+uNWhWO2CaWmb7O2SitZEi
O2UWAI66O2jOIF/RO2kE9VbVO22ipY12O29P6F6/O3XZnJMCO3fQaYO0O3ut0WgYO3+jBo3GO4Az432hO4A1amOhO4cf6ZnUO4eZ
7mPoO4jZklx9O4rrrIkYO4yJCmmqO5WKGHn0O5Z+X3k9O5dd2VOaO5rK2XDOO54d8pL2O55HFJdIO5+I9GUYO6KryH3gO6L3h353
O6+7uWOZO7E+1XbKO8AZkJT+O8FtWmVyO8P/W5NAO8QWNnH0O8alk3ATO8h0CXYtO8oXkGA0O8pqSWwMO9GmOHPGO9I2/nJMO9d/
53mSO93boGDgO+HSjoFgO+U/8pQeO+xva4QOO+79DpyEO/M5jnHZO/N0qHO2O/N4427bO/U9aV7CO/YrQJlyO/bsT5sdO/6HjJ28
PAGnOFO0PAJrgG5xPAjQ2E7aPA5B2WMmPBCXfnJgPBLSBmsSPBMUNIN/PCAW1IDOPCI86WHQPCSILW3XPCcojIoyPCgK1m+iPCq/
GJV2PC6zpXT7PDI5Umr6PDJy1og8PDmm/13aPDqgpltGPEkSKJlLPE2Be3tDPE53EnN4PE55pFMSPE+kfIhjPFLwxFj/PFVPI3Y6
PFW0B2r8PFXFdllrPFb74WL1PFqGJYHxPFtS+pMFPGFSYZ8YPGQMqFM0PGStEpT2PGTjYZ9FPGacoG/9PGbWXVAWPGmcWX4BPH1V
5FjhPILM/ZnQPILyjZiCPIdAo4/wPInHNmAFPI8BYmnePI9bhYfuPI/LlmaYPJC67FFEPJOwGXlmPJvUNJuMPJzPeWI4PJ3tdVBn
PJ9bpFWgPKFBxYajPKkGtWMKPKnrZHf4PK2OLGeoPLCDln6OPLOqVmpPPLlR6JDEPLnbGIEzPLpu2XvIPL1I7Y0TPMHcGldGPMcR
94uCPMxT0FdxPM1CLWOjPM1dVW9UPM4xlHtxPNBmkG/XPNDOuGu/PNOjqXHfPNO9WJUTPOECQnMjPONl1GJJPOwNBmLrPPBRGIX6
PPDDdZEqPPL5/WqWPPfe0XH3PP3+GGCkPQQ1520pPQaAvW4OPQqOgVcRPQ1UQlmOPRSt/XWrPRZVq14sPRcof2zlPRrxVlIfPRuT
L4a4PR5tMVDWPR8C5G3fPSJYmVEpPSZkFWR/PSe0IoA+PSmqLXA4PSoir4bRPSsclZqgPSs9FXzIPSumE1iLPS9rGZc4PTBYDZa7
PTFIm20iPTQ5lFHRPTdNwYkWPTe2UVNoPTin73poPTnOmlPBPTnvPHKcPTqwsX7jPTuniXy5PT37PHp1PUCG8Wx6PUCaQZU1PUEl
qVOiPUNgh4cLPUSpnpB+PUdZWlM8PUdmxILdPUoOb5y4PVBGfITLPVNiw2sKPVN3L4dqPVljDl6FPVyRMpQGPV0sR3OKPWWGr56s
PWjg82YQPXY2GZHcPX4a3lGJPX5B53BiPYHSqZK2PYNsHXalPYQ5N1QwPYdf2E7RPY5bMW96PY650lEvPZPBOo1fPZPzz1KSPZVD
e1/1PZcjIFmUPZs6GYn8PZw9W2esPZzkCp19PaGkRpmgPaQEVpKGPap7F1dtPaqtpGSZPaz3boSvPa3o4VOcPa9vFXh/PbD6wXs/
PbRNSmJyPbeug2xUPbnzEHjqPb+qp31yPcOnTIzrPcj0DJUpPcus9H/EPc41IZsaPdBiRpSTPdEmM4VUPdjVlXCMPd29T5NFPeBw
Zp8NPeO/5Hv9PegGlJBvPeqGIGhaPe+dYItdPfBHmniqPfFYmnUwPfdWap+cPfxbJ1+gPgW30YUePgYQp1vSPg5192jDPhEl71bs
PhJYNGPHPhKL14PxPhKg3XI8PhP4UYzgPhW8t33MPhsAhozPPhvVzFEHPh0xmWc4Ph9MeH3rPiJRoZmAPiU8DItHPicXBFp0PizM
NVLNPizk51MEPi0t/og9Pi1Jv3nJPjab7oifPjcEDFe9PjssGmJaPkEJHJO+PkhxLoUmPksXw4q6Pk5Min4FPlFkpE5BPlQQA2Ey
PlpOk150PlpyJ3ViPlumd3SAPlur0V6APl2cLFrOPl3iJX7pPmJ1hYBiPmdO6W1bPmgXeWUpPmvgM1zhPm3/MZvoPnsWF4sWPoSM
rlmjPoiCn2XcPokpHH/PPomoion6PpGCr1cLPpGGPm34PpaLRWc5Ppd4I1RmPpr0cH55PpwWeIyIPpxsN5jHPqGpqmh+PqG+1JL4
PqPvIJqxPqW15Zd/Pqfx+oc8PqgAhVfwPqwYj04BPq6kOnzrPrCoooLGPrVCh3VNPri2c5WhPr/CKHQWPsSAm4NUPsbqZp2RPstN
pmCRPsu47W9lPs+qzW5iPtMUV1xLPtPOZW/3PtSB72wEPti1gW92Ptnd8l84Pt1T4FaTPt+l/2rnPuTr4ov9PumQtG3lPuoNPmp+
PurihnEfPurkh2tCPvIfQWZlPvb18l4MPvc7ZpHqPvpXlX1QPvsgm1iBPv+ps09jPwRsxl4wPwvGp4UaPwx2xI80PwySEYaCPxKj
E5cYPxTHi2zwPxXs12NaPxY4/IwJPxe0llUBPxjGa4nVPx0z9mGQPyIdGF8TPyLMg24APybbApJkPzD8OWjAP0KkiJyAP0fEOGAS
P1BAjZ0DP1CUMoRyP1FiK5JaP1QJ0nbqP1hPfIFHP1iFspkcP17rg2DGP2YI3m2DP3SeeJubP3q3nHERP30FQ2kCP34kP13wP35L
/XXfP4PZW59NP5cf/U4NP51VTlmgP55KQIZ2P6CeRJo8P6ZwonlUP6fzqlW1P60oz3nzP7cPvmJSP7qsD4a1P7s5z2BKP7+QNnAV
P7/j5lW5P8CXt1OfP8NwOFsiP8lpcE/MP9ITX1PRP9LSrXH1P9QMtnuIP9TXPZLaP9ZiB4CDP9pzGY5TP9+T4Y8WP+IeG141P+b9
3YthP+pkfWl4P+wImGlvP+zumG7nP/Q9Y17QP/bdbmbHP/vwkoaLP/07E5c1QALwEmQVQAmeimJtQAr5mnTHQAwFOFZWQA++qHRX
QBBaj15mQBKn85gJQBkUInEbQB4mO56BQB6P2ZtiQCVeymaPQCaHTW/PQCgAL2EwQCvDQWRVQCyAgofpQDB69I7wQDFI31CpQDVF
J4LMQDbx3Fj5QDzEj2E9QD1VfWNvQD5IF2RyQEBTo1XAQEew44vXQEivioVtQEjIAocJQE5f62+fQFAS02KFQFJOg35FQFZfPXyT
QGEQNIJYQGeGJIheQGv9I4JWQGxDupUqQHZcxYlYQHnZ0p+FQH8kZXHsQINsdWU+QIziJISGQI6bCYVlQI9nyItMQJNU4XrkQJRv
/50HQJZA3n8iQJds61xKQJhKm3WiQJllTH3bQJuIJ2N1QKIdnpdDQKVD7pnGQKVuYF/aQKcxg5XjQKhXc1QbQKkdwpc0QKn5u1EJ
QKpHKJRqQKpr13OmQKqle5ERQLYy+2idQLZvVHwHQLd7smX9QL5+D5TFQL+cznMrQMC8EmtdQMDLHYvwQMGk25fSQMHio1qlQMO/
7Vd+QMgmK5PYQMhifWMXQMqmgomHQM4f9V+kQM8dklWEQNS69nuzQNYksVv4QNyon2UAQN3F+4IPQOB6u4upQODSQmw9QOIlTZSF
QOZqapo7QOcayYpIQOkTdIhNQOr3m11PQOsdx1j8QOs8gIELQPKMsIzxQQHVqWdnQQzhn5u2QQ61J4qMQRb1MpeRQRpYKoe/QRtx
Z5tIQSOaKGXrQSg5hmiaQSu0UWZZQSu81JQJQSy5Un5VQSzyP5mIQS03NHv3QS8d0IfoQTL1cZYeQTThU57VQTahw2oAQTlB3GGJ
QTl3hoVNQT7DW3oUQUArJW2bQUNv2VGnQUcT7W4BQUmngn5sQUnrQ2B1QVgZy2KpQVlMTIFtQVq7OX0HQV0C3HLJQWTCV5foQWWY
kXdWQWhweIkQQW47LVq+QXRdoYzYQXbesFmeQXdNb1+dQXiezH+fQYRYpk46QYdlYV3oQYq+BmqJQYrnX17SQYznGE8VQY9wbJig
QY+s14eCQY+3yVuDQY/801WnQZEyCl5vQZQXaX76QZSefp2nQZWuHGkNQaDO0m0xQaGHK5KiQaZ9HY3sQai4HJiRQbAMaE/mQbB2
zE/oQbTHAmZLQbb/S38aQbr2InLmQb0gt4iBQb2w0leQQb5mrWkYQb+EVVxEQcBIEZc2QcBeJIVOQcOleVkuQcYRf3dQQcajtIqt
QcwYJXa+Qc0+hpDBQc7QUXdLQdaXjoJCQdd0cGLzQd1/TmkAQd28kFMHQeKhh4NkQeLRuY9qQebP21LTQepue2HNQewDcpapQewM
1W/KQe+xbXw8QfE+n3UfQfFuKVB7QfH8bpKEQfL07U6aQfZjjWyLQfefUYkaQf4aaGqPQgfGEZvkQgiXSnREQhCXzZumQhIrfE5I
QhVhT43+QhbfBZEcQheM+n2wQhgzMmS+QhibfGjhQhjgcYvVQh7JEXHlQiCBDJBRQiIMpoAFQigr2YWLQi3yoYeGQjKClFc8QjzX
hHWlQkXO0psEQkYju3MHQkZE/o4RQkZenH8dQkcQ1oPeQkrKX1raQksOHY5xQkxavHixQlLtan7/QlNCp1lNQlS7ZVgUQlVI4Z+g
QlYJ8k6+Qlqmqp4+Ql6hVWM4QmIeX48KQmMEFIAWQmZP+H6YQmliJ1wWQnToCU6GQnd5R1k6QoHNkW2eQoexsJExQok5Y2M6QomK
5Im2QopeSlt5QpFanJCwQpJRNFA5QpK2qoDqQpR6qV9uQptj9351Qpwl22iZQpxRspioQp5P541nQqCq65T9QqaAdITsQqnK3ZZ6
QrH8GHb4Qrj8lHR5Qr9itk/DQsRP5HGPQs3quY5ZQtGUNk9BQtJPjYcMQtUGE2tLQtaWwmraQtfFJ1fIQtm8RYw9QtoONYAzQtwc
Xnk3Qt2ff3OOQuCjZ24sQuKJDJonQuR+F2ihQufv9FQ+QuvQ7WtkQu49mlprQu7V9lWqQvK75FUIQvwSvXBFQv2PImWtQv68uk7V
Qv9DZoYzQwLFAY92QwSLLoZfQwj0/5AbQwu7WJSUQw1y0IYuQw/oz112QxPNbXIuQxXGe3J8QxpGQXZXQxu9nYq8Qx1I8054QyIV
eoC0QyU5XHhyQy63ulLQQy/pKGEgQzElyIOyQzLTyXIaQzUu+4O1QzZNOWD9QzjTIm1JQzjzHVifQzqWp0+6QzsCCpIGQ0Kn1XSc
Q0bOAGpGQ0ck/XxOQ0rTH325Q1B5vosVQ1U6+FKOQ1Vy2G7AQ1Zq4limQ17RInctQ2Z8HWxwQ2sKr1U/Q2/E95hgQ3BHv3riQ3E8
qnobQ3eOV5AFQ31s4o/NQ4d8ZpkvQ47huojUQ5PLMFK3Q5jHMlg5Q5jK/2mrQ5jjmVlLQ56R5YxXQ5/aV5SfQ6oBIZC+Q6ph4HPA
Q67pX4CeQ7HG1Jv6Q7Y7umdzQ7cnNnzCQ7r8JVXuQ7yFCJzSQ8T1c1nUQ8a68GUgQ8igl1ktQ8lM65I7Q8w7Sk6MQ8zaM12JQ9Ig
tk/lQ9VJAVZiQ9msRnp5Q9y7L1q5Q+GBgIAPQ+pJUGivQ+teuE/zQ+2h0nPhQ/DM1WX1Q/kgunpuQ/qrvpZ0Q/ravnjfQ/slpGJX
Q/1ZLnN6Q/7Hwm2RRAl6CFeGRAsY6XvqRA3pF2hIRA7QvmPTRA81E3lXRBFK2nAWRBuNLYfHRByO35dyRBz0to0gRB57spkkRB7p
QFT4RB+WEGg2RCUOJ3rGRCe7poNyRChKZYcFRCitl1O3RC5xB1p1RC7QIIa2RDRax2wARDhcwIhDRDiUjU9hRDln24D5RDx5K5Ep
RD0BqnE4RD9E1IdORD+JWGTaREBJo3SNRECzH18KREIamlYAREYMDnyuREyhwWOQRE5Dk1sgRE966FOURFDrA5XqRFIChGH5RFK0
KnvJRFR/+Gq0RFSwsGI/RFYzmFjBRFe0N5u1RFkGGIQuRFm/jHfjRFyCc5lpRF55U411RGE9BHL6RGRdrolERGTUOpEARGTpdGvl
RGYqHmkSRGbS+VXRRGqIDGKCRG68ClF9RG9GwpL/RHiet5mWRH7gEoDjRH9KuJR1RIBqMG6NRIKuFI8lRIRdN46vRITcOGz8RIb8
hU5QRI/o1GI2RJJVfJ5xRJJ+ZGuhRJP66nNpRJguVFXKRJpEX4Y5RJveTIMIRJx3XJJFRKBET3zbRKRSnlLXRKaW71qNRKbgvH4a
RKoZZV7qRKuoOIPpRLFpeIIWRLMHPoFiRLR7MWqIRLqZZ4+6RLunw3MnRLva/2QZRMHL9mCiRMINE4SERMMmqVuyRMnd4H38RMsm
8FZeRNBTknTxRNa4i1mYRNlWbG5URN2sH4iURN8u5ZAERN+nYWr2ROqwyIOhRPGtqFlJRPSO/XP1RP4c4oq0RP7P1JXFRQAB+puQ
RQIQbHjnRQISdVeXRQ4Wg4t9RRBapGBHRRC5OI7cRRV2vHK9RRiNSmUyRRqYW156RSL4smqRRST7mW7qRSYUjnG3RSZG7G9fRSzS
NFa8RS6wJWIoRTVqPXxoRTYR8E4kRToBamfZRTpE31L7RUEKxlm9RUNZF5AuRUZgFHZ0RUs20mnMRU33QpVmRU95WIL2RVOy5ZPX
RVo2SU74RV2bJnzhRWQVCVl8RWb3Zou7RWjnzk+YRWkcvFhPRW3Sx4cYRW/5e3GVRXg+vHb9RXpKDnt+RXrxU5K5RXy5imbSRX1K
Ymm9RYK1pHv+RYVyE1FzRYaIb3ytRYwjw5VoRY0DZos+RY4f55J+RZSzOX+YRZkg/YmBRZ6bTk/3RacP9n64RacU/oojRak73GhA
RaltSVJFRamlqWDQRbJK+lB/RbY40JpRRblvU4EIRb292VaNRb3XuINFRb89e1WFRb/1F1wNRcBninQ+RcEAzVcFRcFSJJufRcGV
qI6jRcTE+ocNRcVy/1gFRcqrD5rERcspAIyDRc2FrpEPRc+qwZt5RdI5soq/RdJCq3erRdNGfJesRda8YmqmRdksZmwCRd68WVaj
RerJrJ2JRe0RF405Re/VwZVWRfAQrXkTRfDyQIFwRfKkEXqdRfax/Ze3RgYQVoeaRgxo2E5tRhBll5HLRhCps5V7RhDEFZiMRhF6
nFsLRhHjwHDyRh+NdXbURiQ+fGPlRiftE55ERiwJS3LQRiz8+04SRjBm/30QRjPLQ4HIRjS7soBCRjcW7o2iRj4XE2d4Rj4ykFmN
RkF0JpFLRkMMWYJvRkMZ+3HxRkPjV5cnRkwAH2A9RkyBuXutRk3IRn7mRk3xXG4DRlEkkF0bRlWkJW+7RlbCNoc7RltOe2yTRltu
rXjYRl8k/3kvRmMEYI0aRmg72XPeRmoKBY4tRnT3gHVVRnlAyZ0gRoBTsnE5RoFYvluzRoYWsGbZRoZtSluoRofWwI/vRogqNnEw
RotrAlW7RpE5F4KBRpaF9FjHRqBdeZAXRqDtEnOiRqL1eJ7qRqQtD5KxRqyWA3Y4RrT9VmBnRro1LlmIRr2MbJ5bRr8PNYDtRsVt
8J3dRsb0E2ZNRsmRsG7IRs2F4XFqRtDVb4jERtZ1nF9pRtbw3odYRtwcM1z3Rt5Jwk5/Rt9PU4hQRuEZ41KeRuTUHnSCRuaCHJit
RuexqZofRurvDZOnRuu9g3BERu8vSo6nRvBCKJtZRvWyb2AyRvge5VpIRvqkelOLRvw2ElnxRwNXPZqpRwdVaU5PRwkFv1bpRwuP
IJlzRw7jTI+cRxLjiIwCRxOXT3dxRxRAQISDRxSZTp8SRxXWvJKRRxl4qJqcRxvUyntkRyHG8nDHRyPUEVtVRye1vloDRyilPYGF
RypCTHzfRyy4R3aSRzTqkFiXRzW3S4KTRzdwT3eCRzwA42dNR0Dj2XHAR0KNNHdYR0ba42sXR0gdamX3R0klU4bwR02mXI4KR05s
emONR1DQJl7+R1FJVI6eR1JIUYnUR1OPFHBXR1Xuznn+R1YwEJH4R1fSVFevR1i4AJJoR1oPCnu5R1sQ1Wy2R10G05x+R14Vj3vg
R2ETTokMR2M3nnKvR2RaWIu6R2cYn5DDR2dCiW7RR2hTyGU9R2viL2GPR2w4QmdZR2y3flE5R3AO5JTIR3Ntx45RR3RO/HaPR4BH
83XuR4EsalBMR4GnA5J0R4KeG57oR4LVvJEvR4OJBHxKR4pVAmA/R4xeeJIcR5K3gWExR5MZJlGPR5SumoQcR5aWCo7sR59lCnEg
R6HQNVbQR6NvV1IsR6TrL58VR6UqjmbbR6vbHXsrR7O4vW8bR7gVvI9gR7kVYG9OR8AQrGWdR8GAZXLBR85Oq3ssR9ff9ZDhR99h
32d2R9+cm2htR+QHEJONR+XvkYCGR+f6YYX9R+k5fJAgR+89mFgER/E8zl2nR/YMUE6YR/mANmDmR/0dOFubR/4MWn/AR//4DpxW
SADXuWnPSAHJs5zXSAkkBVQ0SAorjWNjSAs1RojsSAuVWZ5ISBloz35+SByTi2/OSCAQv29vSCCsIoUTSCdSRJYESCqM7lHpSC74
KX90SDsRSlNRSDutCFzsSD2MFHxMSD3lal3ySEBSmZiESEPuIYXMSERuR5e1SElTAobLSExNMnpKSFHDFJkFSFKRIZXhSFR53pcz
SFeVrItYSFhgF5bjSFqnVluSSFsAv2NGSFsFPVGUSFvOs3TySF/ykmc3SGNRzYM5SGdBEXztSGd8ZY13SGov1F1VSGqLK2v4SHES
6po+SHJdV1yBSHN5Dm7OSHUCU5cmSH3fWn3ASIAPtF8bSIJw+HrXSIT+X3sJSIbfjJvKSIliQpgGSIpuHm0WSIqVK5mySI4dJpDC
SJDsenD7SJGDYoCgSJKGQWTKSJM54mt4SJYB+29tSJdjE1kOSJnccnKkSJn6+5OzSJ8hGHDlSKUWIFf1SKfKrG79SKfrkFf8SK5n
a4SCSK7lj3dGSMG+5l5ASM7gqm8iSNMBulXtSNR175yrSNhdlWlGSN/StmICSOFbAozcSOVS4nDPSOoEZZLmSOwRWZ7mSOy8KVH1
SPEdWGqNSPI445gHSPMMY4coSPSWmXOFSPWDUFcVSPYmKl3PSPjj9nLbSPwGSY2vSPyz0JMlSQDwCIObSQJFqWmRSQKBZ3XDSQTE
G4FESQev2U5lSQimPZV4SQqt6FxaSQzOcpCBSQ/TAY3NSRFA6I7MSRROsm6CSRkljJK7SR9bN4UgSSaRzmvqSScUG3RWSSkuunUL
SSpLLF9MSTlf63xTSUdO8mqwSUnk7p4wSUwwDm5hSU1bNn+oSVCC017USVKKtoMQSVOSgYH/SV2PlXX5SV/rTonXSWDrgGsDSWJd
C2XZSWSmF1FpSWbt44bWSWiGEZuxSXCtrpX4SXEcRnr9SXSoH1T7SXWqqV3DSXdIVJiNSX23Kps/SYL/LJmvSYY/sZHmSYdh5Z7t
SYgTkHZaSYl7dJCySYoWW5y7SY35CZbbSZFMlG7ESZj8cnaiSZ25iJaCSZ4Y+FLeSZ7w8nOtSZ8KgnacSaKcXF0rSaVJy3fUSab+
KYxfSaqO7mDJSas0JXHMSavpY4hKSazKJ5gPSa3xNlyISbAsC5WLSbEdRWfhSba5+ngySbf2Gmv8SbgHxI1jSbh08pZQSbmrflG0
Sbx/tJK1ScCKiZTSScE97lkQScLER1o+ScSm3JnoScdnyI5+ScpG73ISSc5jbmTTSc9gmnbLSdAbGk7lSdU9ClkCSdrcgFbUSdxT
MntjSeRsYlL0SeWUiV23SewRl57ySexyp1TpSe7lZlx+SfkJY2agSfnTQWYaSfrT8nFFSfzKhVSMSgDLzYa8SgL005jJSgNgLJay
ShDi5ZL+ShEJ45B9ShK87ZCAShRg6VOMShVB83zkSh6f54YnSiB2ZZa8SiuJvodlSi/2tY+5SjBvYYB1SjDxkU9HSjtQd3IrSjzR
iJepSj0ow1VnSkAngGgrSkDxwXm/SkNPvG8KSk7rr5hJSlatklPASlbUdIG4Sljkv1NUSmFg71WUSnAhgJ1QSnRB8JomSnU0MmP5
SnrMEWQMSnz4H1ppSn+2ZG2ISoMrd4bXSoWCz4LFSomk6HR1SouVn3gASoxHbHqhSo0R9FxMSpC8QovuSpDu2X15SpPn/1T3SpuZ
tmu8Sp1SDX3wSp3PvU94Sp4GJWTHSqQwr5q6SqbA73pqSrDECHm2SrGR92NpSrhPuXPkSr/faI8CSsAATG+WSsVLBlfUSscfUp59
SseQ6HXhSsi33XxESsjXKGsmSsmdw13USswvzna3Ss1S+n8XSs4j+5DIStOdxlTHStT5C4INStVvRmNIStXTx2NCStpcUWifStq3
+1e0St/MKIEASuHwiXhjSubfa15hSukTEFwmSvKl7FqKSvLLFmsgSvW1pJZdSvW6K2nFSvlt3FuTSvv68E/fSv9jMV/VSwct7ok7
Swj1hZmJSwtFzpCLSwvNGmYqSwvSR2oBSxLrqG1gSxOulYRJSxUEaV2PSxXico+RSxgbxpy1SxvsRnpWSyUosJpZSyjBIm+KSypc
gXD+SytflJhCSy0/l2ImSy4ZeWK7SzAmM3FySzPsA5aiSzQKz4odSzmFlXhCSzzzZmxyS07U9X8MS1E371uLS1ZkCJJQS1cHq1Va
S1hvU2H4S1jqcm4hS1uKXGOPS2PqsW6vS2Zx93TfS22qE1uMS27fHm04S2/HoFvpS3G0NpoIS3IFb2GDS3KOvFDgS3dLr4DvS3ml
kVZCS3wDrpjRS34yR2qpS3+B21NHS4jEJI7BS4zHxY5SS5LueX1XS5Oa5GePS55AJ28wS6ApxYhHS6G4GIufS6G5cnyVS6UXkliM
S6dxUIcjS6nM7ohyS6uakX7yS6u1S3gbS6xE0ZQ0S7H0OnHaS7z7XGBlS74ffXhdS79kMHNsS8iri57uS8yOalfiS80dyXxlS9bg
6VG3S9dXMWxHS9lsc05gS9lwVnNzS9tl4JCIS9vkFYMHS+MS62aDS+Q+VXCUS+jaIYAJS+6iWV/CS/I6pmBFS/NA+pZpS/NL7ZaD
S/dOOV29S/ppE4wmTAHPUIpqTAJzrVd1TATXMY7QTAVyEU6mTAYd6VYsTAZHQmJsTAZdAVpHTAbRyZg5TAinoJrsTArvPIkRTA6P
wpkDTBk7qn+FTBoODmWaTB0GpIqhTB0QPFX+TB4VNnQ5TB7o12zVTCDIunC1TCLHl4sbTCdYdV7OTCmVBF7uTCq1pJezTCudsXG4
TCyzFJJOTC2ivlafTDAKn1k3TDDOiJLHTDcBRI10TDvTaWVnTEBbYoC9TEU8j4FjTEVLzFaQTEbYXlK0TEdeRIwkTEfhuG2jTEoc
V2aNTEsKTHQqTEvDNXKdTEwqrFcTTE7erIl5TFfXAlvITFhIpGBRTFh/i2EYTFnLJlZxTF7YAZcbTGKWOmroTGX7rmbKTGb/ylfS
TGhg7oShTGwDc2ZnTG+HTXbCTHDY9FkaTHO2i5fYTHZO8V/ITHqOUHMATIFAWlMtTI8UO4DGTJD+km35TJOp0ZZRTJPHNpNWTJj4
45joTJkE/VIBTJkvmH91TJo6PlfMTJwPDoWnTJxg6FB3TJ3NIo8pTJ7ZDYCYTKNjWX3OTKUepp66TKf2PZ1/TKhjz2JgTKqkvFa0
TK8xQIb/TLDlJHtvTLF9OIBbTLNusVRfTLSwGIrPTLdFn4XSTLdV2o8LTLdsoYxBTLntdHshTL0rgF61TL/lv1JOTMIhcHb2TMUl
vFJPTMqGb2nbTMq8mpTQTMziUW4xTM5E81vXTNBEA5m/TNLw2HNGTNWOnG0bTNbr0Ys5TNfyvleATN/6WWSpTOJTf5mKTObUaVmM
TOfmDWLxTOf964fNTOlgKH7XTOmvEXdlTOwqRnuLTOxemJK/TPKOkoyHTPPNcJlYTPRdw2YkTPUVKoeSTPWfEoOkTPZAPYRITPgQ
63s4TPpo+YmxTPuJsWFWTP+bLGZATQL5hn5tTQMq3W5gTQmgKJdbTQtBi4pVTRBPXGK5TRHOCFXITRYuDlQTTR2LcHZ9TR2+hX99
TSW9KYEvTSgPqZXSTS1u+HnQTS3z6FVSTS/8+3MQTTQx7F0PTTUBXpiUTTmYL5ohTT5UCnCSTUON3nzETUYEgZ39TUlYi5XVTUtw
/3pTTUuDPI/1TUyNVGAxTU2lLWGGTVG2/IslTVyyk1fDTV8k1XUcTWBJaVccTWGwyVGcTWciZYxHTWihmmMGTWlmzYT7TWodpVFq
TWrj7WYhTWyD04U8TW1SQ5zTTW8onJKgTW9YDId/TXAkYY9ETXI2r3efTXNObGjZTXTCBoPUTXTR5lijTXUKnG3iTXZoxoM/TXhQ
rIj+TXlaX05TTX0G6Z+UTX4gKlfQTYaiUYs3TYeqMGOsTY63BFP6TZHfNoWBTZVZ3YqRTZa4AGHlTZiqCH6tTZ0FMGRlTZ5osZwG
TaHIY48UTadWvHR9TahP/FlKTavgF08UTa7FB4EaTbPZdmteTbQ7YYPvTbVOzE5mTb4Oc4kyTcI56VGhTcb6GmlMTchGFJGITcnS
n23+TcwTs48NTc62TmpWTdCAFlW6TdWRzoHlTdbYvYu8TduRi5e0TdvNaldyTdwapmeXTeLjXX3ETe1RgpsXTe3WoJr4TfB8tZ23
TfFGJYl6TfGdCF+5TfuYqV0yTf8VYk87TgbZF3MmTgyYV23NTg8+sIsiTg98lHzFThApelHUThjjjY4/ThxdGY0FTiFPOmdCTilK
H2edTjOZZZPDTjRSQVqLTjkpd1XFTkAaMpxsTkHmCFfaTkLK4JAhTkeKKGlqTkhb5H+XTkhuK1UpTkh0yVZGTkzxF4CmTk2CT4GU
Tk5/nHn5Tk+QGYxOTlDEaX9ZTlUkDGyVTldMIoEhTlf+OIY4TlyzUX/GTl0rCXpETmE75pGTTmYOEFe8Tm4FJ12iTm9WuXxxTnaM
fGfbTnnMmJPsTnzlqHgoTn4eVIANTn/TKGNFToC8/1SnToFR5V5BToQa1VZEToRdOndOToYLbWpJToeUTlMFTogn4osBTo5BJ3rI
To5oF46wTpE0LnitTpE19207TpGcMHfFTpXvenkDTpc8LJLrTpqvSVgMTprxo5pnTpsLfE/gTpt64YZUTqyXy5sBTq0qRZj3TrW1
S2ZJTrdo6ofrTsMPL1/oTskDMYCBTsmGCZwWTtl62Fs+TtoIC27HTtpXUn8TTuSIWGurTuYzp4uZTuZzcYWkTum/JIJkTu6VE3L/
Tu+FtXwSTvJUI18MTvUuqF6sTvWqXHpITvjs7pSvTvwtA1tNTwKtcY3eTwOM+l1tTxMp2Ja+TxhYc2jrTyFrHE6fTyKHjXSqTyKa
3YBFTyMDoGhVTyMjl2vjTyOJ4GyxTyeZW3ZKTynwElUcTyw+kIysTyzwAICKTy/GM1nPTzUt74+8TzoSzoX5TzsOG5vrTz7OylS2
Tz8QqIkpT0Q3pE+TT0vY2nBzT0vsyXLzT026BZukT1O3JZqFT1RJn2+4T1WFhnxaT1bt8W9FT1i+R13uT1roeHPYT1t0yIU+T2Kc
xFvPT2Ng5oKST2lXjF2YT2l7Q4IlT3ZBc2GuT3it1HQ2T3mIHmUXT3s47IPXT38VW1/HT4Apj3J4T4HgDV4+T4IWSlWkT4TTUIB9
T4YOfmr4T4fR554vT4woroSoT4wqu53fT4/PG2PNT5kLeYrXT6EFuGwqT6LYInaLT6T6tmPFT6ik7ZhYT6yzWWtHT63FuFFmT69i
xoBjT7ko1Vh6T7pM7n/cT7pfImlUT8cjAYWZT9DvuFxIT9Fb/XokT9GYk2FvT9+LkE5KT+J8RFlUT+aLuXjGT+j21odaT+o/2YvB
T+wo0IQoT+xDMlIdT+0DKF+xT/GQzlPNT/J3cU5kT/V8WnU2T/iKs1R/T/zTznt0UAAHDFZQUAHobVCfUAOW65bOUA1kRoeTUA5d
Y3OkUBBymnSQUBF5El0TUBpeD07dUBv8eJz4UCLe2V/sUCc1Dk59UC7swJZHUDJAC11KUDTz5J2FUDepeoxZUDjlfJCHUDoHX2Zy
UDw8fXR4UDykHZ5VUD9WLWSTUEUlilX/UEwXt3e3UE3bsZ3rUFBFDJ1rUFCCvFflUFEgapIJUFdqsHKaUFee+pagUFuB8ozBUF4o
lU90UGAX+HiWUGILu1kjUGIwCFJEUGoiNlJGUGp4WVbCUGytkmLpUG19vneKUG/0wWlXUHdP+ZmaUH36EFTAUIDAMVC+UIjrl4Du
UIoDJnyZUI0yY3HIUJrSFGdIUKCRkWetUKLMBV5IUKUheHd1UKYP+5TZUKbPJ5HrULTwS3zsULZkB2IAULbkxWbMULhpLpWQULh6
T2JcULmIA56xUL5ZJWZcUL8ZeXjIUMCGGU5+UMLJGYCZUMO/UpyKUMblEpyjUMlvT1yxUMnqFmMZUMzYkVpBUNJqj5EYUNXmsWBs
UNYk21NMUNi1JoHRUNnj0FfzUNyaUX6PUN0++XcGUN2572oSUOZDtoS3UOeQs3k5UOmLGJ0dUO3fbVCeUPLxRpJsUPSZZWZFUPSz
dFq6UPgzf5phUP8qNpaeUQJpM2xgUQX6jY0OUQjJFI6IUQ0m4E/NURA1FWj5URRJC2E7URTLX4tSURlBQVVpURmRfp3SURmY+lde
URmqdpdnURnqS5yqUR9S14h9UR+/z3RiUR/823wFUSUVxFlIUSg6jWfsUSwMv55DUSyf1WjYUTdxH1zFUTmrjVWwUTueOnggUT94
tVwQUUG5J1kEUUI0MpLUUUZCT3lOUUawGVcYUUjN0WZMUUl7NIYcUUrFfll2UUrRCXp3UUt/ZmFhUVBn3VcGUVHkZZv4UVeD3JQT
UVprB1Q7UV31C07bUWX77557UWuLNGmEUXeXM1uEUXnT4YYJUXoLZE4vUXtQ6F9OUX4pCovoUYId3Y6yUYcRU1igUYjjgnYIUY2R
a3AoUY6zNmt6UZJ2apbnUZMc1pj5UZefLJA5UZ/itmlxUaGeOYY+UaK/Lp4/UaSaImSJUan4kV9fUaolkFoVUau6IIa6Ua9yDZns
UbKCBJIoUbKLG5xIUbZmfH+sUb3RalV1Ub75ElFFUcCeJJDGUcYA+IFRUcd+vXtOUcjdNJcRUcmUtXJzUdZst5kKUdnpHocUUd+Q
TW/tUeLDj1Q6Ue0ft04EUfX+DYrkUfahiGzcUfaoFH06Ufg9np3hUfx/IZBhUf/Tk3UQUgiv9VfOUgnvvJ4gUg7JRVpYUg+M+H07
UhGiio1/UhY+K1/gUhx8AlEtUh2UFHlJUh5tVmvTUiBztVkeUirpmlq0UisvnlU8UixXHpaAUi8HyY68UjbzXFEwUjdSFoaaUjlD
+IM+Uks7QnrMUkuMCp+SUk7GnI8dUlAxE4/3Ulc2AH7cUle2D4IDUliamYIdUlprNXSVUlynT19QUlyyW3FjUmIQaldOUmPN53RQ
UmnQao+QUmn+rFvLUnAAA2+ZUnK5pWHwUnQUvVaUUnWssW3eUnWvBVhVUnfJaHYdUnpJLo2PUn4A03aYUn9r5lFYUoD84n4qUohs
eYLyUotv94pDUpUdmGQlUpdXRWVfUpfwuGSmUpm32XGaUppWZX4RUp5Le22sUp7jb301UqJomF0eUqWiLms5UqjijXDeUqzwgJIX
UrGMHmj7UrM0soylUrSa74y+UrbpMJztUreqW13iUrjscl96Ur7qb5vXUr9lGGToUsR4Q3oXUsYY4F1eUsYsU1naUse/PX1HUs1V
R3J3UtQVUXZpUtarpXAxUt3oJZFXUubYPZ5OUulQ+nuDUumAK35YUunQBIKvUvHUN5Z5UvR7pXL4UvXQ2IHGUvdZsE8SUvmqNI4e
Uvp6351dUvtQdl6gUvyFpl7HUv3bx3nBUv5MbX97UwAG3IEjUwGo7G50UwM1qmtFUwx+pJh+Uw0p3HAOUxBq3J3VUxWiWlplUxzg
gmTEUx5d7VoBUyHZ1nhBUyRCqHsXUy/4RFFRUzZBbW9kUzm8zn85Uzrq35PjUzwwxmhqU0CAI290U0JerJQcU0zUv2PnU098WIV3
U0+JG11FU1npqF/iU15Aml/hU2H1X4WWU2MMUW+bU2VL4WvRU2Wt/2LLU2ZrbZghU2maaFHGU22T8l3IU2+Fk2goU3Bi03UeU3LD
pF/8U3iNrpMbU3tnZ2XVU4OQy4QpU4bTzIAGU4j8IYe3U433EImWU5joHZ1UU5l9aXvFU54vkE4gU6OMcXUiU6RKsZ8vU6YXwpmr
U6Y0UmwZU6br0Z6LU6jMMIg3U6ypdnRdU66NB5N5U7j1K5YmU7pFVYBKU7r4D1hqU8MLDIVaU86Ea3yHU9wIYoO6U9wRqXxhU98D
EWXlU+SqH5t2U+nePWXhU+z9GIkbU+88kE/RU++izlhoU/MbC1hKU/SORXCmU/TpeouVU/kQN3lMU/ogxlOtU/th/46gU/9kTYHg
VAGZtHcvVAezFm33VA6fVo9sVA992ItUVBBU31GpVBIGAY5GVBKvzFxtVBUdxGHDVBbRnVy5VBi82pQ2VBmil2moVBrcvZGdVCUP
sXppVCgNtGuKVCj6VmFoVC3tnX4zVDZeN4YLVDbcEGzgVDjGDn2kVDkkXFOGVD3OCZCgVD8sD5UwVD/Vf3luVEK/NoNlVEf5EmtE
VE1HqJoOVE8W5XJEVFCp618EVGClHFWcVGCzPZeHVGKFYF15VGMCBWvIVGROym2xVGRZ5HS4VGujYVNeVG+BIJLAVHvNIlNdVIIM
+lGbVIJcZJeIVIgbsnD/VIjC7JRYVIjbqGVmVIpr4mgqVIsvM24QVI/YQpIhVJNMTXIJVJP0OXleVJiYCldkVJ4fwoDiVKNJ4Vgi
VKNq64Y8VKkrZ10xVLDRWZKmVLEj44rpVLclNI2qVL4/hZ1oVMFseYAkVMJvGlDxVMjEgoIFVMpzU4H2VNQYUXiDVNXJgZC8VNww
sob8VN2mFGk7VN3VA3iiVOKZqoIQVOQzQXmKVOUlbVZjVOWApHbOVOXNP3mGVOgfplGYVPXBM1NtVPdPwl3qVP2InFICVP6gnV5+
VQYuZY8JVQb1u2rDVQe/tJ1TVQu96ViZVQxJMlnSVQyGNWnBVRT/cY+CVRePkXJyVR4iGJmwVR/0XWqSVSH03I8jVSJan4dWVSJd
SpcXVSLuHm/2VShVxWwIVSwqgFVgVTQohYD7VTYsmllMVTf6bFCnVTgX1lF5VTgoTl88VTmVgJ8zVUAPt1+BVUUn3IjdVUo7QU4Z
VVdRz3dsVVldEIa9VV8RflR5VWGwpGiDVWZR7ZQYVWZdtpLIVWekI4jxVWkxp22cVW12M5YoVXLFI13fVXYxu423VYYfy5S9VYaY
rJ6vVYbRJFvQVYc8HVokVYuzKozHVY9IjI7fVY/KgJO9VZSpYFvbVZki8oQrVZmImZMMVakf8mNwVarh+JloVbTtJHvmVbncD4WP
Vb2mYI3/Vb+/DFjkVcdfcJhUVceY7Ge0VcxEbWThVcyBaW03VdAD3U+gVdbqBlhmVdc20XptVdmQW5TCVdz0nZLeVd8gNJ+QVeFs
bFDrVeRKGo3hVeeQXJk8Vejuw1iKVepfjFvkVexeBE8YVfdSFJ06VfgYXpmbVftq1IgdVfva9ncOVfx1t4guVgIqIYtfVgTumZgl
VgUzm4RCVgYpF5c9VhBxx0/SVhGJLWGTVhMg0JHMVhb6coJ7VhmsPVezViDlrmlyViMYhnGDVifRyJ9IVihx3pp9ViuXKXTaVjLr
ul+/VjcLoogEVjp2HJuvVjun54UEVj3KO2wxVkvsvl5bVk0/IZ5LVk5gEYVcVlUdtU52Vl6xN55SVl61rJx8VmBi334cVmP1enM0
VmRrqJqAVmoAcY0KVmreL3X9VnRngHO5Vnfw3nyGVnqQ9Vw9VnwLpn0zVn+jymXqVn/6NU9FVoDeD1dJVoMUE4zTVoNeIYjCVoPc
eXFBVoTQlHn7VoYfr10GVotLrU77Vo6UBVCQVpKwdo2lVpQf4IU6VpQ7zZQqVpRMn5HkVpT6qmiqVqRQXWoqVqtt+1nNVqtzzk88
Vqv0/XMcVq0AlVSkVq7gcJEBVrC/yn3/VrHs2ItCVrRc029MVraQ546BVrgxKHihVri77m3DVr5MZXmWVr6ScFMZVsCxmnJOVsLO
imYrVsbcA15lVsobPJbeVsvLPopvVs5s4IWmVs9Wn5aXVs/YfGhwVtiEDY5aVtlVm4+JVtzAs5wmVt0a1HYDVt3nn1bJVt4sZVtu
Vt/rPGvfVuKXDE/BVupzrV8rVvFsQFa2VvHOJ5mdVvxSZk5NVv04hY1EVv4qbIHvVwDUZnGhVwI+dGigVwNGT3MFVwbQQlYjVwf4
93MYVwgM7IDXVwyXVVJ+VxA3t1PDVxJo0nhKVxw1aZbEVx0ze1xRVx353YPDVyvW61DXVyx2MXUGVy4RsHyjVzZfzpTuVzbgb5hO
VznKGnZJVzsncpsGVzuHdFXsV0D+umgVV0J2yn1gV0NAHn6DV0hnyGKQV0oA7lFeV0xVopRAV06jdoZOV1PAU4s1V1jq55KQV2SG
VVBvV2XWV3ObV2hT/pTmV2kxcowuV27ZyU6hV28JyXMbV3BBFotBV3PuiWGyV3SYFoK6V3rVjmmPV4dLnHepV44LKHYOV5HuWVyj
V5XsUIacV5vbDomgV57TgH3oV6PZ0JW2V6QSN1NPV6j5mVsUV6ubrIx8V67+snmFV6+YOF+XV7HfF3tKV7rQQm1+V7tHCWPzV7xF
VWf9V8JjHVZwV8g/G1k4V8qr3pICV8zFXVV6V84kLIAOV8+uVX0ZV9Bn1VY8V9HA9IEGV9i2SZ6DV9kRSHSuV9mBrXgVV+KTBmuM
V+QI6HcXV+gduXIiV+qBek+JV+0uX1IDV/cDlYm5WALiaHjtWAUzrIz5WAXoS4pQWAXuII0sWAY0ZX+cWAaFMGHdWAiJaG3VWA0N
31fvWBJov4CHWBazIX+5WBcPIm8hWBmJlZf4WBq8B3tJWB7GlVmBWB/TEJeUWCPmQ5pyWCXFV2fyWCZPf26oWCeijIN3WCinT4Xv
WDFGKXOxWDLNLoWoWDg2mVgYWDxtpZVSWDzl8lv9WD/tkmvnWEEr22imWEVeJYx0WEjBQU+bWEjxuGoWWEokTJvUWEzN4npnWFyj
Jk5AWF5j22pBWF9GSG9dWGElzX0wWGQhW2vZWGTsMlvZWGmR24+uWG1bsnF1WG9ARoT9WHYQgpiHWHgKsHywWHgMj05OWHtm6VVi
WICK2JlTWIDdXXArWIEiRYneWIYwxZiPWIbjY3gtWIdbvpTYWId415yDWIqBlXvzWIsUYl4tWJJydXmbWJbhpmpyWJ8rwYADWKMn
xZhbWKVqKIsCWKciMYrlWKpOh1VbWKsa4llpWK1Bs2K3WLDUnpSQWLjYrZ9tWMAHQFbYWMJJ253KWMXzL5NoWMp8IW6FWMwYaVWT
WNAaj4AvWNivwVIaWN1ejGIWWN1zUprVWN5RJZFHWOQ8h5R3WOS1hXWUWO45InPTWO84ul5YWPKyVXdCWPLyu3MxWPQK6lz1WPSC
wVszWPr6lm0HWPyb/FPQWP3z7GXaWP4A13KpWP8tsYuuWQBQxV6jWQJB72WZWQOprVG5WRIRN25YWRL4vWntWRO2JWXfWR9IcVG1
WSL8dJNqWSPnv15UWSvPlW7fWS5SYlpOWS92eY1oWTTEM2VxWToZNlvGWTpwoHj+WTrUxVkpWTw+Dk6nWTzjCG4gWUE84oJTWUE/
Omo0WUFz15SgWUReg4jkWUTTPoWDWUbMrZwtWUdX+lU6WUs0apDRWVSsUJNmWVeCbXaDWVoN1n6ZWVzCS5x2WWc4p3bhWWsU0pDl
WWyIIGm4WW3pV5ooWW5jEGVXWXFfy3bnWXF3A3GyWXJ+X5DPWXKQjo4lWXgUY33BWYSjaViCWYlWMXq7WYodIonlWZNJk1SrWZc6
3JjWWZoUTIQsWZ6ZZFegWZ8YiWwJWZ8hhGfmWaF1rIl0WaKJJI+KWaNgpnmuWaPk2Z4VWaQTxGE+Wayud4LIWa7Ls5mxWbBcSnJS
WbH8FmrcWbJb9WFpWbXlVpGhWbnHR5aIWbrA2WxoWbuo6I9MWb3vh1xlWcGhG25zWcHi0lClWcW2SHUFWciiepMUWcooy2C6WcqA
d5l9WcqgmGFLWdGc1nn4WdTs4V2bWdedcoTOWdl3LVKNWdn8EWVuWdwPek7YWd9mzJPvWd+5ZXjLWehfLm7CWeo5B2W0WesiAloJ
WfAsdmRZWfIlNpW3WfiYkVmnWfxOCmf0Wf8phFHEWgAIkWhxWgAo45MTWgZIjltbWgnhXF10WgrTHJOmWg0sWVC5WhZOEZjpWhZ7
143wWhe75H1tWhrPZ4r7Wh/e4mP2WiB/FG0VWiSQcJKuWiUtt3FDWihiAIiZWi9YG3sGWjBqFmFZWjQpvHHiWkBdK1T8WkEbYlPj
WkIVbVljWkvkBlBWWkwWmlw1WlTCzlT+WldeK3nmWloMf5KIWly44oAfWl+aQVqmWmPdEVlBWmX8B5IdWmkuyJD2WmlnYJlaWnAN
foA1WnBoiV0gWnGoJ5mzWnMRQXoaWnPXmH0mWninCYOdWnivQ580WnsXIZ1bWnv+hZATWoII652VWoUjWYBnWoXCv2IiWo2fw1rh
WpPCSFDFWpVarpmTWps7mpG3WpwzSZMvWpyGGV5HWqZb935EWqdSnF8iWrfCwW8vWsIrhJ58WsL5M1QIWsPIiJQ4WsiEcU4+Wswx
BpzMWs3pZ3+9WtMZSpP+WtRC9YsNWtt4J3v7Wt/fSpQvWuI1h4pHWucM3ZncWuxXfH5iWvG0n3GiWvKhm3BTWvK1zFGCWvUSonuM
WvY7zU5xWvfN2JIZWvobYpo1WwAjolc5WwhW1pt/Wwm7rYdtWwz5x430Ww0TUXDAWw2TeJEzWw5755GKWw/NgFQeWxE4iV+pWxF3
zVs1Wxh+65oGWxicb1ABWxlhUnZAWxrRYYdoWx4eRViYWx5rIVYdWyW8kY09WyoR7odNWypcRlK+WyuJQozQWzJjNWPkWzNc8Wf7
WzWE8HUKWzcjF3YaWzkJAWEBWzxh0pnZWz3hKpqIW0KU82TbW0S9+JVhW0Ub9IFcW0knc3jmW1Ad/FKUW1FLcIYRW1k8lGNtW1+h
fGfkW2HLa4N9W2LI3JqwW2PGonQiW2SJm26iW2WycXAcW21oqZf6W28qLmKzW2+Sop8XW2/3gIJVW3PU8VpdW4PrupJvW4TVnITZ
W4ZySXRqW4aWJY5hW41tL53jW45BvV3xW5OLxVeMW5eYT3AnW5vcvGciW6GHplDjW6Xe91khW6Yp3nIXW6Z0k3zzW6tiRHF8W62T
wYsRW66gSmpxW7Ck/V7PW7E/ZIjnW7JrM3ffW7gEBVoFW7+Qwl76W8MdG2q9W8SWlGo4W8gyPVIWW8j1YmaWW8m+VmPLW8r+6GlH
W8spbo2xW890JG09W9Fg2ZuWW9LmcWm6W9Nq4GYNW9epS3RBW+cXzWH8W+ktvVcjW+0qUWicW/OS/465W/VAvl2jW/Yb4mQRXAIY
Jo+UXANmw2GlXAOSYXc+XA2NgHb+XA3wF4WfXA+XO5tyXBNpc3kxXBRXuYibXB7lJIEdXB8uGW3BXCZYPm7BXCeaTHkjXCoLhI/x
XC7+RZGpXDjSjFYvXEEoemCxXEGOqV1ZXEHWPHsMXE7X0pVdXFhHbWUWXFpk0VjUXF9PEWbpXF+x51dqXGBhQIGuXG3C6HCjXHJy
jl6ZXHQUSH1pXIhANWCaXIsocoizXIzN7288XI+iJ2/VXJE96ZPuXJGo048+XJHjOVIIXJQhOHiXXJYdqoAnXJmP6YWbXJyZS2Nr
XJy2hGxSXJ/U6259XKE8PlTlXKMkV4/9XKVD02C8XKbaiWPZXLPjl4SLXMFn758OXMIs/GPdXMXcEph6XMiDtZXbXMmdx4Z3XMpB
6nrlXNqckZaRXNzUCYuqXOoqgoYXXO3rB3Q7XO7RNZF1XO8DzlfhXO+FgmoMXPT7uYiGXPatGJvaXPehz3N1XP1Nd597XP/XTX61
XQMfrHJcXQPbuJv0XQglVFHjXQmMBGrMXQpSlYA3XQp6RHNNXRBLjVGSXRRobpG/XRYzLoDEXSq312TDXSsMrXI0XSwpCWa5XSzq
tX1vXTdGlHTUXTuHiV7dXT4TvJlQXT+bmVAZXT/6B5sHXUFUMWRNXUGwZ15jXUQBi5d2XUVnG3F2XUj3G11aXU1BKVHfXU5SPX1l
XU7262iAXVF5K4VRXVZo/IXLXVcU24NmXVu/EH+IXWKEeFztXWP7N4LBXWkFOVWWXWutHIgeXXKzG1VjXXN6r3LIXXaDYYUuXX+W
FHosXYA6OIZDXYTRG06SXYVfiJp4XYdmlJMwXYfZrnRMXYjXmmwbXYppFZHBXY76fGMyXZOo1lw/XZaXklxAXZymGGC9XaQSBonm
XaXAzlW0XaX2aWpMXakTrGVIXalTXIDfXa3zjnWNXa5VsWyDXa7Mx1bZXbBwBmUTXbfWh2GqXbpq+4alXb/9s1OdXcDQYW1XXcPK
zJcJXch9iVdmXcniqVMTXcwGU5eKXdK3qmw4Xdb9dGI8XdlHUotXXdxB02V6Xdx7X1u/Xd7YD39gXeAcPV6cXeCbYn2DXeFHj4Mi
XeiUDFgVXe2x3pSOXe278X6+XfLeG4jPXfQadY6fXfRzUYISXfZUCGnzXfeUFZcqXfka+WJOXfnoDlA9XfuK/I61Xf85qnGvXgEW
YIeZXgPPSFJhXgSDsVb1XgTPIHMJXgYyi1LwXgZW6ZVEXgpC5oZ5XgwC1o6HXgy17m6fXg5ACImqXhHf8psLXhHhkoJlXhMX2G/q
XhTsUXImXhcdcnKFXhexjJhcXhmQnXXKXhmn8mf6Xh429XY5XiqTsITjXjCHHXZDXjkHqmloXjx+YGZkXj+L3mH7XkLdkljiXklS
alioXkvpV12CXlCbK5FrXlcEtFlOXlcQem8MXl4nNYEBXl8qcJ6IXmAFJXZ1XmRk02KWXmohx1AeXmuEkmUPXm3dq2FKXm4o0FNy
Xm66LJavXnG3TWI6XnPN/WCIXnP6TmktXnRw5G2yXnqn7JL0Xn5V552ZXoEVpWkkXoVpy1w3XoWKHZnnXoXsSZ07XossWZGkXow8
/mD5Xo44zpN1XpBR/l5aXpb3S1BJXpt82ZK3Xp93gZ9hXqFOd1+oXqQwS4rMXqbr3H/2XqdjXWMTXqkR2XYiXqunP3TuXqxp7ZRH
Xq5Gc1LYXq7MgZLVXrbaynWaXrcCC3ZLXrcwTHcmXrdFfE/9Xrg7VHVCXrg+Zn94XrkCLpRhXrmCXZw7XroUWWFJXr3JTXCvXsAI
5mh7XsQAu11fXsX5Vl1JXs0SsmAOXs3NzoyzXtMah5QuXtQ964eWXtVTim8DXticjXKQXtv+EVWjXt6BVE+dXt+7Q5mRXuD3pIpf
XuD841QYXuHDjZFVXuNotXJGXuPosZ+NXuWboW3gXus/0GLvXutPFY1rXu2Ov4HVXu4NHYagXvIAxZ6pXvuNTHQGXvumBm2+Xv8d
X3CoXv+XxmJqXwAyz2nqXwDaWItuXwHC54BaXwWLm23nXwh2x2r0Xwp4b1n1XxTmsFbIXxqRTlF+XxtxeprFXx+7foo5XyRVgHKb
XyTGr23yXyVy3le4Xyli+XFpXyoaiWz/XyshH3suXyuPRZkdXyw7QFq3XzMk52MoXzhablUQXzyr62l1Xz8H8WWYX0LZWn5mX0RT
xJp3X0Vra4ceX0lSz5HGX000olB9X1M+m3oYX1eQJYvgX1puhlsYX15jbmY1X2Bt5pEIX2FSvYMyX2Vh32g+X2W1N1vJX3DCX3rv
X3LJYVCMX3VnrJY/X36CzIpJX4Es+IJHX4QwZHbbX4aCaXZIX4n/mYUqX46t0or3X48qsVziX4/na2mpX5GDYGUJX5VQs3iQX5Z7
Q26VX5xV61I8X6WW05djX6sYA2qjX7ehOpbiX7iKz3D1X7lyQ2WGX7u/vWTVX7wtoHByX72AsIXgX8dQ6XT0X8kkRmSIX8mgtHYH
X85GpZh/X882UFq4X9Gx9WtyX9NFf4pWX9tgPlm8X9unPnziX+ZLMXcPX+oo2FlQX+wmxnbPX/K1EJ9oX/QrxF7wX/t4lHm0YAg7
hVZ6YAocTHljYAulUV4hYAzed2fYYA3lCF5QYBBX8VWSYBDtZGFcYBO8TVOnYB2jqH6gYB7p4nJJYB73OZZwYB/LTpvIYCJu0ZKp
YCKIVGmOYCPC9lrMYCWy95ZEYCZv7XdAYCsihYiTYCu4/XHVYC5SKGhgYDf3uI8aYDkVvWS0YDkxaHqoYDsxgYXEYDyH/W/GYD0i
FJ4EYELpcHGGYEndXIzbYE4ZPlCtYFFUdXqlYFZtKmZrYFbPEFLtYFkprnLHYFyMe4MKYGCvZnIHYGD5m17iYGKqaGmhYGPbZYLz
YGR1vk6eYGUdaV3QYGXyq2uIYGgginn8YGgmT3KfYGi6PFnEYGu+sm1NYHDDD08TYHIhwFq/YHOQUpOlYHPPU4u/YHSUs1wSYHx0
qmDNYHzlNFJXYH1INIKJYH3AM1mKYH+5vVqxYIKru2IbYIN5tWKfYISkLpWyYIUWVZKTYIgixliJYIrAcVWYYI3dU4fzYJ4Gd5RP
YKi6el6OYKr+8VypYK25Y3w/YK6FAYK5YLBc23KiYLhlJIPaYLnRb1xzYLxWSGhkYL9q7HP3YMDj+pUKYMJlJWNZYMMY7G5IYMZC
p499YMbv4nyUYNZ9H4khYNeMTZ8pYNkq/3/aYNrRWZwLYN8ZB4t+YOH+CVaqYOUvf58AYOqPq1IKYOsffmHoYO3eFWUoYPIda2Pq
YPkKeHPzYPqE557JYPt8rIrHYQGf75TrYQXAJ2sEYQlPC4nOYQqdE1bBYQ4KqnhAYRBxlnO4YREDpXenYRGwcnGMYRR3+WgbYRWi
ynn1YRiI0oMdYR1t2WLiYSIKBZmlYSYUMWdXYSaTUosmYSbrcobmYSlOYGfGYS/qNFOuYTC1/GyzYTFg/XKSYTMIsG2ZYTQSQ2Hz
YTy3o1NhYTzeSnxZYT3YwoBVYT/Dgmg4YUR/Q3owYUaufVV0YUbxT3YeYUb0JZcNYUeKuJg7YUuOu2G6YU+IU1qHYVIbwpXrYVJA
DG2KYVMvglSYYVRfyFI+YVYO32cIYVcCho9YYVpI/lgGYWCmMHzYYWMuCmdbYWYtRnQlYWbv74G7YWv5OZ8mYW2ZEX3FYXGUJoXW
YXbFF32XYXkZK4A7YXsJS30gYXtuMnOlYX7GkZjCYX8ZUVyWYX+ACmUKYX+5h5RfYYajC3FOYYgKApYZYYrXQpiDYZHndXZfYZbe
WFxkYaEdh5guYabGS5scYanQ4VwAYa2YRYbdYbExq3meYbU91X4+YbhgHGtlYb1bF18+Yb2TcJEnYb7cqYVvYcOuB2/cYcURKGTO
YcdFyFc1Yce+dXSjYcjjJZ78YdEKHVswYdKOuU8XYdOIX5VuYdQTH46CYdX4NXGHYdY54W7tYd13g3ntYevoC3BPYewC/2luYezk
Snk0Yez2W2EOYe/eElD8YfOJSp5YYfYzxJHSYf0fxHKWYf2tVlD9Yf3iGmu0YgXIF4w8YgX1I4tyYhGO+lYyYhJJjZ3IYhKAgZPo
YhKMtGEMYhehH4zJYiNPfnn2YiWgpI8kYic9BoPOYie1SmIcYipVflLIYjDlCohkYjPQq3X8Yja7GXG9Yj8KlG8YYkKKe4tvYkMe
i2PMYkQqrm4wYkSHmXCVYkVgEFzjYkiBJ3c8YkrC/F0/Yk3zVVZTYlUxyJCtYlu+pHBQYl0gmpyfYl3aoVpmYmRWcY+hYmX03XEc
YmYh+GwfYmrOtWBXYm93EXgFYm/avZXfYnDGWY5eYnSyElQ9YnUPfmm5YnUQ6VgaYnXdUowKYnjpQ1uIYnnPyHTSYnpTwZJUYn59
EojSYn78wHOeYoF6X54hYoGtD1w0YoIaSIUKYpaW5VVIYpnvymfnYqHGI1v7YqHSilBIYqbKZmoQYqdFwFzQYrFZKWXbYrGBaIo0
YrNE2JYSYrWwWIYOYrnMSmZdYrqP/lS+Yrt8IHIDYruxvIZiYrxe1YbpYsJy+m9QYsVaJGV8YsYdD1qaYsdAU120Ysj992QCYssb
BF+cYs5kRF92YtBlops6YtEVqmrpYtGaw2l+YtHa4WCAYtVHXYyhYtWiuZOgYtuOQoYpYt10IVKzYug6eVp/YuiHqk5XYuu57lSP
Yu11nIBEYvEcsH39YvXJU1/YYvt642J7Yvx5ok7vYwAhXG1vYwBLVJz+YwYTWIiQYwrIspRSYwvQ3FncYwxdo10KYwyU2pNnYxKN
HHQzYxMAWpdTYxkUKYhnYx7v+1f2YyEkXmX2YyGM2l7oYzHel4s8YzU7lJu/YzhQG2CgYziXR31VYzjKHE8fYzkKTFlRYzpLPV+7
YzwMY4MrY0lghH2EY0pQypHTY00+h4E+Y1Zu0GRjY17NJHRfY1/bEWCnY2OpWZ2LY2nqxYxFY3CRqZIuY3CcDnl2Y3MlOHnSY3g3
v3v/Y3r0h5DFY3spqW7TY3zAq2HWY33uNGnZY35C2FKnY37SGIYyY3+1g4DHY4MWLFKTY4gF+WYgY5BVWHolY5CLIJbgY5ZfPlCS
Y5f/AIH4Y5xfFnGdY54HA3hJY6O815UgY6PXw3CCY6qpOmOvY6rgXXvoY6zrKlInY6zxU4xJY67kXZI9Y6/8O26QY7NDEZRNY7Pi
MFXmY7iKPmTNY8HlRXKhY8KpV2/NY8L1fW91Y8Ow6J0PY8Sf2p82Y8qcqpjMY8xzopJ/Y84pOJorY8/8Q1+eY9GChWf5Y9KfQHOu
Y9f5bHOjY9rSS1zEY9yyM5yhY92q0G3QY94Ivl2xY954yFC1Y+VCgGV1Y+o4bF3VY+sV5J2/Y+2B0o9lY+20HnH7Y+44lFXOY/Ei
xnq2Y/UQk1qsY/XOd4OXY/Z3cIMxY/hn7G30Y/sJ5IQEY/xVfpBlY/9EpVH+ZADkFVnvZAIJbnU3ZAYaj47uZAe20E+ZZAvuzlje
ZAxeLIdLZA8ginIEZA+DH3Q8ZBLj34cxZBdIEnZZZBnF6Z7gZCDO0pvcZCM23m0dZCajRXLXZCpRf3AlZCzYvIH+ZC8+koVSZDN6
8XyzZDV/e25yZDdiOn6jZEBa7o+yZERe84hWZEo7qnAeZEy9jZEMZE1DQ1YhZFbyRI/YZFiUnlwIZFoXbGQgZGEmpokzZGtFVmsa
ZGuFa3yoZHFq0IW4ZHNHLHcLZHTeCX3QZHWpVF8ZZHmItpHIZH60anZoZIDya2xeZILEjI41ZIOWyZCOZIcbQHQxZIshEm/8ZJLh
k1HgZJR/QGgmZJovs3NvZJxDWotbZKD3+VL9ZKdeEHnEZLHKNU8kZLH4slRQZLr6P59kZLy034iRZL4QBWSNZMEjHYtNZMfeGIGx
ZM4+bmTMZNEbmXAKZNQ0dpLYZNVv/41qZNnZ8lKgZNxn0JgwZODbxpW+ZOMMe1BwZORPrk4RZOTO+oevZOltNljWZO1aJHPuZO2a
vnF3ZO58FphiZO84fJ5hZPAAjofVZPKcdXQMZPNSHFjfZPPH+1AcZPR7NIpTZPbJSnOsZQSOc4i9ZQScL1SAZQTG/Y5MZQjCDHI5
ZQrbrIeeZQulMn5HZQ5LAJx9ZRwDn36qZSainF4QZSoejXphZSqbtIkgZSyPZZObZS3X9VonZTLuL4EOZTs3eZpmZU2qI1L8ZU21
T1k0ZU3TtZTlZU6Nr4YhZVZAYH7JZVdKymx4ZV5XEXW4ZWG8XpIaZWhejlnJZWrHxV21ZW9+RJJuZXCMWoo9ZXEtGV9IZXYrwmWL
ZXe75o7SZXfHc4OZZXf8MWYOZXxi4o9dZX87KHgZZYAKYVt+ZYwcXFfbZZnmXGfFZaGJsHgLZaLdbIfkZaS4TV+GZaiqzWuJZait
elgdZapSOWHEZas232x9Za5yJoyVZa7MZo08ZbLMFVV4ZbbBO07qZbicBY3aZc5+iVzUZdELvY+eZdFBkZc3ZdOZ508CZdSoRYBr
ZdWqBX/KZdcHGG0UZddsk3UaZderu1qBZd96P4k2ZeQNH1GzZejGklJZZekCHnA9ZeyYDFrIZe8uFHNEZfCxk3o7ZfIU84tJZfPi
RnjiZfQGBGcZZff8E2uvZfhi9Zl3Zf2XLpgUZgH2hGY7ZgRFQHJHZgby11bbZgomG1arZg2eCmsPZhI5B2+qZhKvGnkHZhNjgW5+
ZhUjN2s+ZhZY7k+wZhZ2qpRyZhdtXWi7Zhp6zGquZiJQTYaJZiRV9olcZicEj1RpZihpvU+AZirK+JJcZiw34YCNZix6yX0jZi/y
mk+sZjDuq4QyZjQj3ontZjT+ZF5cZjnnklwGZjx8Z4rsZkfuSWa+ZkvH65rYZlb1Gp17ZlgrY5Z1Zl2d4183ZmUVQZHyZmZdJlRD
Zmmr64F5ZmsdC2aeZmuOdm6LZnXSAnLGZnePlYEXZn8fT2UcZoZO35auZocMFHVLZosQuYPJZo4XZWvtZo+INmJiZo+cGX5GZpDl
g3yaZpHr/ldFZpU7FpBSZppXyG7PZp3ghHojZp7LtWe+Zp75m5EKZqL8kX7iZqY4ZJd8ZqmQgW5JZq7vnoqZZq8FyHXkZrKKbXvl
ZrLln1JxZrOYAHHnZrQ4XI6LZrRwy3FmZrk+d2EzZrpGelG7ZsRB+5R0ZsWCL21KZsedtn8xZsoXGI+pZszJJ2BgZszKpYV/ZtCv
DGFfZtGPq5qGZtG69GSLZtKAG1eLZtMR8FagZta5ZGtGZtccTJMYZttgl3NuZuCzAoMJZuHdIYkAZuLSBXo6Zu1COnGlZu7j7lnG
ZvHXL5XGZvPAqV7aZvWn6JHaZvhwnpMEZvqVoGfdZwH4/ICEZwYpGYssZwiwy35pZwmXIm0wZwn1sXUDZwtqOoi2ZxFVbmWmZxfk
xmMbZx1cSpEJZyOvlJpEZyPKWJ32ZyRuJmD7ZylwGWqxZyl5h47RZyy90Z4IZzCePluBZzGQVWXNZzMF3X2OZzMhdGUzZzPbPlWR
ZzShKnWGZz1/VlLWZz3TOn9NZz8MMYgLZz/D5olLZ0FAmoe+Z0Fnu37+Z0V6GZMRZ0Wf8InvZ0WutZI/Z0n9DGKBZ0rFW4jqZ0wP
yZwSZ1BwG5CsZ1VyU3WgZ1bPHFzKZ1eHplejZ1mSUlVEZ1qdVpMJZ1wYR2RWZ14qDZdwZ2ObpVHdZ2QgwlMRZ2WpUnqVZ2d8zFQn
Z2eK9Y74Z2tLOGW8Z3Ur7ohBZ3Vy21EzZ3YX+J67Z3ZDVVzWZ3cbF31hZ4gbkHPbZ4nFlWZTZ4qfTGl3Z4s4gYNhZ4zfc2jzZ5CN
dJtCZ5JxF1mxZ5LbV2nXZ5eb35POZ6L9e3w6Z6NzSnUPZ6jBmmTgZ6unKo7YZ7GlHZDwZ7NtuHysZ7QK2k5yZ7iqlnrDZ7tDSWvN
Z772SFwjZ8OuT1JMZ8tz91cKZ86f9XDrZ8+rJZ5UZ9AhUn1OZ9Pg4ovkZ9hgr2jmZ9l+WGsvZ9wAnVOQZ95LdX+GZ+AdgZ+ZZ+La
TF0kZ+UiB2ttZ+ZjTVCHZ+oc0lqpZ/MmqYuPZ/SbZ1J2Z/d7v1VLZ/iqDW5eZ/qv8naTaAFQXFMGaAGIk1fraAZNE4BQaAf53IMO
aBA+ZpNfaBCqZnlPaBS0xYhqaBiiYpebaBp1t3qtaB4bnoZyaB7Jh5X6aDKCDmduaDisN5zRaDplz3mUaDxN5XFYaD/5RIUcaEdI
3WbJaEn5VX2JaEvUMXLkaE/Tio9AaFD432MraFH/lVJlaFe1k07BaFfmUHTLaF0XN5CraF54/YgOaGftJ1A+aGquEXRTaHG8CokI
aHIDpoTiaHPiLJ4LaHiPuZaHaI7qCmJoaI+HeJnKaJXpbWT5aJiAXIafaJna+poZaJvIupKHaJ2proAuaKB4eXCFaKTf+Y39aKff
xJxjaKkO+ZBsaKuRaH9oaK4vvnYpaLIQJk+FaLR8vnIpaLarbYNraLclKoJbaLpA2oy/aLwWT4akaL6lVZA+aL8VS27+aL+TeV/0
aMs3gowMaMvMFWheaMzJgnPFaNXxrmghaNZWIoQ5aNk5snSLaNuBJYk8aN4D/1TCaOAKZ3ceaOBRwlvhaODHOo5FaOMyLm6xaOfF
SpsoaOqXApc5aOtTsI1paO4IpY2paPGhHV8LaPNpMGHmaPWRwmyyaPm7anqjaPrCoXebaQXZ4mayaQwP7mJFaRMOAJfFaRYTJGgc
aRY4imrtaRjkgG40aRlDen9maRnCHVwKaRrX+pMSaR/UZHqEaSDVIYN4aSgbGlfVaSsKnVLSaSw6Cl5MaS77UpImaS8FOmzQaTJl
H4pnaTRGGoetaTlW5289aT8LlZ+HaT+vw1yUaUCTNmdjaUIAmndKaUceyJZZaUeRhYypaUt87WjBaUyFRmlRaU64rFWzaU7tzJTi
aVROuH4/aVYpMGplaVa5012caVpIb1R+aVzEWHzNaV6TtnZwaV7jT3xHaWB7z3f5aWGolVgqaWRlKXtYaWcH8oSeaWreU3RwaWyk
x1w+aWy+yIKcaWzw6Z0FaW/KmmHXaXBQZXeQaXE5ImliaXSGyXs8aXY7p0/naX0WalMpaX9kFWWgaX+KPIp8aX/AiIFpaYQAhHr7
aYeAm1wsaZF1OlcoaZKoclmlaZaEMnJWaZkIHIqNaZ0ntVnFaaDV2pwXaaK+ElAUaarapnH+aa4zr5aqaa/dQGDXabBT9XQrabHj
qFi9abnCtHDJabxTTZ3Dab5Hq3CIab63WpG2acISR1TEac30pJ3padMH9HaHadVZLV7Madb6pW/QadnBk5OuadqK2Go7adskvFDo
adtV2ngzad+VSYCnaeHmgIhoaeWK2XGOaebKfpuLaei68YQlaerbHWxXaeylPWNSafMLiXNKafSCcGhOaf9654ySagDw+1Vragkc
PmLEagrdhnvyagyLnF4aag3LPVgnahmiRoqoah2icFUxaiJtcoHNaiMsGoiEaiS7QnsCaiVLhG2gaiZsXW6/aiiBFYOWairZrpU7
airhx4HaaixKopVjai7UTV+0ai8cCYFxajJ9nVOTajYjLHQjakKFb3PQakY5JFfZaknj8pwPak0yX53UalMy/V55alg46YdPallJ
QV+4alp2DpOJal9nX3LRamBiCYgsamM5Al7JamR5TFQyamTcNJe9amanipysamr85GJuam1Ugogbam4hM5NDam4rXVFKanJ865ld
anL2R5zHanVU0k7janYW0XewanlDgGyaanqu3F5qansxOobMan1iDn4SaoQavlxxaoYRXGaLaonNnJ6eaoqSYWIBaovPVFHAao5o
LGkwao82IG1Dao+xHYYvapAIW3+HapnB9GENapnP4HXNappvrnrsapxaYoLtap0CA5N4ap592Zq+ap68NHncaqUf2XsbaqffVmLf
aqfqY19Raqvvd3wTaqyfGptaaq2QgmmCarPjTlDZarV3a2rxardPYVPfarhA64KiarikkVPKarkon3Abar1xOGN4asZyJXbtaspp
z1zVasw1G393as2VMZ4patKquY+iatj4vVnpatoZc1LuauY+PpiKavDMUn9MavG05Yl7av00eVBsav/w1YJ3awhQZl78awmlwY4A
awqD+HP4awt8Wo/zaw3VuFwfaw4XJ42Jaw/qd50UaxQRCpJnaxQrs37OaxWSX3V0axuUU5d7ax1lCG87aya0TnOYaykcO2zSayks
7Xw3ayvrP3jyazCTv4ZEazD2hY3bazMSLW+aazXN33SsazzTNXgeazzqSIF9a0Hwe5pSa0KPZlama0R765ina0S+ClE+a0WDmIYZ
a0a7i1gHa0ecz2D2a0fY03yKa0pPkG9Pa0stmk5Ha019dpEsa0/2plTja1FXWn02a1w41WO6a2LGtX1ja2Nho5Q9a2bu5lb2a2cG
2lywa23HEF+6a3Rb6IIAa3Uwgo0Ba3vWVJtva4A8TFE/a4fa0VHba4qjuVSda4rUoXbaa42+1HdDa46abYPMa5BXH2hNa5JLQmLO
a5VO/3Rha5it3mVRa5rnnVo6a5w44Yfca57sCFU1a6HBAXPaa6HDJIWJa6VHdWOba6pBuITla7ERk1EBa7JTWmSea7vQ1Zmma770
CmsWa79DQ3ORa8PRP5NPa8WUAZfVa8aewF+Ra8mSzIMza9bxElpqa9gbyZg+a+VEop5fa/LEZFnsa/QkHWjga/i3h3Mqa/mJbmdO
bAW97oC/bAZTBZyWbBdOEWjMbCHVhnB5bCHYomm0bCHc1n1ibCNi2oqJbCo7FIMnbCv3BWNPbDcyP2dqbDj7954NbDw/D3nobEGn
1IrRbELMaHZlbEYVRo7kbE0jpk7HbE+Gmo2ObFHBbWJMbFvE6V+ZbFyICHeebGFq5FD0bGMSg1AobGMdjli8bGPrN5NcbGnEgHXd
bGra6GzibGsw7X3lbGyweJBjbG4A+3pSbHGBbHZ6bHL2U5CSbHOR7Yx6bHhjwJYgbHmyPW1SbHpEU2fubHt78pa5bHzxzWkPbIBU
vW9ebIeY9JF3bIjVxoREbIrzp5mjbI0IgpGfbJMoOWkdbJP8O3GnbJRaEmPhbJYwGY3cbJiVOo2rbJoMkInhbJq+/k7ZbJtZn2Pp
bJtm8WIabJ2fkXOrbKUKR1CKbK21e4CMbLAGfJRcbLEssFicbLMSP5J3bLcrEHDibLdYE4NIbLtc3VJ0bLx42m8XbL62rXxtbL+g
21yTbMJQZ4O3bMM4fpIHbMlMWZ2cbNHW3lllbNb1qpUhbNgBrlqtbNiqFmFIbNnd1FMbbOGvcFbibOV1Mm1YbO2O9XFibPKuzW25
bPSUDXVfbPZwEJkTbPdFTn9UbPnLnGG1bP5r3oYsbP6tpHlgbP7rDZknbQCOj2INbQKXPIIwbQaaXWIfbQ2HiYp0bRMQ+12MbRcy
ZWKYbRm/wJc/bRom7ZI8bRqLvF1ubR2p/YPWbSEeKpkHbSYFQZTzbSjoPYJDbSy5q1bmbTXGs5s5bTZ96HgpbTcA7mvibTdw4pE2
bTsQElGLbT45yoI1bUBNCXwebUD2rHfTbUu7MZfpbVITApRZbVQQEpivbVodznkrbVska1vebVyQc2bwbV7QFX67bV/2x4wibWRd
t42kbWWyp3tGbWvVwoEWbXBTso0RbXORsp76bXWtsW+NbXwX42VtbX1a5X2WbYE3tnf/bYZic04jbYnNxVIpbYzhe5FxbY3qhoCw
bY5LKpcHbZAZnJVDbZBtPJfgbZM3oHMVbZawEX/DbZp7JoJwbZv6BZdXbaHuXGmvbaIgBJuCbaK5J4bebaYsNX/nbaiAzHHcbawL
NFOwbbQKd3v4bbhXypwQbbxORG5Ebb/9OHQPbcQvf3Onbcs8encwbczi4VYXbc0bD3Y/bdAXQoU7bdpdcF7LbduyX2qBbd4hapdu
beJH3JYvbeLZA5FibeQFv3ICbeRLkIf7benPm5R9bewnlVRnbe3oM1Z7be5jeXK6bfF2OXc/bfLb323PbfPPYpRVbgDmLGXpbgJV
7YUvbgWHvHbXbgcQWIbUbgw4FH7rbhBlVnrabhHqcU8lbhQUV54dbhdm0l3Wbh8dhVaMbiHTI5k9biKyoHesbiit3Jb7bimLv2CQ
bisWD0+fbi8I9V9BbjU8EHI1bjdTD1RybjgWUplXbjkXcGxqbjrl45cUbjvp8nrSbjwlM3mdbj5UM1n4bkRZuHWnbkq2s1QPbky9
rpNJbk2Fh3Sxbk8ytotkblPzAllvblR0ElHPblZn9oO8bla1P39lblmqsYwQblnNZYUzbluH8V65bluaJWtDbl4qf4C5bmALfG/p
bmSHqJ2EbmhemWfLbm0ovVf7bm6AipRbbm6zxIcmbnDGkIj4bnOCVXeNbngw/Gm3bn5qWlclboMSSWRCboZtV2J2bozHKlblbpIk
cpwnbpPAT3mZbpqc8XnXbp+jIngJbqJ15H30bqW5jFXCbqYQZGCtbqfBzpaGbqwwhl4lbrImZXujbrOXJHlnbrQuVoqLbrebq4Ck
briHLpqobsFEMHT4bsF0BnEKbsKm+V6SbsL8ppribsOHpm6cbsYDb5zpbsiaBp3TbtF0c1cAbtHqC5L5btMIKIZhbtvg/53HbuK9
5nCPbuMNGJjtbuMszomzbuhOdp5/bunNE37qbus52mM1buw2PlGKbuyA1Ftcbvcoj1tBbvn1pGMpbvo66WsFbv2Jp1T1bv+umWn2
bwOphJ2jbwczlmmKbw2MeJJ9bxbmdlQlbxn39YOSbxo/12G8bxqpL2kLbx5R3oGHbx/DxGItbyaR/HmBbyh4MJ8Kbynd9IF2byx8
CYNQbzAhImOYbzKLoY2RbzSyy38LbzTIKHltbz36aoPob0BxqWQcb0Wecoktb0eTUWMRb0zzcHTpb09bVoNEb1DrY3vcb1ENT2Da
b1I/Im9zb1Lqn5WAb1aQn3b1b1bw2Hw+b2ANz3Kgb2qNuYxxb3bLIIjab3gd1Gitb3pOmGQJb4Gf/3e9b4XxxGDFb4wh8Znkb4yx
PosMb43FYpIDb5TKRpuKb5gS34jhb5h1O1e2b5tqnYAEb5zn1E6Rb5z6wXL2b6SWMpxBb6jm/YmIb6qDF55sb6zqTmzGb6+D5nld
b7Fewk4Mb7HGFE4ib7nzgIGhb7vrIZYQb7yLK2wSb8CjhWzWb8URzIoIb81YBoD1b87/J2Y6b9NB71peb9bO2H9vb9uCM1o5b+It
p5kLb+Yyk1DBb+jmZm0nb+5wpH6Nb/KBxmxlb/OnBJYyb/ccsVFab/fanV94b/5FwYb3b/+xQ5NrcAkEmZGxcA8xD5bscA/ZWlds
cBbK24f+cBbiyWt+cBczoZXYcBhX1nFRcBoFW4bDcCImkm5ScCXNMlBicCjn1XBGcCqbn4/dcDFMeF1xcDbcH5GAcDe8xlu0cDin
eZeCcDsExF+YcD97I1mEcEOOOJt9cESdsZg8cEXjGF0FcEccmGkucEkL03XPcEmIs5xOcEmsRlygcEtkfGcEcE2eMIaUcFQ6UVQU
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
dd3L3Yywdd3iTU5Edd6w442VdeEhwXwndeK1OHbBdeSp72rOdek555hFdemD/HkUdfK+n1qedfilOp+BdfvjzlYVdf+zUlZbdgXc
1GMYdhEXUU7idhIAiHnTdhPW+l+ndhPsi2medhR/uZbTdhpvDVrFdhxlEH/wdiSR0pCldioGFE8hdi7Gl5oxdi7REoeAdjLb2U6H
djTiZXAydja92lWpdjbu+G5ndjnH4pBedjwhF2A5dkXu5VSXdkwyunxgdlr2KIUWdmG6CIT2dmIDuXlydmPwBHdjdmQUtVyKdmYL
rG3cdmfwLmhddmuxPY9DdmwRKmIJdmxE+0+1dm8nHlMOdnk9apbldn1FP5f1dn42+29/doWiu1MNdoZJQk7MdomtaWtOdonpEZpY
doryFnpxdotEvIN1dox+eFtodpA9X5jvdpSjIndXdpWO7pyFdpZGLXsBdpf115Rxdpq9OFaXdpwBb1m5dp1vPoD+dqkc/Za1dqk1
0Y5ddqzhcWltdq6/0Jnhdq9nnptsdrWd7oEPdrjLjVrydrmbrWEDdrsQ0Je8dr6i2Iomdr91ZG0YdstVQ145dsumNZA/dswcnYJZ
dtA6i5KodtTpfmlCdtUlA4gadtq/4pvidtt451VKdt5GbGTYdt82h4kidt9rY5SDdt/u0mUeduGUIFpQduHMyJttdubmYHWXduvX
IHmhdu01d4gPdvDSKYVZdvW+z30KdvY9uF5ZdvepS4d3dvfpkZzrdvm43mOBdv+X2Y/LdwnmLlizdwxTipu7dxIrAE62dxJNi51X
dxrfQlTUdyLtOWdrdy1+ZGKSdzRIzV17d0NR6l8Ad0OvnVLod1Gq1Vlad1Lq6lEod1MFCoEYd1PmM5x6d1pEa38Id1zgJl0vd18W
6IMUd198B4AAd2KSpZBYd2OSC1dVd2XcpIKMd2i/w4dAd2lP52rld2uJ0JECd2u+hWHTd2xN32D+d2ymRZLwd3P6YnBld3VRLVza
d3Z4009qd3bJTX0kd3q445Pid4Ef2Fnud4a3RlkNd4j0sGfOd5Rze5vZd5bJeHxPd5ggxYTgd5i4l5zcd5qovHZNd5xT43Q9d5yu
aGPsd54tZloWd6aKu2Idd6l9q4Kqd6p7j4OId7HGVYXJd7L7qk7Ld7OrtoaRd7idpIL+d7vaNY5Id7wklX1nd79Qznjkd8DKTGtB
d8JpQVXTd8SGNX08d8XGQ5kCd8afZI5wd8md+o8qd8vvhm6Ud8yrAHtWd9g732jdd9mLLo17d9vJWZ3od+Cnl2yOd+VvMmKrd+na
U3i7d++CNU8jd/iNc5LSd/+2k1CCeADOOJ+ReALiIYoUeAXU4ngjeAaWI4VzeActclT6eAkQPVdLeBGSt5PLeBU9vo4qeBo9v2Ac
eBuTr50seB2G3V1UeB39OJoReB6RvJozeCGdXnK4eCMHd2+AeCjF+GdReC6g5JfAeC7Oq3NPeDFAw2QreDGN1JeieDJrgomyeDYt
IU61eDZADm1seDd3T3o1eD43k37keEPRS5EEeElii2Z4eFDcN5GjeFKeSF1QeFRuY38jeFk40GekeFwF+HjceF2WOnu/eGU5wXkN
eGfDkFireGqjNWZ8eG1o52SAeHNOK2zTeHQiYFWDeHSAXlFZeHb7coFSeHv/ql8BeH3mkm8xeIAoBGMceIDOKFM+eIMZNZzieIzP
Gl3KeI0TmF/TeI9z3HwreI+Jq2U/eJPL9m6TeJZrBJTneJq7plGseJzIo2JVeKbWh2cpeKqxXWoheKx4m1jYeK0iTonLeK3yNFTV
eK4HMlH7eLFXCYA/eLHKo0/YeLlCI4YTeLtLeZ0xeLxDCVQoeMAb4E6LeMA6d2c+eMgb/H1JeMxe64NWeMyRqlVQeNQskpHWeNZw
mIn1eNgxhJwTeNoVhmgAeNunYnGqeNxSB3+yeN5C65l7eN+vnlzneOfEjF60eOjdTWjGeOrrsHeYeOwwOFbteO+GsGleePC0MYE5
ePDVdm9TePdbLlbcePgnd2lheP562lwTeP+SrZUSeQS0T5s9eQbock6ZeQiGbIcteQnH2G2qeQvxuVdPeRBWj2VpeRBdl4S6eREo
VoASeRPK4oVDeRUJdJoLeRbpaYQDeRtNRGXLeSEtJJ2leSPv7GbgeScVQGgSeSntaYIzeSoaoIiCeSyjVHkkeTPKoG/BeTW+bFOR
eTYWdZdHeT5D3ZSkeUE5H2eReURzJU7KeUd15ppoeUexzXL1eUn8QojfeUu3sJFweU7yyokUeU+DoorBeVBG636SeVO3iHhteVO6
80/JeVPU8HlGeVYBqVvceVasE4/yeVcmBIiJeVjH1oKFeVkW6nZdeVmN2F4zeVw8rHofeV4opmYxeWDJ6pqSeWFaY3yWeWJj/VxX
eWMkKpXteWMkL1YQeWogrmdxeWr5E48YeWytaXCweXDcwleweXRHF1IXeXt1+ZxPeXuZ+52AeYAl9YQXeYIq0mgJeYc1hZMaeYmI
Ym10eY1HeWJLeY6m2ZGDeZKewH8OeZfMl1d6eZo+pnuUeZp8OnIxeZr6ZYxbeZxozZpeeaR8r5M/eazypZ5qea3WWXLVea6/Nnsq
ebSHNnZsebuVx2gvecLFSZQUecLZQGsnecfqzJXeecqY+pzGec6EKJnWec9NgFWVedJyLZZuedZKIV99edgdgn5Med/sFoSMeeu+
cHM5ee9RBFQgefWTfnGcefcWwG9Yef6QdJucef8/A2g/ef/xzZcwegTC7W3zegZMSJnyeglHpYOKeg2l0k7ceg7irptoehLLsnRN
ehbmx16EehgdvY/ieh9+m1DHeiEdwFleeiMYjWi1eiYMN2UIeiyVmZpLeje03J0rejlrQYUkejvR2YYFej8Jr2wjekJKyooZekw0
dWG4ek9wTXfxelYHInzMelYd9psnelZgW4i3elc/CYbGelf0rphHeltCf5D6eludM2Lbel2PCYoAel+vq5FgemDj8nP8emYqwp0e
emnnjVy/emu6MVIHenCS2JN/enPpUpgTenXlCo9penou2YkOen63KJYwen+XnYO/eoVoxZJ7eos2ZFAOeoxEaIRSeox6loV8epNz
On3HepbXTWOXepgoGWEvepyo0ng4ep7GOH9wep99kZcrep/jK0+0eqEzalU3eqMaK20reqXeepzZeqgfMoBWeqmcXViIeqyDP11m
eq1Wj2+teq/cNIk5erJSoXXgerJq33anerMqRn6AerY0goTDer017pC4esOeGZhLesecvVU4esvzwme1etFsQp6qetGtZ2kZetb+
XHiSeteMa2BBetmeqV+DetxINZ3+et0UxGT7euK8eFJzeunbB3cZeuvDd5RRevC9PY24evV6PoM4evlEBJbRevqijodsevxZ1pFz
ev8AKJxRewEkr5XnewHb2GrCewJLsJV+ewNzpZocewuduFStewyVjZMjexWdFXhpexi8Elj+ex/7Hm1deyN0hVlxeyQnLmJAeynA
xk5Cey2ZTI0Mey5lQW3Cey50A3pUey9F1W1ce0Ho/H2be0PGj4lxe0WSxZiue0Xajldje0d04WnAe0uL2JBNe1Eb1HkIe1JyNnyp
e1VLi1UEe1Zh+4xte1i8bYCze1qRVI6De1sS3Z1Se17YNHN9e2ZQjGL8e2cGCXx1e2g842wYe2g9YGoXe2tURlBAe3EdgYvGe3FP
eFMqe3PIr3fke3VgpH/he3rr8F/Pe3tdU4p/e3vUymbke3wLEobke3yXYZJde4WBeVUYe4+1Z5YIe5PKS2E5e5pYP5RQe5w3PJbP
e50Krpr9e6GXnF7ve6JT5541e6TWdpwie6kFMGLje6pSfIRfe6q583cKe6sHN3TNe7DuFHEve7O3m1rGe7Uzo5u+e7q0XJC9e8X1
XnXae8+0/XEze9ZMZJMse9gb61cBe9jE3nAGe+Pz2FP9e+Wgl52he+WoPn1Pe+/Axldpe/Cc55Foe/JdhJYte/O1B5rAe/T9fnk/
fAFGPmK0fAJ84Xa0fALha1P3fAOls3y/fAWgvn7QfAa5R4p6fAo+t07WfA9I93wpfBCve37MfBFmrFCIfBGDfHdJfBO8fHMwfBfx
fZxefBlFP2+vfBnXvV9efBsGw55KfBtBKV+JfB+i7J2CfCbMJGNHfCnBDYj7fCorGoRRfCqyK06pfCsbjWBWfC3dQGLAfDgS+YDd
fDgmNpppfDgo6WeTfDhZ9ogwfDjMcGUsfDk/P5KkfDnHlnpLfD1fbFNffD6b8GSjfD8VIITvfD9Mv5LyfEei+piwfExU/lQJfE63
UHkgfFEJuHwOfFMHuVHFfFUFIlaYfFn7b4ZlfFqdPpjXfG1Na4zZfG3fIXWwfHSCao3LfHmqSlrwfHogK1CUfHviGGAGfHyVs2+S
fH0puYaXfIBMwHFNfICp3mgMfIPLj0/7fIXwPnXbfIpYUZWafIr3mVEYfIttKoE9fJCYx48ufJCckp1tfJPtQJvOfJYDBG70fJgw
PJWefJkKN3ILfKM+9oXkfKN5dlbzfKd+5l8WfKfLmlqbfKo3BGuEfLNJJ5pzfLQH83ZHfLaa3YVMfLlTY1JbfMXj/VRFfNBpVowO
fNfMHYQ9fNtnLIHnfN74u1yNfOFKopn7fOPfzol8fOanPVs5fObgwmw1fOoff3otfOrcdpVHfO5J5VHxfPJws2mcfPirSpBZfPxm
U5DjfP+TyXAtfQpjVWK9fQpy2FpJfQyDTYNYfRM5nJx1fRbWimGhfR8/lGArfSODJ3T/fSkEMU86fSlzzHTbfSx8sE7SfS1Phnrd
fTCFboYVfTF7ZJm3fTRCs395fT70+1F4fUb/t2M0fUguFHuJfUzIQ5gBfU4Ud09NfVAjOVnifVGWxmOafVL6ApzOfVd2hlI4fVh9
6ml6fVqLu5k5fV8zyl1jfV9G+Y1AfV/DdoBTfV/Y3lSWfWHNAI7pfWSYIIEsfWS9z3GsfXOkp2wDfXeROWCmfXifBWvufXisfZsb
fXsprnurfXt/s2UBfXu8G16pfYo1X2jOfY7MoWdsfZE55FIkfZW5JnYrfZcA8HNtfZiA+39WfZx+4ZPwfZ4NjGLnfZ8C9ZgdfaQK
0mMifaS5CICpfae3kFgyfa0ZDo8vfa2HBJtUfa4pAH6Kfa7AX12yfa+Sc2CBfbiLo583fbsL3HLjfb2ZooZFfcWRuGg0fctZ85Lb
fcwNx35Tfc1UTGpkfc/fSlJwfd1BUV/LfeKoqXdvffO1OoElffSJbnlKffTB2pdAffktsF6iffoaf5YsffrrMmb2ffuctGxhff7L
amzqfgLDmXeXfgeotn4Hfggbo1njfg1wf1wRfhTRGV9DfiCHhlorfiOQKlZ/fib8+3JBfirbzH4Wfi6E2FkbfjCrDHTPfjIi8IOA
fjRc11N+fjjpJn8/fj15N2J8fj7nAGS4fkLidoWzfkR3zoNjfkTYM13TfkoiyYR6fkxfDlYlfkz8Y1Dpfk3kbp3nfk8moF28flIf
h37WflQVGGCuflgqkWH9fmLdDlMhfmMc3Ikdfmz84YJIfm4Y8Jo9fnBL8YNwfnJZJ5KFfnN1GVr6fnSRjlGqfnZScZW1fnbSRmNu
fnk5LlGxfnlp31WPfozcqJNvfpw82YuRfqAka1OgfqE8rJBffq1ENJbpfrGkjIk6frRSoYZvfr5MUYj1fr+d6pgRfsYzdnF7fshn
G1v1fspAn4yufs4Bx2BAfs605XeUftDTJX0rftOWjVCOftXH5myZftg262Tnft7znZhXft+JC2VAfuB0/5a0futQoXG6fu2iOZMO
fu6RimfWfvEvGW+nfvPa9YsGfvd0NFwwfviNgFhSfv1HrlIrfv22ZFnVfwHE1VbMfwISAJDnfwYjf4nzfwkiVmVOfwmOqlEqfwzu
f3A+fw96vZLTfxMQTH3IfxTwgZs8fxd0DoLefxkTlnC+fx+BwpCDfyBPvYp9fyFKLFlffyNgMVcgfymff2GWfyyfhY1MfyzguY6Y
fzGEOVpEfznGiVc/fzuZ7Yi/fz0kaFfyf0VKi5WNf0ecNGLPf0i8BHynf0lAgGaZf0vEfZ9Kf1HKeVGMf1K134oCf1o6OX5zf107
dFUAf2ctyGOif2tu3JTsf2wKCn88f3BKJHZmf3Ta3VO1f3T+GV+yf3Xz+nBbf3hhcFmif4AhgmSaf4pQyILnf4qvl1akf4rmiVWa
f4uRDp7Xf47JhHw2f5A+xU6zf5hXxoNRf5jqr1Tof5mAb3KVf6f26F5zf6im+GaHf6oHTJa6f65Qq3QKf66ByI9cf7ZA4npHf79x
CFC4f8bLo1obf8fP+1vFf8gccFlqf80hUm75f82FQIXKf823i1mGf9ArCZvef9Z85XNJf9h3eWKMf9pyZ2Ccf9uyQYvYf9xTklIU
f+PY6V86f+o1bHlSf+pKP1vNf+5eiJj1f+8d5ZYzf/FqemlIf/NeMXdnf/Q6GY55f/SuSoBGf/W2CGTLf/YfWmFsf/rhRXzmf/04
kFBff/3TgnhQgAAFOVk+gASC83CKgAr9j24MgA2eyZBxgA7B6Vf+gBNtj1OSgBnkVYTQgBpch2XtgBy+lVR4gCjGBVjlgDTruIzI
gDddEljKgD5gjX9OgD53r3XSgD6ZS04ygEK1f04egELaEVwrgEd6jZRwgE1/M2stgFCSxnMlgFIjvoBYgFKrT2p2gFQ3e5NkgFTG
3pMggFWg4XxngFi6B5lcgFlwW2QTgFrAt0+HgGBCSWGOgGYOenh8gG08ylL1gG+H4XmPgHCx2WUNgHNSy5isgHtuEZKKgHx/qlRc
gH1lFI1UgH6oVZ0/gIDF+FfFgIKLTFzMgIpBolHwgJLqsXfpgJXeFpsjgKGSwm+dgK++HI3AgLxFgHMMgLyUnGwOgMGlh5VvgMH3
hImMgMtB+5RWgNGB64+9gNGSV2C0gNG3I3pigNI+YoNbgNKNvn+DgNTm+pRDgNUQymEJgNxDm4R0gN22k1wagOAxlZI0gOB41IP8
gOFGa4YmgOFlqI3PgOL28YH3gOQBunLcgOSO0HWQgOZNaXfngOZdy4K0gOgTEpAxgPHVbWipgPgJDW7ZgPiCbpl/gPvqXl1zgP31
b5UYgQMgbpW5gQSd8Fb8gQTrRk5YgQc8uX3CgRtsBW5CgSQ+e4ACgSpRxGIDgSrSm4h+gS106mWBgTK+lm/ZgTi8WGgagTvMiFQE
gUHrn3xygUYQt4dpgUgCD4KOgUnMdV82gVEdRlZYgVG1VmDOgVQgLYzqgVg8dFGwgVhvqXYwgViHmWN5gVpVFYIOgVwtRmhHgWDx
THxAgWV40Va1gWjMzk+kgWjuzVGIgWn74mkegWpr61ivgWwoepTvgW+UDHSzgXEylW5cgXaC7F4RgXlfR5DHgXzdMonrgX8EXV4y
gYBjaJwsgYHkL54zgYPnvJFMgYZzLW9pgYuCuXqXgZHnm1/UgZNrjHkBgZdb4ZDXgZmq3Zx4gZn1YFLPgZoJTof0gaKMm2YAgaRM
Hp+igaaH2Y2Lgacwo3wBga7BEVbHgbau1ZORgbyAynTngcLI9pktgcm5XncfgdqNo5r2gd4ZM3Z5gd4p9naxgd8l3lY9geXBxljq
ged2T4yFge9p22ySgfEE+nllgfcHF2i9gfeRtniIgfj37HcEgf65W11cggL+qIWKggNZKpzqggPMCpITggSeJ1V3ggWfPGtmgghi
cmbmggk1jYWOggnsFnRFghKGJJzWghUT3Wtngh2y9o/rgh7YPofUgiEqDY3EgiMcs3i4giWJuHt3gibdEHPdgiqTLHW/giwi4mEI
gjFonU6sgjT3WlDegjYB5W9rgjYZOI50gjsgaFgWgj+rNZiTgkJTp5WzgkcUTVcfgkwECW2Tglkn7mx1glnfJHZqgl5lNlSlgmIp
zJPagmJB6E7QgmJYk1iNgmQ2XnhGgmyc132MgnFT3Z1CgnRg5omFgnTc73QZgnVMKJ4Xgn9y05UBgoHj0JqJgoTpT2Azgoz6y5HJ
go5Z3FKmgpBWpoeggpStXJoCgpUyi3J0gpmfTIlJgpqefZDKgpualmvLgp5BRlYmgp/zao9Qgqe8dF6kgqxXLk8Hgq1IJXVQgrP6
RZmPgrVKjYY3grgd5n2RgrhC2536grshymi+gr4oUZZXgr/BvGQGgsQx+J53gsWZ9X/0gsf4Al3vgslA738Fgsnrg1LkgtGcJpHh
gtty1HhlgtxhVIdwgt35BHoJgt5qq24LguJvvYHTgubMmXs6gugWI40qgupkaXhhgvHfVYybgvLU93CngvOe61x4gvSBTVI/gvVG
XFmAgvVzdGbrgvaKCpDLgv8gMXSIgwwjoptYgxZqVZf/gxwuM5ddgyW54VTKgygStXoOgyj5qXHTgylDD4U0gymb7ZS6gypAGFGR
gyz0qYtngy/9d3eZgzFbFF/kgzUq2J33gzrH524Hgz6h5os2gz7qAlBpg0gIS1Dig0mxTm5Gg0oh73KPg0pft3yBg04EqJWUg1Oy
aVJIg1kS54Vqg1wvc3HPg13WqYWNg2RqWW8pg2kz0WC5g2laKWbng2vgCpnjg22S6WXng3FJDYLwg3KP9Vyzg3QJ7HFwg3q2aJcf
g34YpWlNg4XhF4bZg4X5kpzsg41DdE6jg48s+JRXg5ImpF9Cg5UDyHNcg5hs75szg5pkhXS+g56jnHAYg5/rsoMLg6YbLol4g6gv
gIPNg6jAFGGCg6kxIpXIg6qVP152g6w/h3iag68VbHYSg7Xv4o/Cg7mHeHnIg8Cj0W+zg8MmL3cUg8UcDIf8g8dxoH+rg8yaOlta
g9Lt8IK7g9NqRHpJg9UZAnBwg9VTQV/zg9bfeGOtg9hJi46Og+kjlHmyg+nUIWpRg+5xkI0Lg/0664oMhAGTUpcMhAIT+m/IhAQ8
m2AAhAYz0F2fhAyiRIjohAzAeJuqhA/GL5AJhBP8oHsjhBSnTWh9hBS44pTNhBVapHhEhBZWmoKChCJL+WU0hCO4JVB0hCYSvGAf
hCYiM3HYhCwVr452hDNtNoYEhDOPh1/rhDTKV40DhDbGpF3FhDhWrnT2hEEpxYbbhEMke37zhENG4mhKhEbfu40uhEr7oVYZhEww
FWjihE3AoJNEhFAPZZGHhFGJiY5ghFb+YpwvhFpXEZ6mhFxJz3TohF8/kFw2hGHZanFJhGcaXnEPhGc6lGgnhHPd7l/whHPiSJkQ
hHiIDFNDhICEdJJwhIFqc28ghIUGj5uihIZxi1AyhIgRO1cqhIoAzZAvhIuRjl2OhI1a0pvMhI4BQn2mhI8oc1mRhJC9bFA8hJOV
pXtVhJeLqGW2hJgr01g7hKCm31fchKjdApychK6EeW2lhLFdC5q4hLaSU1ImhLj2v1KshLj+8WbLhLyjpGz9hL9YJVMwhMCua5N6
hMYdhmNshMbJM4vIhNTk6FHehNaVc3M1hNd0WGxNhN15DnrChN55pndwhOQkZ3WzhOXHUE6qhOkF5WV0hOx58oj6hO03eJQohPH7
b1mPhPWMA5e/hP+M/3U0hQP1MFO/hQljFG8lhRYPqIeUhR57Cpz8hSLq0GF4hSRxDpYchSXg45WwhSk9IGPihSqfpnOAhSwlb4rW
hSycL2+PhS6aMo0EhTPUdGy1hThP4J15hT1alZpMhT5hvFBQhT60hG4XhUCxwotchUFX4JenhUfk4oYlhU7ZeX9shVUGNXjvhVb0
gYq4hV5LdXRVhWFCHoUxhWHKH5NIhWIuRoxPhWKMbVSUhWLWnptDhWNdu3ONhWgeLIUfhWpMhE4bhXDEzXxbhXIKmIKrhXMkTXcS
hXNf04OGhXP+8WDRhXhNo5o6hXmrSm7VhYPzzn4ohYth5WIvhY2lSHCOhY4yWk9thZNwiFR6hZz6l2O3havy55xDhbNClogfhbap
fp2whcN4iZW6hcPJs3bghc1z+ovRhdOT7mW5hduLWl4nhd1bT1Tshd2D82WFhd3xSH9XheGkHVqVheHBwmpqheKTUoWCheTUlVeU
heeCT20lheimRVTXhei3tmBThfQEHFSihffmNG8AhfkZZ4j3hfpcp2Nihf37q4GMhgEyIm6BhgO2VXPRhgR8v3qshgWEe4Skhgcy
sWhShgj9x5efhgoIRk5ohgw2G04chg/YQ3BqhhKVX1kchhO3vIgFhhVjHFLvhhjlG18Yhhp1PIvchhzXGJ6ohh4+W13GhiJvOYrn
hiMNv5L8hiQVEJvYhihAO2kyhilSvp3vhi7Lj4Jshi+7+HV7hjA1CXCLhjL6GHEyhjtScI+XhjthF5IthjxNNF3khj1FP5IBhj9t
PVH6hkDg9FUWhkd6pYHQhknMwV1khkv1zXnHhlFLaHmlhlRen1/FhlpTdE9whmVOH512hmZZ/JIrhmbsXV1phmtvuGSohnGb+3X1
hnIjlJGmhnXhd1hChoSS14Ykhobuu1VshocCDXprhop0nWfHho628Z8QhpbXjZiehpowZYx2hqAy3Z+khqKmsZ1ZhqPIulQVhqSQ
0WLdhqTdSHdihqe3r5+OhrMFbl7/hrQQRXP/hrVF3of3hrqgLHRshr78DYTthsRc8mkIhsdlal/vhssAxXXzhs+B85vfhtF4VIjG
hthXK0/xhtiI+JaWht1MWYqQhuKdYIIGhuRGHVEEhuasCVmohvZv5GYnhv1XxY3fhv25eWljhv78uleKhwAIs5NghwBdO3E3hwcK
Um02hwiRWFPchwkcNF9KhxR0qWvehxsn8pX+hxwyslgChyCVe3SmhyGhQU8mhyMCTGCehybtam3qhyun+VixhzC51nJwhzkFKo0S
hzsz7nYgh0SSD1x7h0fOnHQeh0yy0YGth1CsQZpdh1Idy56Nh1Jc3WkFh1PoD5RCh1Su63T9h1mx1ZTeh18cgFNKh184629yh193
Ok8Ph2Th9FRah2gbkFaJh3DclXUqh3MAO40mh3UCUGTph3Xl/Gj2h3X5i2aMh3bSv4vFh3rx8ok0h3uSClJdh340LGxsh4FwlIRe
h4GbtYmeh4gpgoKmh4nuI2ECh5J+wY06h5Oh7lYfh5U29mGxh5ZLTG7Mh5txMYNLh51ukX5Ph6F7Gozyh6S5e3sIh6Txv40Wh6US
Y3/th6YqE4TUh6ZcBYgMh6eEp3XZh6uxsHrgh64yjGBPh7FGGFP7h7KE22dHh7TOAJVyh7YMZXNQh7bcY3seh8Qr3YHCh8uHfIA8
h9AxEFO6h9QU7V7gh9XfZ1AYh+Ij34d2h+RA6ogYh+RSp1DJh+SkIJJ2h+gVe12rh+7yj5d0h/Cs34xjh/TNZ4Loh/XGeFmmh/Y1
HHk+h/hq6pZ8h/mZ5Hh0h/motZKZiApxLJBriBIcaJISiB2cbn1xiCl4qV0EiC2psWoHiDJ6c5R2iDLnQ4b0iDULUW5kiDU1A3HR
iDoTgWX5iD05HmOwiD94w3fQiEQxmpTxiEm9tlTyiFBG/ZP3iFCs307XiFLBLWmziFhsZYfviF+XymxriF+kVF+9iGRhYpqEiGSA
+mdliGkZQVybiGlCkmMAiGnFl3jJiGvwyYuyiG+pOVMoiHD7XXmViHKi0IxdiHaUy3qKiHgVx5xriHjgmmj4iHlb21Y2iHzPP5rj
iH2mRIawiH9VNVkSiINhWnUJiIWi2l1CiIYOkXZ2iIglQ2fMiIlj/V7jiI41k4HBiI6jF2F9iJb3C547iJlvJoBAiKFlenyAiKUC
aIATiKXRqVkkiKl7HWGNiK25IX6hiK79IGKOiK8O9mbRiLTeQIRYiLT3lW6ZiLdkKItjiLgYFYj/iLyKH1VOiMc1I3fziMxBa2Mz
iM8jwX6RiNbaU4c5iNd1h403iNe1nZ4JiNhzPJPciN/bQ1nBiOAZYWAIiODWOViliOskzouIiOt6+GFiiO3rqm4liPS2RYNtiPfB
C4hRiPt6DWzbiPyWNokfiP28AFxciP6a6oBliQhk1JvsiQk2CmogiQqGypvCiQtBmJmoiQ/jlZ9JiQ/k/FmDiRhO0pwciRm9LlHz
iRsr8XooiSg7G2rmiSr2SJ2WiStNnnuXiStjnpUfiTGZi1DLiTM/MZAIiTviA07/iTvyj35IiTz4NlAHiUK1nX0ciUX8P2+jiUYH
X2oEiUuf0H5RiUz3G42yiU5Ud3ZziU9K34dHiVUN2o/piVWB+myBiVWbaVoaiVrfN4cliV4C4YuGiWMd8YjgiWchVZ+DiWh0T23o
iWoEaoGEiWx7t5jfiWy2G06CiW7bVIstiXIQtIMTiXQV8mkgiXRAiZ6hiXUFE3nMiXhMsHLoiXqocl0DiXy0I4zCiYB6a3lriYDq
1Wj9iZgcKXRciZjuJZetiZlnpVNSiZ/DOZ2oiaZ0sFSjiaey05YFia0cm1zpia4YI2ytibIqyWMuibJvZ04KibOX+1z8ibVdfoXP
ibeQjZ2Tibl2LYaxibpirHjAicHx0IWVicKVHJMficYIyZplidjVWHHridrGz3zaidsWFH0oidunZJVkid8fbFsdiec/qH4AiegG
eJrkie0WsHveifSPJpNzifWQkZSlifW7wFlyif2lPFF3if5oVXftif8IelvWigPRn4NNigXE3Fx0igyjq5vhig1pXZVCig5i5Ypx
ig99MFlmihGDRn3sihORYGS/ihzyfFJuih1tp4J/iiR9BmXMiiZ9hFnIijBGVJOcijIT05cdijmN2VWCijtVPE9miju8zHm3ij32
XpPSij7Qrpkgij/+wJk1ikA6wZOLikB3GnJoikNpXHUgikkdzoEtik3xlGmjik7PM454ik9/b0/8ilIHuGWcilQ0UIWjilrG65PG
il8PIYb9imZxB1dhimg+Foruimhh8lTNimv9SY9bim17/n4lim30FXvGinpGfngvioIb9XB9ioclqlrHiof25YIaiohpWXikiopr
JZxGiovkLlemipSMeJGripXVHI6aipbBLptOipxiy1eiipyHNGuDiqto+Fbwiq3HeXfiiq6H85i2ireSOF5Jirpn0pLcirqJG1uF
ir0JtY5jisCq63CNisRF5WZOissBz3T5is+TDYwPitH5iZNwitMqipaPitOGq383iuL+enJdiueNH4FYiumpvmPDiuoqH2xkiu0e
PYPwivPa8Xo+iwnXmZqCiwwoUmwLiwyRlYVGiw0YnXimixpj9mWRixyZ7F02iyh3YVp2iy1k/oXRiy+naF4NizaHoZVJizo1SZrt
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
kkZuaGJRkkyI3ZHnkk7L+k+5klbeWoW+kldWYHB4klen6pkWkljAHFbrklkAZGOyklqFC4ENklqyFp4MklyLq3uNkmBcTpQdkmcU
plIokm3YpmbNknBogluvknm6T4M2knq0A3Fskn07elAMkoczjXrzkoijx37DkouJr5Rokou3AWZekovXdnDYko0VYmfcko6fFGtc
kpATP3ZykpApdoGekpRk/58hkpTw0FwokqT3NIN5kqiWy1p+kqt0GHePkq1cfng0krN+mX7GkrRiOWVGkrVaz5egkrZTbYbiksLV
IJ0wksM/R5h2ksPQ207Nks6bZl2KktAsLlQBkthLUJzNktkJBHj3ktlv/Ga8kt31O4QRkuARkWqGkuO0yWbWkuWE+YXukuygxZLK
kvR7Hprfkvnj1GwVkv/Bo5klkwJh8WLNkwKJC3N7kwKTD5vvkwb94XH5kweJHZD+kwjFbombkwybzHdSkxXE5JLqkxxL/o//kxzN
JYf9kyMv0WRskyshCXfRkzDE6HNqkzY+BZ7AkznM+4dZkzokqG6dk0CH44hfk0Y0TE9Gk0aB5GB6k0uKY2Y4k00z1Hetk08C5YcD
k1cWyZjAk1pe+IVfk113zm/Ck2DiUGvak2XU3mROk2p4B4Fsk2vIQVidk3AtaIDNk3KMdIxok3QD92aSk3QfFJfLk3ieam44k31u
7JCPk4GSeWptk4KiNF08k4bJM57ek4f+y3WRk4sjUFkLk431o1GNk4717lQ2k4+E12KKk5PO6GF7k5ZCHk7Uk5ZhOJM7k5o78WiQ
k5tdxFpAk53f61wnk6B33I9yk6KsLlR0k6WGx4RNk6xH8p7Gk69hxlTfk6+R0m4Rk7C6K3xuk7NiQ5ywk7SVm4RXk7ZU+YXFk7gb
P2W3k7hPBnD0k7uk4JaZk77GLXHQk78FHpkYk8O5hHAZk8VBnoSQk8cSC4qWk8nPfltqk8qm5YJok8yHPopkk88tK169k8+fu5SZ
k8+zTow7k9JsdIuok9OTN4g/k9Ps3pYLk9iThVgwk92uvI+Bk+G6t4bFk+SArYjIk+cml1iGk+cuf42jk+0E0JX5k+6BwV9bk+76
TXFIk/GMKnoBk/InsWwHk/QTuHE+k/iC0mzok/zGjYFVk/6+D3falAAsWVaZlAFe+Y1alAQ7zmsQlAST3IAYlAZ/FFzAlBOFg2d/
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
maM+e3iYmaO2uU92maZ4G3bcma2Om35rma7WWoPZmbPSOoR2mbZeTGYymbmmpFRkmbm5OV4Ombxx/3POmb4ui5DvmcPdlnXVmdEb
6k/jmdJO4G+RmdN/IZAomdPhv2rwmddvt2NQmdh19oqlmdplKWmnmdtCXHKomeCrVGPOmeGhw2mJmeNYTIXxmeY7hU5SmecHqXDf
mej3D2clmexTeYz1me1JIV6YmfCBqYLfmfcXOJrImfgZIV4KmfryclfCmfzcjpY7mgFUUVTBmgJlL07+mgLs72BcmgNHJmA7mgN/
LlP/mgVLX4cqmghID4jmmglbMG8rmgoqU20OmgtAs2TBmgtl3ohVmg8oUGOgmhEIzX3PmhVUlVWsmhk1M1R7mhsbeYpMmiERiGu9
miVYio9rmijMKlPHmiqa3Jg1mi01DI4rmi+QBFB8mjDQtoJfmjgKaGaTmjiPClxPmj+0eJXHmkJPgYtKmkoUYVo7mkyxwFwCmk08
BZUamk4ej4Vimk6QmVYimk9WLWE4mk/g1lbymlKoSZTBmlR5WHdemlWRPY+jmlX/hGPEmlhLDXZbmlnziFjrmloEE2bhmlq6AHoT
mlznkHCQml5m+IWvml/nuo1cmmZVrIgJmmegfHHbmmskqFtdmm8NKJn2mnNRP4mKmnUVZJlnmnU/aZ5FmnqAJoZ6mnrt5oRHmn1r
o1Kvmn/cc1KXmoPAf5ZWmoRWjl64moV0G2/ymottG3VFmo7hvISUmpImL218mpJFboijmpJzoI/ImpL82XbxmpnkyU9umpo4hWym
mpv8/GTdmpylKni1mp1tvGpAmp2R1malmqZ6kWQNmqqUip+EmqvHzImLmq+Ln5AVmrAwSWy0mrMUi5tPmrSVEHvBmsGneX/SmsJB
nYMams3gKItGms+SpmqfmtInCmcSmtXXu1zumtaZwk7TmtpVqXsRmtv4xYQnmt3TKo59mt8C1HuGmuiB7FaxmulJNlhAmuzZC2WD
mvZYuIXmmvZnq2Ormvev034Vmvsjr3zvmv9DxVD+mwPvJ2wtmwhp0Hw7mxA8+mRcmxFGg3G1mxhhwoiNmxnWyliWmxoxInELmyCO
cWHfmyJ8R5tFmyuaXWqzmy4BiWAimzG9DHAfmzq+tnUrm0uULYq+m1PkXJ37m1Z9/Zpxm1253Yc+m2Bv85y+m2NloJV/m2XiG3ir
m2YDqXfem2YR+Igtm2bML2VJm2ifxIlsm258sVV7m3F33ZJ4m3ZG6E+Em3gb4YEKm30eungDm4JgWnkwm4XND5BJm4sPH1r1m5Lp
737Um5kIR1UZm55dGX3tm56cGWNem6OxdJE9m6Sm1mIjm6V3n22Gm6d8I58um6pNL3dfm6rn8p9Xm63d2WC/m7I0/5WXm7NJ0Jf+
m7VJ/Zpam7VSGGWIm7eWEmdVm7tWKGeQm7wqzVCBm70hT2okm8Hb22pim8UPg1wdm8r/8p2Pm9DmWZeOm9HuJnd3m9WFhljMm9uo
d5Zxm9wT5ngMm9xcUmKZm92M62Cjm9/w54Wpm+h7VZWIm/QNYpfIm/X+c4N8nACQ/WdWnAEst2AtnAifPo0ZnApOgZwgnBDIqZ8M
nBEArIGYnBRh9YOinBltUIoPnBn04lpjnB3PJmTFnCYfdWOGnCoaXU8snCpCwoSBnC6W1XUxnDCOzlAhnDV4gWZtnDf/t3kznDuB
Wm1TnD1XEH4ZnD/k1prnnEEtzVPInEM4fVBSnESWwnV5nElv8Fd9nE2YLXHunE2rAIocnE6jcVZ0nFXfGoXlnFarxWQ4nFba23mY
nFoEtHS2nFtu1Z9HnGP3mmwdnGfDKJjBnGoXHo4DnGvNF2z3nGzR4WcQnG0Ru3xGnHHRZGt5nHaa2k9onH859ILQnISgjV8/nIeu
p07pnIe6A50InIlgM5g/nI6Rt5lGnJOLUVwEnJP6h4wanJmrCGEanJ4Wa28fnJ4694nTnKH6XJk7nKhi312+nK6rVoI8nLE8yG5N
nLOcF4d6nLfdQWRXnLvAtIKnnL1642bVnL8m/mQynMU7YHLxnMbCDIQ7nMu8xlSFnM6ge2rbnNViWoVbnNWlPGYMnNg/3FjAnNm7
RFVFnN/4T2s4nODlV48AnPBmDY9SnPSfr3mAnPVaGIGDnPoPEHlvnP113lXknP3SeVQsnQAdJm7xnQ+YTnGUnRSkX4xunRrOfYMR
nRs/t17ynR6PdGfwnR+ec1j4nSANwYDZnSAtQlVunSJHY04JnSgPQXY2nSpqflbGnSuyZ3FZnSyizp8InS0ivp0AnS546FdHnTDC
rV3enTHhEGZbnTH0h5PEnTH36VtjnTJ3eoZNnTKhN20cnTPb0X2LnTxsuV9vnTza0JQBnT4jM1bKnUJEzJdanUJ1mXtnnUMzdlUu
nURt2XYunUa4LYTcnUj8mJiqnUrQR1/5nUyemXLZnVBRmH0WnVPJ/VRVnVc5OInknVfQr5CUnVgxJVe6nWSZ3GE6nWa5SIa+nWi6
xFOvnWrOlmEinXGyQpxdnXI3RWTqnXLMW2E2nXft3XaZnXmuJGavnX1f6m5anYOfcVPonYT7uHi/nZfqxFrQnaIiQpThnaWQ2lnK
nadkmFqhnak6umwwnalo0oG8nazZaZ0onbEy2HA1nbLEdIh/nbZBdX1RnbckDJVtnbg5y2APnbojRlQjnb77r5n1ndBJz46JndN4
2GDzndqs9WYHndzoU4gNnd6BhX9IneDyAorAneHQkY+tneSB3k+3nebBtVtOnetJqmDSnewSoIVLnfd3yF4Wnfkub5ylnflqLVEh
nfxJ55U8nf4JfYpFngJzv2VgngjEoJPQngpnUX4rng0AqWdKng11v0/Ong23VpD1ng+h+lkInhFVJ4k3nhKejIN6nhLoeYGTnhyh
IZHtnh5TS5BQniFsmYxVniMuaIbxniNg/WyFnisnepSmnjB0Ro5vnjE2VFKEnjH6jp+GnjKn/153njoU6JeynjuNQ3DxnkJTbV3g
nkKMSojLnksv8YMWnkvvDmZ2nk/oj3WMnk/z0I3MnlHuc4gxnljKD4ucnl2R8pd3nl7jYoG3nmCZj4mUnmIxZZo3nnE9EVv/nnKE
CH8knnPoYWlYnngUqk6gnn9h8Z2bnoLQYYpRnoX0SHvdnogSA5nunohXwZAenpJWRmwenpeUd048np7+VHa6np+ZZIZonqBFapVn
nqMdxHjFnqWWhJ5QnqeZkoJPnqlYjmyHnq5zu3qPnq54Xnx6nrRmYJO7nrVlJpT/nrXfIn2unrrjmGrRnr0rBU9CnsToJnJAnsqm
DnDnns5yjHy4ntORU1VBntRR5pTwntvnBXuBnt9DGWYvnt+lk3MgnufFTFCznuii75minumtiX5Nnuw46JH1nvDsJIKxnvVU/p3g
nvb0+ZT0nvwp3IsSnv8AombjnwEPGmt3nwWoB5DynwX9j4YHnweEv2EGnwiOkYrznwjQGVTunwzGRpXvnw2RuHmmnxA+64urnxGA
9Wx8nx11d2XUnx+5Xo1CnyBvvWrHnyPZ1HJqnyQGcHXqnyVBaVb4nyWIDnhenymMGIF+ny0YgGKtny/J/GC2nzKRCGYtnzdCDmQ+
nzorNHtRnz3LYlSRnz/m8YeXn0D/mGxVn0Kh8E6kn0SsXHGjn0cd/VcQn0dXl4USn0sfbGXAn05UIFRCn06qWHxrn07T2GAan1Aa
2G+Bn1D7BY5sn1SiqU8rn1lJT55tn2LNJ1ehn2WO15t8n2XvnFPLn2btLGG3n2iww37Kn2j0I3g8n3AG8J3En3Gq31MAn3kvxFbN
n34QWmkmn4D9glrvn4ekhINnn4966Jfwn4/R3Gown5FVYIsUn5T6Tnzwn50EtHUmn57N6XBfn6NYblwgn6SuMpvgn6cPumVVn6fP
Amgxn6kOEm94n61+e5Dsn7O00G16n7RJ0mCsn7cTVmhRn7frDZEdn7hzB38tn7tq14oqn7wbA2p7n71M9W4Kn75P4Hejn8KAHIra
n8P95lhMn8olIXTXn8sUto1Kn8vAd0/vn8wztYGnn9Bd/WsLn9D56oj0n9QLhZFkn91h8nnjn938eIGmn+Pmxpt+n+QnEGRdn+bL
xoPKn+mZRWyfn+x4HHiEn+1DWVJAn/B9pHV8n/OPhU5Zn/Q5Bpkwn/c1OHbun/lRcZ8Ln/0XZlByn/1GbY9On/1PDXOwn/5OG1CF
n/7OOVFhn/9/pHLDoAMTVG4JoAcYI4J1oBny6FAGoBqnKn8uoBxUi3WxoB2pIoOuoB2xoU8goB8cEl5RoCMBg4CVoCtog4mXoC7Y
HYuloDFIVppioDMy9mxdoDUanV7hoDbPw3gsoDcApoduoDcm7Fy4oDejXp8joDfacFMloDzXTVkZoD/RO3xYoEfxymBzoElFb5hw
oEl8U4P5oE112I3KoFR3eJFZoF3TtowDoF7cpHAFoGJmz2pgoGU91pGOoGtcdGOmoG5vNnG5oG8BN5QyoHGt5GinoHHLS5qsoHLn
MVqSoHMp+JMHoHPhvYx5oHYzW5EhoH5ilpOOoIAPSXghoIN1WWZ0oInxv437oIxU0ou5oJNDrpXEoJYymI5ioJdXc3GzoJeJymoi
oJuowoLPoJ1qmYYdoJ6XX5Q1oKJIEntgoKKm/osJoKX5nmbAoKbaUXd2oKgBFWCOoK4dGW6moK/hpmtJoLJ4y51LoLS9vY9aoLYe
6YisoLdzxoB+oLeK5mjeoLlAQ4nQoLl2VnuPoLuJsoWdoL1xe2ehoL3YyYYSoL55RGy3oL6wmYPSoMWBsnGfoMb0N4TdoMpXNpr0
oNfViWCMoNhTn06woNlhwFouoNsslGAvoN66qWCNoN/4rVKboOIYTl7foOTP55pPoOTlKpStoOU36XWuoOdObWvyoOfNEpCRoOgM
uovaoOquSJQZoPJOG4HtoPJlkptwoPrsvJdFoP0Z357OoP6I4ZEboQBXLXcToQDj51BDoQJbSYTGoQLl8FucoQTOznTCoQrk71XV
oQ0L6Hp2oRI3CI2moRKItnSMoRd4eGVToRspCXLsoRwM3XtCoR7yFFv3oSUKpnVxoSnEn5dYoS7azYLvoTM98nX/oTa4nXSaoURK
jpEVoUTmWYfToUbWXHKroUeRxYg7oUxuJpzvoU7WVG0+oU8+KZ+PoVP7BYAqoWGJUIURoWWMkXOBoWkE73o9oWxycmpZoW9NO4MF
oXMEaY/SoXgbOXR7oYXir1lnoYhSoIRioYi4jZrCoY0daVcpoZIXe3aXoZO6NmieoZQh506WoZYt3oUYoZZd11T9oZo3mH5ooZuL
gJ7YoZ/Tb4MooakfWlU5oan5fI5HoatOC4ThoavX6X86oa9Z9Z9GobF491ASobKKfXcCobTf35LiobXWM1XZoboisI6WocCmd4ad
ocDC7HBeocmPdZ8CodCR72+XodIhb4L7odJYZJCuodacAo4codhhK1hIoeHcS3uyoePPWZQhoePaBmeroeY2VWjToecdlJJDoerp
SYIXofOtamdDofO2LX0CofZxy1+jofef7lUkofj1H36wofpp9G+QofwXJW/rofx7+pnqof5l4plfof7hHXTJogGzMVOJogXC5G6t
ogZkJIiOogaW/pxUogevu4wHogjBQ5XiohJWnXNxohxbf1F8oiGm8H8AoiIA2IiqoiKZIpAdoiOTk4lRoiWBUouLoinWd1u7oir/
5FR8oi6gNF1vojBgKWJ6ojHIQWD3ojJtb1sSojWpr3dNojXvz5Q6ojbe1FoMojg4C28jojoQwX+2ojpO22GfojsuJ0+nojyO8mwp
okKgwIwcokfsrGdUok+VtogVolJNCE8Moleb/4Lxolse43t4ol5+eI+/omTD7JKJomfK8XuhomjpwI+Mom2ZCIgXom8c1VzkonJ4
X4kGonrpm1O7on7SfYTJoot994vQopFm5XrBopPTCYNPopdLnnUzopepf2RvopqXqoy5opti312aoqSZ/H1IoqiikFH9oq0G1Wbx
orE33X6/orPOulSmorbX3mqEormv7VXLorwnHXkuoryZGVHTosIqTpTqosK9GYbnosQEW1LAoskc+pMAos8ukp1kotTFK4GXotUo
dIO5otojfGDBot1Pw2K6ouCTsWi0ouQ9GXProuX6IHubouvf9H/movL6LE5wovgr7pq9ovoEulaRovxmg5PCov5/loAiowm9dHtl
ow20KI4QoxMZbl7poxOTF50SoxRuGGQ5oxdHEpx5oxdLd2Baoxm/K5Qioxu5x4ChoxvU8JXdoxxTaVpGoxzm2ZJ6ox6p7ZROox/6
aFsfoyACsnfHoyZVR2/boyolLYKGoyvsCma7oyvws2qVozdAwF6BozhsdWNNozmMAXCYoz0mho0Vo0PGd1Syo0S8s29Io0qudGiI
o0vk/HdIo1BTBXwNo1GEUZ71o1dmDGLIo1kBOmuqo1t2lpxJo2SD33S8o2yUB33No20p+msdo21TVJG+o22GQZJEo27zwYJNo3n6
S1Mro4LbF15Fo4UhkZwyo4baWpBVo4qHD4eVo4qjKWBoo4te/4GJo5ianE+9o5rVcIe2o570mGKqo6XSnk+8o6Xa3XiAo6Yc8Yli
o6dS+Vcso6e0l2Kvo6n+eG6Wo7MjW2yko7YoKlhQo7cCLGs6o7gDLpgOo74PE2Tmo8Ae1IbEo8InV24Uo8cOjJkZo80Oq2dAo9CV
WZqNo9TrO59po9Wc15xto9cf7VKro965BXiFo+Lm42iTo+qnQlSQo++lzJR5o/J1poO+o/Tf2oVmo/iLWGDfo/6/qGkDo/+QpIiA
pAiZN4KXpAmlhH+EpA3hSGAHpA9dC4cppBMGnZlspBPR053spBVmuWmIpBWLkIrFpCXomXpjpCbFjnZ+pCvN+5X7pC80uHdmpDSD
FppqpDUPMFxUpDjc5poQpDv06WMqpD2qb5JmpD417XX+pEL4lINqpEMOuFOspEMvWWLUpEaolHA0pEbqXGhnpEwHs1uQpEyqSo76
pE646pzmpE8s5WtApE9ujZpbpFDcz4uxpFG0ZVUCpFoFdJpGpFzdw1KQpF0sOVC8pGN9p5fKpGSBZWc1pGWiDZPfpGfLQnfhpGn+
mpO1pG5W914CpHCVxJxEpHGEk4EMpHzx8FW4pH4bnniupH9NmZP9pH+HLodzpIAU4XTOpILLNYvEpIlVsoNspIyh2nTmpJM9FF1h
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
qetC9JTJqexWP3xXqe00UWgUqe2TrGJ1qfEdkXA2qfW63WXzqfgwWHsvqfj6IIRaqfxClFPhqf1b1lLbqf3GaIumqgThdmGnqgjF
a5dpqg5hCJazqg9MzZMWqhIOgXh6qhteJViDqh71KVI7qiAQ4WpHqiDMG41HqiHIVoT5qiZihonoqikly26YqimG8Y5zqitvE2oF
qivdVWSlqizquHPDqiz0GXINqjAew2mAqjGf8XI4qjKMXWxQqjmskFcCqj5zqZNuqkGCiYAsqkL+OJxYqkNeZI+xqkO21HZtqkUl
LVhyqk8rNWKhqlN7rZChqlS+foGAqlWt4XCdqlYcFnNdqlZ49XFtql+K12RnqmLwX21EqmPJSoo1qmf76o21qmpo3k8cqnD4pGa3
qnGIjGR+qnWitm2JqnY2VluAqn1lg4dIqn48joq9qoKMq5FtqobA+ZPlqoc/L3D9qok7qmvXqo5geWMOqo9Dx3z0qo9Mbo0lqpAz
wYQtqpFUMXAJqpJTMJfOqpTmBXckqpzAv5JIqqQOa3CHqqULZ1dnqqqM/5Ciqq2vUmi8qq9FW47Kqq/ZCH+MqrI6NGJzqrQwDJs+
qrS+D4EcqrlyQ1dTqrn8GYcOqrxiNoAVqr6eFp1Wqr+QEoPVqsc82pnCqsppvZ+Mqs7cGJbDqtBq8nDsqtWC84dxqtaBCWsJqt6w
dJOfqt9LFFDqqt/peVBKquGX8VKDquHPSYP0quHta2CwquLIUouhquinqXycqui2AZiGqui4dJHVqusL9ozXqu3S1Hy6qu456H3h
qvQbZFrtqvfnjFecqvkaYGzZqwIm953XqwKUiZUuqwWzhYoTqwq5xJ91qwx+QIhsqw94o3JxqxBLlp5NqxmJ8Zk6qx2LZnVvqx4V
DJY0qyKfimtXqySxvHpMqyc4JZ5mqyuETWYmqzBlPoEwqzK9znyYqzPSgpqlqzX3ioStqzX4nWDwqzjmToV0qznr3lJvqzx4PFTa
qz1CJHxdqz5kN2q1qz7suIgWq0ZaNoVeq0gnhVhdq0i18X31q0mlyleIq01gH384q1QaAo+Pq1U3921/q118F5j/q13UuYzKq1/H
bGrKq2GJjGdkq2JnG5HAq2hUv5Sdq2soU54Yq2/IQWFOq3l/eFk1q3sJ/2Wsq332E2YTq4FSE4ZYq4cq35m2q4wKb2lOq40MzZft
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
r3iZjYD8r3oVEG+Er3qPGG4Cr4AxWHY9r4kIMXKDr4sXLJexr4yKY2WOr46vZU7gr48k0F4Tr5Neu1FJr5eMUJAYr5e+loqkr5oN
rm61r6KvmmNJr6ZyPWOUr6aPMWQur62pcGV/r7SH4m6Ar7ab53oPr7qAk4o+r7ql6Wglr7rAV1PUr7ufcJWxr72vyYXTr76UK4c2
r8C5OJrhr8NXootEr8hTwU8dr8iiWGVdr8x/Nnlpr9Af2YiVr9SjEIerr96oRmPIr98HwGndr9+YE3C8r9/ZNVatr+N4D1oCr+aP
mYZMr+l48HYNr+peynpdr+3i51eer+9UU14xr/W5cYEur/Y14H+gr/bFfVF/r/jbnWD8r/wEBHg1r//DQoqVsAc1MIZ8sAepnHBd
sA1FlJqzsBKtZoKdsBX/yoG1sBrvU5zPsB5vzJutsB/QMJSPsCC90pzusCDlz2vDsCVdhJBKsCZ8V4arsCwMhWZusCxTZ4GRsC7a
pIE3sC/DHGnOsDPoZXHqsDV3EGcksDjQ12ACsD/CanEasEKJimxPsEWvGXwAsEcih5Q7sEi96GkMsEoKLmp0sFJLwXXWsFKndoSy
sFa7z21VsFpGEJQ3sFtlQHLFsFtra4HosGbT1oPLsGiy93nysGkLi2nRsGzO2mdgsHj7WFG2sHsc2HzOsH1Yalm1sIJpzo4hsIaq
fpyusIiypVbxsIoVL4eksI5wJlodsJdg65LBsJekvYB5sKJc0GiysKVncpSMsKeNpGWKsKgswpUnsKtLhJZasKtTGoq5sLaKkITM
sLaiuWnrsLb0oVtUsLdMglUNsLmCFoxksLm3GIGisL6Yc5eksL7cq3UOsMCE95BqsMGGImm1sMHg0GepsMMrRH+ZsMTl/XjUsMfT
zlinsMiz4HIGsMjtC11NsMqDfVJysMtXMpnasMxoe1Y0sM25Vmd5sNo3DpOjsNqfFX1KsNwUeHHwsOKxqpWnsOK3dZCasORhJGO9
sOY3LGHtsOeLz54OsOmRN2mXsPn7Apb1sQbCu3QSsQkMl3DbsQrJk1kBsQsq+Z69sQ4q6XFEsRODt2+CsRP/T2TRsRQBe1CysRXm
Ums2sRakOGeasRvoJIBfsR0SfoScsSPgi5SLsSa0F5dSsSgDn29CsS9O62PPsTfU9YxnsT4PplJtsT/HCXmQsUD3fVZqsUD/H4/k
sUq+lHsAsVhFxpALsVjnHF8HsV/KdGk9sWT8hWe9sWafXoCfsWfkF2lzsW7UCpuhsXhepFr0sXkCsFndsXq5zFWXsX1rwFJYsX8F
DYV1sYYq/nDCsYd+kGnmsYnmEIujsZLLdHPIsZPJQ2mSsZc02YmCsZ38IVrpsZ6zN1YNsZ+5HVu1sabU0YW7sadK3WSWsahfMmKe
sanNc5yUsatQTZcksa1ECmIUsa9cO34wsbHj21fLsbRikI9RsbfnjWSFsbs8/Z+Csb90koYWscCaKI33scJT/X4uscZFNnKjscaY
B2B/sccl2p7Qscuz4GGBscx41Imwsc497HNlsc6mZ1C2sdD51WrVsdTpEFnksdgYi57KsdkD9WEEsdmJCo93sd58kHO8seKiD4Zd
seYUJZthsecSnGhysedecId0sekAiYkssezFNFiHsfJ/44qasfgHY2nVsgR7UFUbsgnerVQOshY8151yshfE73+eshf4223dshqk
YHR2sh8v0HFTsiqMoJCmsixtYZTjsi6E4pyosi+mHJlxsi/St5KnsjErwX3Rsjh7MVSzsjnpwmOusj6EK2kqskaIvpgvsk2dMXe/
slRKrIPzslW4V42AslaxL5Vwslsya4wesmMO7mzpsmPJm1+usmVbPWqYsmnojnQJsmnpyobtsm03Kpzosm5XS1lFsnSkXGbqsnfj
5FtRsnyruWlrsn3USE4Usn/v3YwnsoBWp5GEsoZ7FYFzsoiUalhWsokahI3iso3XoHYjso/mKF5SspL97JA2spMsillVspPTWH2Y
spjEYV8esplU8I4LsppqcINcspwAa51BspyVdorxspz3SFWosqKTUVdBsqQTeXwxsqQ4g1nQsqcEvXJ9sqety4TosqrHuosgsqtn
14l2srUNa2u1srbzc3mosrqgtmzHsrq6IJDAsrvbupuzssIqyJS3ssfFkXMsss3iaH+mss7RiYZPstCY/Z1+stGP1p7PstIAGXfs
stJ3hHhUstK/cWhzstLK+FBustPPkFPrstPp9FEOstle2W67suSXxY4osuZFn5GCsulTLI8FsusmzVWLsutZCWR6su8x2n7xsu9t
GFkgsu+35oX7svByo5RvsvC51m2WsvHLEHMisve6wnWysvkly4wGsvucGHM4sv0TxJCQsv47KYb7swFjUJlDswUwmJW0swkT4HEQ
swxhi4PFswxyDH+bsw1SoJlRsxcgq3wEsxhumnSdsykZJlQrsywuiF5esy6BopEGszDNVGUQszG4InvVszMVWXDNszPgPVTWszWC
o2EmszqPTH79sz1VBFkzs0BOsnJQs04G2Wgis1igK4xCs1oYxmn+s1yxF1OBs2AkjG2ts2HPhJHKs2fOYWTAs2oVyITFs3HSeHEo
s3Jz73nus3PzTHBus3Q4j4cGs3uC4mTcs3/FRoiys4BWd3Zhs4I/gFi3s4MhlHVGs4g3yYSns4nHvIKus4n0L1Sbs4qmK33ys4qp
lnuTs4sLMmlgs4zQim4Ss4/k5G5Ps5KAdZx3s5gxUl7As50qcXaEs59JQWTfs59OuWzks59QllqOs6EGkpgSs6Kq43gas6O2H1yn
s6fltG38s6pkBWuts6y8dYmVs62k+VeDs7NImYTXs7/zHJ9Ms8Ad21jFs8nJumPbs8uWa1MUs9SLIGh/s9Szr5X0s9YM8pdOs9a1
9XCis9qhW4Igs990h2frs+aHk5bKs+rQ+Vd2s+sviGHqs+8pGIALs/H5rlmbs/Q6Xn7Hs/awXoSSs/rXzpZPs/wfJIddtAiQ7Zvm
tA9sdHyhtBUF7F2mtBjGMHOytCKLx4uKtCfoFYwYtClgMWdptCnhyZi9tC8g41HCtC8vRGXDtDEwVpMitDRECIaKtDdsQIQ4tDpl
H1BYtDzq4p5otEgFG1qotEnvuXs7tEtoB44ytEwZBZJgtE7g+GD6tFPKpoT3tFPZEY2EtFrTSJ6StFx1TIwStF83pXNDtF/XUIok
tHQbT1X7tHQqmGJbtHeHOpW9tHke6WRHtH0Y2FmStITxrF2WtIsRLHomtI0GU3DUtJCtM3iJtJDZLXj8tJNCumMBtJXk81J7tJaG
t1g+tJgcIHTItJu/O3BmtJv2fE55tKaZ5YFBtKj4+48wtKuRTHOCtK1Nk3CatK1RFIvZtK/2t19htLE2vJMetLSjyng9tLbON3UE
tMDOtI3utMoyPY6UtMzxw4LgtM/TFloZtNEJp3KBtNTpWU/VtNabkFP4tNoH8Y8ttNoxR335tN5mjmEetOEPFn9ttOS9WU7utOVu
OFz6tOnIdHXTtOnK91DNtOtg2mxztPFuOngGtPM3PnfmtPr0LnXRtP3HfpPTtQpKmZ9QtQx+f2n3tQ2ywoXbtQ++62rdtRGrbIMA
tRKnYF4XtRLwBpsftROS4JuXtRRRtGactR5//VBgtSIul1dRtSQYpk4HtSdwUYSXtSl2dGRmtSowYXaAtSzbEmtZtS0qxozWtTBg
DVOrtTVzcIYotTbFSm6etTmDnGTstTvNXIKLtTy9aYyetT62+nPjtUCO63EBtUR2aohZtUS92Im/tUT2d4t1tUmpJIJ6tU/OtoSu
tVAW/4kutVDUrVEVtVQoTJxftVdFS3b6tV0WgVj0tWGa8F6atWaRfHydtWqANmP1tWsTvovytW+y7JVstXKnElvgtXPoMZsDtXs0
ZWeltXxVkZaVtX2iaopKtX3VmZXatX32bI7xtYYjXU9ztYcdOFt9tYyOLWFttZA8417ZtZJLEondtZehyJtJtZ5+23DEtZ8wtlHq
taDF4HGQtaKpJHuqtaOGbmI1taky4XYstayai4+LtazaxG68ta66dZYqtbDjqIaItbVhHZmstcNODYr5tcNYNZ0qtcOHPVmXtcX5
X2x7tdC+z214tdLY+l91tdw0umaUteTwDnK/teka8Wggtemx0mAWteoJMZArtezAXHO1tfCpCYkotfeFwFITtffqdoF7tfrua3YT
tgAUHoZXtgXuX2UVtg0Yd1UqthHIOIdvthgqaF0ZtiCxLH0AtiSP1VNZtiVQ4l7ttiVRm41ytipd02twtjDddpIPtjI7QJVitjlb
YFC0tjmNKmzUtj1rC1l9tj6XPmJDtj79SE6/tj/r+32otkX/mGj8tkdVPYVHtklkZ3YEtkn9qpF6tk9UKFggtk/IjVrntk/f9XUI
tlAhqGBStlN0mJBOtlU+v1t6tlpADIjTtlpj/1OZtl6FDIE6tmImYYrGtmiZi3WvtmqU33qRtnGkqlKWtnSzMWuxtnWxTV/4tnuP
2k82tn3Xn2cYtn6WP1EbtoOvsFsjtoV8Z2jwtoe6CH41topAZoE1tpFBtoectpKN3GTutpLMhomntpkwgoVVtp1+5Grytp4Lr1c3
tqCd35DMtqNEIFaEtrLX81VctrkQ4J6Xtrt1yXqWtr0/tFvqtr5FRIaOtr6jr5MVtsU5x4WItsjIS5frtskrKpsetsq9rl8yttgP
kn/3ttspEIY/tt2G3Y+7tuZCVH2etuZsZplqtuq5A5vqtutnyn6btu8I7W9jtvD+sF+EtvELyWIetvQD0I8GtvaT2FnZtvfb83zT
tvqyUZRItv9lXHQhtwChUpfMtwfIQFNXtws5cmwTtwvqhJi0tw6pWX6LtxECX1rTtxKJjJyttxLX0If5txRsnVhZtxVRIoBUtxut
V50Ztx4Ah34htyFs6YBRtyQqeE8AtyS6IFSKtyXzWHs+tyX81XlztylFhE9/tynRJFJRtzIcp2a0tzMLzlYJtzQBKJrOtzXcWo44
tzcr2XGJtzo5fZEQtzqclZnmt0FV95eBt0I233Oat0JAxpVbt0LRR2cNt0Q6mJ3xt0V+UlqUt0Yf5oOOt1OVnm5pt1fDZX8Wt1gh
bW4/t1glMJfUt1tkKJ8dt1v8iIcat1x6wlYgt19FVJuRt2ATiVBat2DcG5xCt2IREF5Xt2K7a5c+t2l3fHl8t2r9WW/it2y7d2rU
t234IlfKt3B/N2S8t3GVb35Kt3PuN30Dt3S+0Gk3t3dqtlRPt3jQ6XPqt3shwIJ9t4N40pP0t4YyRIr/t4e/vHvUt4kG0pEjt4p1
M1qut4sLpJOqt43CZJQ+t5AyP2tut5HH3HBKt5pvR4FPt5sjNInPt6KVNFIyt6tI2JuAt6yiEWX4t7TbK4Q1t7YvC3lst7apkHGw
t70eDlGyt78bS2znt8MBv5XXt8WWtlZ5t8Z4XGCFt8oc3X9it80BiYeht825yo7Lt9mnsI/8t9n5h3Vut+JQeVl/t+SKfYKst+VV
3osKt+kTNl3ct+vmZl3/t+wpEWwKt/C8nnCqt/OaWXuOt/XTrGeAt/ic42Ajt/rxF1Btt/7vPWgTuARvknIQuA0Gflt0uA83FnIR
uBGWGl0WuBb73FJBuB2qromPuB/m7nO+uCOeYJRBuCWztZ2QuCeiuIPIuCsTBILjuCwcVlJpuC4a55VTuDTOH35UuDfrmnG+uDin
XJuVuDoOMI00uDxVAp5kuD4thZ2iuD9N/o6XuEY+fVnYuEfAAWeOuEicuo28uEqAEXsauEryN4YauEsB6opCuE04IpiXuE10M2NU
uFGteYRcuFW0dnoKuFdFC3VruFxFLlGduF4UQIHSuGEasGUkuGO1j3pluGRTYY+suGWqQn+nuGYiQ32+uGzKVXZ8uG0AO4huuG8X
55MnuHBgL58luHG0zFlzuHS26E66uHkurWpauHo3j1vyuHzhHFm/uIGn2JlmuIPUKGT9uIYZcm5tuImb+3EOuIzVH1ZAuI732ogD
uJD3z5DWuJb6kpe4uJkBc5ShuJpme256uJ7DMJSjuKG+SHcBuKL9ynW+uKgIyVzeuKsjOZjbuK74LIzFuLXZ7ptXuLfM332IuLkr
7VmCuL2NGFxsuMAHVJv/uMBPbWA3uMEucHFSuMoYRncDuM7QiZ0JuNLvYotPuNXwWYgyuNf9dmMVuNuxlY4muNwPo4xsuNyfVX2S
uOA5Go8EuOEdinl7uOK1vZQCuOK9QU9IuORY0Ig0uOknflc2uO2gfGOIuPE5KHh9uPXAH0+juPdhNo/TuPiUn3DRuPq5WnDauQEm
F4eFuQOnZnBxuQX5Y3z2uQpYhYBsuRFF+mDtuRk4k4uOuRyTe5OUuR0rSVILuR3LMpUkuR52+mBhuR77a2auuSHn72MIuSM5ImtS
uSSQ026MuS9dKFEuuS/jaFUiuTRHrE4huTns7I5BuUFWWFqnuUP25Vh/uUxj6099uU4wnndruVKsXF7VuVV1k5C5uVp534whuWDl
5olZuWK1O2WTuWSn3p3PuWnThHvuuWyKLIGduWzbsX9SuW1OlnthuW1RroGvuW5OXmSyuXHsjGMjuXZvxWztuXoelpHHuXyVXWW/
uYSxpnasuYarCGTvuYj6gZOQuY+JFnGNuZ373pqOuaDYSJwJuaMDQ1pLuaj74JCxuanwMlvOuao0s49kubKKPYZHubSOk3dMubbV
kY/PubcaL57iubjL7oxIucmFU54qucoERpXUucu/zn4TudnaZIIpueGvOHeoueWumlGaueYUx3LfuejI2XiouetYRW7Xue+hVYSY
ufMcoVMDufQ3tnevuff4gWoRufol85Mtuf3rHWNcugrggmL4ugumzF4Aug3XB5Lzug33R2HsuhWaUWpLuhqze2qsuh9ie2shuiC3
y2QxuiIyAFXXuiLY9G2ouiP1knp7ujAB1XCrujcfJJcjujgXY3S/ujlfkZYnuj8drXc1ukSRXnGoukp9QHPfukzE5WhJulEZBX+L
ulQSaZ62ulyxVYHuul0xb4vCul5cu5cQumHrO30/umJ8Nm+3umbfP3giumjB9FDRuml6l4Vyum40/XuRunJkHIsTunT0g3RIunaI
SXVyunsHoJQruoPVkWAhuo8xllXWuo+2P4zSupjIpV0OupxrjmZQuqF/ooK4uqTuHJOSuqfx+F9duqqsEXjEuqvzpoPGuq5+q4cs
urJJlJYYurNNOVcMurNu32IuurOzS15rurdXKH5AuruWTXrTur5c6YTbusAou4deusFSzXMuusHSnJ7DusO+VmJjusvJ218ous4P
lp1DutO7q1iQutQY1ZhNutd9IWYzutd+bWM5uthhx1HLutkhqFNgutnQdG4puuJroXj2uuNONmL7uuhM5F79uupJwVZ9uuqx5V8n
uvMxgI5kuvhEFHgnuvkIMIcPuvt6eX1Muv/tyZAcuwPhwE7AuwYoyZ1IuwqH6psFuxOwRFPPuyEa64qUuyGScHTvuyLL5V4EuyN/
UFdduydoDJAmuy+fSFgKuzOU/GAYuziDeZXWuzq9HIhGuz2kbFebu0IoZZDZu0YjDmEbu0sV0Y0ru1EfYVOlu1PSWHCbu1P6LJj2
u1lo3psTu1q642BCu18Mr1Ssu2K3cIoGu2Sg84mhu2WOI1Aqu2siv4nfu2ublGEnu28l8JdJu3Gm0lXHu3bpoZcxu3fZJFlAu4Ih
NHUSu463NoKlu5Qjam6Ju5fpRZPUu53wIIgQu6Z1XZ6ku6vDsYtqu62iq2YPu66h038+u68Y2XTTu7Aa44fwu7S0gFa3u7Uo+pl2
u7dJjomZu7qEA5Uiu74FxZgfu8MbXW01u8uZN3f0u8z0Tmjbu87cjnXxu9BylHsYu9CbJm5Vu9LBzk8+u9TSmn+vu9ncGY9Lu9pu
toc/u9/eTnrJu+E7E04su+cXRZCeu+klO3+iu/ac7YE7u/a3umOcu/g0zlMyu/pn3m5Qu/q1lHPlu//NSl8PvAEnm29EvAGKzn1F
vAg5BXqSvA5SOYyyvBARcZ48vBAzH1gPvBC5/1FxvBDdkGu5vBToC5OyvBif7FKovBi0gG1ovBtZn4Z1vB5U6WmMvB8xoXcpvCiy
tHwPvC0oslVvvDK7+4b4vDTdi1fEvD3malc6vEiXd1KGvEpqKZGivE6cjYRzvFREw3h5vFdXg5hDvFgMe2NXvF8b6IrVvGHFJ3ZR
vGLE5VZBvGQvqV2TvGXTbY8QvGh0ZpzevG75xX6HvG8nNXZvvG9RjX0IvHDbLnNmvHKuxWjUvHX4eGEjvHl4ioivvHut/X3dvHzl
en0svH/mZ1ZzvIQEbWHevIY64W6zvIe0j3o/vJKBBFjyvJVmwJRuvJc+7We/vJ7yrFUSvKDQm4zvvKEuLIJhvKKnCV7svKVrl4h8
vKZlxWM9vKasql5nvKzWKX+1vK5nL4qyvLDPlYTrvLRhhpS2vLUE4nwdvLpKFZ0WvMCsx5FpvMFCBZ5XvMdGuJdgvNEK1HNjvNIU
D1ymvNYg7HXtvNt9ln36vOAK/2srvOGOhGnNvOX3TZEHvO4v2IJ2vO62+3UXvPCJnYOovPJjVFjIvPa5No8xvPkK2YYtvPp2W3CB
vQbTkV3+vQjCtJB0vQ/hdXBMvRCDUmFdvRcbyGf/vRomfWUavRqDXZD/vR9EmoIfvR+nFZBavSBRo3UyvSGnsX42vSGwLZrxvSUE
aoY7vSYRB3d+vSbjeHlFvSdNJmumvSqPlGBOvSsBAFVxvTKTa2LSvTnYoVvBvTnkoFyqvTqByFIcvTqU3Iz8vTr76IC+vT82YIHr
vUHXgJnSvUKIs5AHvUT1Fmz1vUT4UpHfvVE/YFaOvVIPsn63vVanC4UMvVftiXHJvVsY0YTRvVtLH4LsvVxsrWYpvV0qFYPAvV7g
HmKRvWAAaneWvWCho1xDvWDK2VpovWEB0o94vWTukHS0vWd3lHu1vWkeIJHEvWwSOmxjvW6EYlaAvXG2KVlPvXIVUnwbvXXjFGCL
vX+cAZ5PvYBSzmbyvYBiaGZ7vYSz5HwavYjBak4YvZAEun1SvZD1aIxqvZD8q4FKvZHjgJZDvZS8fYPuvZZkMZ0pvZbYwYewvaEm
b2UrvaNhVYZQvaRsKYsOvbZi0W/+vbxUAHxRvb38rmVbvb9+KY1tvcK1VlgevcLf2VYbvcb4gH7lvci5VGqevdECkVQavdEcQk5e
vdYMBFpiveCPmZqYvenTEVslve2d/JvQve/4r5InvfOInVcyvfRreJv3vfpl84kVvf3u6lZ2vgg4dV+tvgvsOZSyvgv0tITuvgws
HFVTvg9UYHmxvhUa9XnrvhgpmVltvhhflpyZvhnn346Kvh37E3rpvh5n7Xjzvh67ZFahviJKwYbyviKICmwPviOqYZjUviTn+28H
viku+0+ivit1o2MNviy4wGPBvjC+JHBCvjEDsWKJvjHrd3lQvjIrxId7vjLrSmgevjN31o8PvjdesVpzvjtMHJW7vkUY4ppOvk9A
YoV5vlAoj07ovlMB11tYvllo72YDvlnwz4yovmDTu3rLvmFFqH21vmFHcGe8vmG60FzLvmMox5kBvmuIGnJDvm2nsIiYvm38p39y
vm+wjXz1vnVQW1sZvnVsyYAlvnz8xnAwvoEAU3pkvoI/rnE6vodvSm43vpMDnYJdvpNYF4Qavpn5ZX9avpoiiXwXvpsOanDSvp8F
z1CovqKU7m5bvqXPg5pAvqiujVrYvqoDumaFvrDiNJTavrLe9YXfvrZmL10UvrxZEXXvvsPvq43tvskzkHLEvsmru44svsr3n2Px
vsyBfXxwvtLq+24rvttFqm4Vvt/J7lfgvt/s32iuvuIKf3NXvuIpLn9QvukI0lWIvuz6SHRKvu+Oj1oLvvHe9VMgvva04IPdvvh5
nlBZvvldLJnNvvmuQXexvv49gJlrvwFZKY2YvwZuy5FJvwiXXnTEvwkkoGrSvwt44ZhmvwxvzZjnvw46GmC+vxEpCnS7vxPcq1YS
vxRoZW0mvxxS1Hx3vx7Ty5BBvye4zW0jvyuIRmHpvy8wRlRevy9kumlVvzn5aI/nvzyPd2cgvz1jVnPVvz2OKFefv0BnF4knv0B7
vXW9v0G/X2y/v0dHWFX1v0nBg2cdv1j5SYPEv13+FZwzv2SorIJLv2UhDIpjv2bHD13Ov2eLfGXBv2tDglbLv3JtdFdWv3XCjng5
v3eGb4Olv3lp9ZGwv3qJOGCTv37Wx2W4v4Ct0Gkrv4lCL2cyv4uTQ1Piv4vwWphrv48IRk/sv5RulJEtv5WgYlnDv5jSxnK7v5wJ
+owUv6Xijl7Iv60L+nQbv68tY5YGv7V4fWOMv7a+5H/Yv7pGIH3Gv7r8d4DSv8KCh3Iev8RvjIyqv9I/gWQiv9rhtX13v9vhQE+B
v+HhsWtpv+Q2HY62v+byHl1dv+hEqGddv+lu2oLav/r8911rv/uC331CwAQgTWN6wARqQn6CwAWv7XafwAx4d1KywAy2dJj+wA37
BU+ewBGEaIMgwBJmhoZgwBNy24RLwBPMrotgwBQ0044+wBTw/n4jwBflPov7wBrsgppFwCBmhoHAwCFyQFbSwCGbDpETwCQ8iVUj
wCiBinXOwCojI4yCwDF0CJ9fwDXRt3qQwDZjm1gIwDeW3plOwDgl05V3wD0e3HkFwD1HYXFkwD5MuGMSwD9wX38EwEB7c5VUwEbR
Non/wEev2o1vwE8QMYW0wFGIXX8PwFOryZeXwFPZdlpMwFRGz29KwFccpnhHwFta/m1qwF6eelzrwF9h/m3vwGKMp1NbwGuPqJRp
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
xARYqJ+IxAg7c5yixAlNbJ5exAw2NWpKxBVcnGrgxBXP614uxBbufHDwxB0iql6UxB2oBmHkxB9T2popxCUgD3kJxC+agpY5xDFm
oG27xDSH83rNxDe04nUjxDkhIW42xDnTO1GTxD0cXY+zxD1TsYrexD+5/Vy6xEJ7EnGWxEXI9JuFxEnP/VdcxFGAnpVqxFa3HlsH
xFpScJ4oxFvEsmv1xFyXQXY1xFyhtJ3mxF65u1bAxF8LtVWrxGeseZWGxGoicYUdxGqStGpDxGt0vWrsxGwej23WxHCnhE+/xHdy
eGN9xHeq+VfBxHl21E9axHo4s5z9xHuxN2PAxIEkDJMyxIQ8hnLwxIdpxpr8xIgHH2OdxIi+Ik8/xIz3G20txI2oV2KXxJA99Zsg
xJBE62mfxJB5JJlBxJKhuJUXxJUBEXBtxJkjfZznxJvf21mzxKKrjXaQxKVy14aDxKbSlo4IxKmbeX/HxKxGVWz+xK6/I3xixK75
F4aqxLD+kIdCxLExpo1ixLGNJFIvxLWDJGemxLbFAmgFxLbFgHHDxLf+Y2ezxLgqx43BxMBsOGeFxMCdNWcKxMibKnQNxMpSBY0d
xMsDd4xexM1pc1UrxM9y3YyjxNC00F9ZxNeTPWlDxNxy/5McxN24ZGUCxO24n5KLxO9Jym6RxP2w2HDgxQE/Z3iVxQGHbnOqxQmf
/F//xQ7/jZAtxRRFFWv7xRa42XKyxSOeqJtlxSc2wU9nxSg7z26XxShrPpaQxSsT74bsxS3BHmrNxTPTPl/txTdTCnVBxTe+5Jgm
xTxs0VpRxT3yUHnfxUi76WfixUoaMYa5xU4t04MqxU62AITzxVD7Tmc6xVGWZ2npxVM2mnaWxVo6o3ibxVqTDmqFxV93P4kFxWFQ
7X23xWSLXpO8xWovbpxhxXRyyG57xXfT4Yo3xXt9TJl5xX0PDGVBxX5donW5xYRfNo9TxYWAz3+kxYbfiViexYeU/YJKxYjHEZDq
xYo5PGxFxZctz4SpxZ31EGF6xaWhbWc0xaqroFn5xatJ0pfRxa1FKXsWxa2qu1UoxbELwnDvxbTY+n4MxbpgD2PvxbtbmI1Xxbv4
ol7Wxb6WOHkyxcIeC3AuxcX22ogmxceVEI+WxchcfJHNxdDOJn+3xdOzB1vKxdelOZBPxdz3QZRkxd6qHnA5xeC/fVvwxeDs1FDV
xeM1/2DWxegaiVg8xfB88mjaxfjle04rxfuQYJE1xf1ODJZUxf99wYd8xgO6sZboxgYTymYExgjGL2gdxgkhbXB1xg6BaVnTxhL1
wWVZxhMEFWRSxhMc+nWsxhfC5pj7xhpeXGZDxiIPNmFnxiu2/2K2xivhum1fxi1MJk8oxjvLHF80xjw7a5GZxkLQSVCZxktBI053
xkwI5ZtGxk0R/X2Vxk1LrH/Lxk41iFB1xlZAk113xlmdgoXtxlo483y3xlzaKoR+xmozgHNMxmpEOl9wxm1fDJy8xm2ofnHpxnWK
y3SixnWXoo0bxnZrdX6uxnfyKIeixnjCuZ2rxnmsq4DKxnu/X2c2xoHs8I7axog51VG9xok5I5nDxowHoV/bxowTeHtTxo/be1vC
xpKhbnVoxpUxr3UNxpZji2jfxp3wVnYRxqGueGtYxqVU8V44xq+qBYa0xrHjEFbhxrVL7lIbxr4xKVq8xr5QDY4Cxr83gYUOxsMz
AH5qxsYPtk7zxsb60ZiVxsjUwWRJxstGa4Gyxs4aHm8Nxs5KV5lCxtAwYYabxtQsZlrdxtcDNIHFxtcoQVODxte66Xajxtielpoj
xtkFu3bZxuNPYl/9xuZgRGyUxubk3GYWxudB6XgNxu1MPmJpxu/hel3RxvNVn10RxvbPZZZSxvqato/AxwJ8KlxexwSL840/xwi3
U3n3xwurEG32xwzz6FdXxw4sUo2ZxxSS4ID0xxeWdG9ZxxnSX4Eqxxq2FVawxx30rJfvxyHHzXmNxyJNvGjPxyRXlJdRxyiGcIH1
xyq2aX//xzBrUmyhxzDCGWNDxzM4jHKsxzXfxXTYxziZoXmjxzrljG45xz5kIFSwx0NkxZbvx0PecVoKx0Yn6WFkx0pvj3Jvx1MZ
OZWlx1dntVMYx1fDOoDlx1l6dYArx1sg1oJFx13eUpdZx1+cfpb9x2Bcvliux2DyMZBpx2MOImLWx2RsOW7jx2YxZYuYx2Y6zZyn
x2k+lp4Tx2o+5FCTx2qs/1msx23ujZqDx3NOs1oex3NUdl1Lx3Os6H8Dx3ducF5dx4XZb2t9x4k+Y5Xxx4yTjFgAx475lpy0x5DA
fWAKx5X+ZW0vx5bMpobCx5uDvXVzx6AxRU7hx6Koe1FLx6X2lHnkx6f2JHeAx6obiU95x6t5zWujx62y4H8zx65Ui30Yx6+UOWex
x7QLwHO/x7eAp2D1x7hPblIZx7lyIlfsx7qNN2Vqx8G9tn49x8IKgFsQx8QV/5l+x8UeBWuWx8o21lzPx8zEaVr5x86LoFBUx9CW
h458x9Q6qopgx9kHnpvBx9l+G449x9zA8FHVx+IIBIq1x+YMVFcix/B/1nedx/Z7m3Hox/dKFnsmx/kWY5tRx/l4Om3Sx/zp5Wtx
yACQWm0AyAPAfmgPyBHdDm7zyBKPdpdUyBc4AH+6yBhBK4vvyBkJM5SpyBtKBnI7yBu+O55zyB2z25cCyCF8MmScyCJx1oYryCVO
0noMyCa9jpw5yCbRxmq8yCjLzoyQyCs4hojbyC0s1Vk8yC7CxIGKyDAZymZzyDIZt5sryDKJS4avyDR1Q2NoyDR1mZejyDutppgZ
yEC6aGwByELurI1ZyEMFhonSyEckLFyHyEg1BX0qyEg2Fld3yErXPmYYyFKSzIR/yFaZI2USyFyR659myF24Ao4EyF4Ow3vjyF6o
KHm8yGd91ZUzyGgGVoCOyGoVTmIwyGocfFzfyHDfoYHWyHX0jE8nyHcG4X8nyHdbiIk4yHkZa1CdyHv2CHNOyHyTXJyTyH5iRWYi
yIhzppasyIo4k1glyJCGV2amyJKTZGhMyJO8oH24yJRGXVunyJbteYwsyJfB4HQ3yJfCN4KzyJn4T1rlyKH7jIahyKLriG4WyKMz
LGsiyKSosI/tyKhDek42yKh4nnwDyK6QQ4ClyLKItIVsyLMH1GahyLM5x04AyLOAEGbsyLhPj4P1yLoouJvpyL+Md5U0yMB1fWXw
yMHIxZ1xyMMUM1eyyMOsdXFXyMQt75KlyMSujFyLyMW9PXCEyMZKLXM7yMZa23R0yMvRN4OYyM38EGtzyM4lcpFmyNHDalcXyNaO
D3MoyN0DWHkYyN9dFZzzyOGZGVNsyOIyEJFnyOwDlodkyO2SMpZeyPRRJE6KyPiEjpgEyQSXJWb3yQx4MJfEyQ2e/ozjyQ9Np4KK
yRASQFXnyRBCNF+qyRLsyJMdyRN9AmSnyRYnZV4oyRb6S4VTyRg2w32pyRiK6pSIyRlxDpFhyR3wNWi6ySIYTHaFySPTSICjySaM
Q5pXyShSEHr/ySlkKV7xySu6BZ1MyS+6GIJiyTT2v3nWyTf0C2cRyTrWBmDeyUAKj5wDyUmB622wyUqDl5nVyUwqyJi1yUxsEIpl
yVHkVE9ZyVUVh45XyWJoNpnMyWWhFYgryWZLEX7eyXLT4nuvyXNk5msAyXPOX34gyXadB4e7yXjGpFSOyXwtv3ufyXw5oZ5ayXx1
MIaoyYD9AHv6yYFW62GbyYcgmZ1cyYliFHiRyYt4sWcvyYzuJmcXyYzuUJkryY0iyp8qyY9asFM2yZD8WH/gyZmlx3rHyZ60jpiB
yZ7x5l6HyaATemnEyag5W5m7yakmK1Mzyas1aZ9EyauhlJhaybFI9GQKybMQIH4QybTbf1RAybnb7mc9ybrT9W/FybxZCYS9ycHc
v5O6ycWxxnaNycgcx3xByctoF4YYyct5i31Zycx4xFeFyc77v5kfyc9523m6yc+XCVXhydJQ4YhaydJVOWolydj1Ho+3yd+90F9m
yeKgcGVoyeic/YXwyfEszYTCyfFdeHZkyfMq/lyvyfNamWLRyffDSFKfyffcP2FUygTRp5GJygpTM4atygz8cFB2yhDDoZrgyhGD
zZySyhHNK1U+yhIIP38GyhrZxYfbyhvtf2QoyhyKiFmQyiKCsXvZyiZMv103yjG2MVz4yjYpWJQsyjf6R2+lyjyY1GJmyj7DmV32
ykaappdkykrI24T8ylAR83OVylCzSJNGylKsUnp8ylOmQFKHylUczYQiyls8lYGkyl3ZB1vnymIvGYXAymM+IVE3ymYzgHIvymbD
3oaWymcvuZH5ym7FalnrynCkZGDYynQLuYGoynXcIHwYynYYY1vHynbG2YdBynsgBY38ynwlDVVeyn430JaLyn8yvlykyoCIyYyt
yoty7VGWyow9NZouyo1IdlX0ypC3WGubypMxf4CbypbUN29uyp0IV4kcyp+DolroyqFauGvOyqMdK1kFyqehlF8myqoDfp4myqrn
32w3yqtoDYaYyqucE5oAyq1g3JRjyrBxVo0+yrQSGIXcyr/v+l3mysDYp3Q4ysP2wnwgysVP3olwyskc0FIQys03Jl/7ytLCaGSE
ytb/X1DUyto0Lm4EytsOslrRyuB9R5kOyuYZLmD/yurcW29iyu42aXcJyvNqQW9WyvSG2WOFyvjDsGjQyvu0uXVnyv2skmcDywFX
2pXyywUcxpb+ywZanHmRywaCvlBGywmH32ulywsMfpx7yxKPHI5/yxT8dFqPyxgly3onyx5KrV0uyyA2cE45yyP3jWcoyyRPsZkA
yyTU0Wk5yyX2SGioyyvFpGoOyzD+EHFCyzJgKpbUyzMRsm26yzYa71qqyzoZolg/yzuYnXUuy0N1r2nWy0cUlGwoy0rP1Zosy1Pf
l4MDy1rAmJj6y1wGKWWry11dGmJdy18+RG+Oy2GQu3C4y2K8hJX1y2nsZIaSy2rWzlOyy2wyNIp5y3BYAFDPy3EJjYDry3MMrolp
y3YE7GuQy3oqgmnjy3o9OJu5y3qid5l4y4BCRnIyy4Hf25Iey4d7Dk/ry4e/aFCRy4qETm2Ly4tjrpyOy5PUd04ty5Yl6V8ay5cJ
lZMDy5kHJWVKy5kl/4TKy56zJGxDy5+1cJ6Py6Ja/m2uy6gN/lphy6rBsI9Fy6z/f4fCy7DLUnj0y7mqlFV/y7sHnW8Vy7talmjn
y7wtAI1dy76muXyIy76tn3V4y7/yb49uy8YfIpYay8v33GdYy9Mgxlbvy9avFHIOy9hAsl9Vy9hHw5SBy9tSOmGIy+E3T1uJy+Nx
9Vxhy+WhEXsiy+YOnFM/y+koAF++y/BXcm0qy/qyIlwzy/1eHpVgy/2qNH2CzAAOgoxrzAa1in2zzAqwhlY+zAzEz2s1zAzSRmbE
zA4fEl/ezBHL2WJZzBNXvW08zBUaI2bozBtLxmmTzByFhnjdzB1Y63N3zB5ZHY2HzCfUmGYFzCunmYjezCyW1WIIzC0sjF7czC1X
oFz0zDRZso0AzDgUal9FzDuQPXpbzD65KHlAzEA00GGazEGta2XgzEHKWYP2zEIbZWIYzEaYMWSSzEbRCZRazEiICJyCzEiiw4V6
zEuV3VR1zE3GWV6GzFHlJoCPzFk4Bnp6zFlPtlCmzFrrBmgBzFzIrlD7zF5oo1MKzGIQCmwrzGTn7VsnzGWcZFQXzGZOhYILzGhx
Xo9ozGt4QY/bzG0UHXybzIATS26+zIV8T3wLzIgbV35BzItNGlb7zJInHpsVzJVJqlFczJXVSWw7zKA642VrzKDd5F6TzKQ3cF+A
zKb7r4OpzKd+XXF/zKk8GYBOzKlwvV4ezK9zr4s9zLCLBlIVzLH2G4aQzLP7eVyXzLSJkWGszLcsUXhazLdNxIftzLlc4pa/zLom
y5p8zLxuaY7gzL6uTH+/zMN9NXclzMTec2O8zMarn4DPzMcZgoWszMgPF2LKzMmBmVm6zMvZaoA4zMv/pZG6zMzgc4lqzMzkrWPj
zM845IbzzNWLgVCDzN5A34mrzOFNNInJzOiNgpkNzOxDVXBjzO0z7FJJzPUQ9ofYzPaampx0zP5nJm+kzQDIxYiuzQTTA2tNzQmc
RGEWzQ7OdFeJzRViTJlvzRXqZ4tQzRi4+U8NzSAKzJa4zSC1iFWuzSPPflkmzSiJDptAzSvthp64zS1GUVOkzTHhj258zTNcwZk/
zTSEEp5+zTdKZ4RVzUO9motpzUWFrloyzUWrm4rbzUrTZFRrzUwH0m/TzVXAKlpvzVdWipRezVeQh5z0zVnCqIRpzWivyowAzWnF
sou4zWo1kIGPzW817GNYzXE3l2rAzXH5qoi4zXKDh4ymzXOv93NRzYIOT26GzYOzyFVyzYTL2Y4NzYWQBWdMzYZ5WVr3zYu+rmdS
zYwqGo9/zZKbG3BUzZLt71qkzZQDSXddzZVCIWW7zZboGFwJzZjr0l9+zZo+rVCrzZ/kj4Mbzal0oXSezbvkh2Wpzb72R1SfzcBU
LJJpzcSzWpM4zcXDLYJazcjjJHAUzcmEjXfSzcnYoVSNzcoCHG+gzc8XXHwkzc/Y6Y6QzdJlTFuGzdRVTnmOzdZlIW+6zdoKxHQg
zdrvjWcuzds4NFJNzds/Omanzdtur4cuzd6aomZ+zeKPSlOpzedSSWBNzev8rJJNzewgbX3eze5JU4Yfze/XEp1gzfJHz1myzfME
T4t6zfeP84CdzgSaU4bYzg599I/DzhA9+1FbzhUbuWETzhehm1Rszh+SooyRziZ7/pW4zid+iYMBziqRaWU7ziqh94ghziuq2YQ2
ziydv3JXziy4lXl+zi2mypPJzjGeglS/zjM0+IWEzjRxtpXKzjR01YeOzjTUX2fPzjjZ2mPKzjxWin2/zj7aHonAzj/P/HW3zkHe
iHkSzkS6Dn5azkmx0mwszlAYwl8kzlM7spYHzlXW9llczlZkGWv/zl2lrY8ezmG86Gi3zmL2FlttzmRFCFYpzmb2sHbzzmeyoG5l
zmvUcV/DznBtB1tyznCBZnkfznakVZw6zndR2I3Wzne1HlxWzngg9n11znoa1lx1zn2tonN+zn3rf1fTzomfyoBSzovwl4tozoy/
unSwzo3M3ICSzo/4mn75zpJBElY5zpJYEo1wzpUEuW8dzpcZMIefzpjs0XKKzpkqeXPMzp6s55fqzqmvJ5Ffzq+2a3VUzrU2E2G7
zsG3l16xzsoNkIdrzsr9CZffzs8twX1+ztKdCpUoztsrTomOztvrppNjzt6ID2qDzt9ZqE/UzuAzfl1WzuH8d1LRzuXXLltLzulG
I2e3zuwzFnIBzu8V4YTpzveFV3Y0zv5yEoDbzwZGe22Bzwl7FoDVzxDyoHr6zxNa44U/zxSwIoP6zxcusVCPzxtoEnNLzx5UqISi
zyHQBFRvzyLL9oohzyNYwWUtzyPJc1hfzyRwwl54zyZmbIfizyem+nijzyiCbWFjzyo2VJw1zy/W3JcszzDsfmXszzEBNX6xzzjR
JlQGzzs+5IVJzz18g1rzzz/eZWbQz0G6UE9Sz0l/VGdhz0t894kEz0unZ5doz1BEbXFhz1EB/ZZjz1GjyJbmz1Jpv1X4z1XoZoxi
z1YweZXZz1rTeYOPz191c1mqz2AYsFFXz2M+KVP0z2hD1JNhz2xOzl/nz26+7XiUz28xo5eDz3AoxGuRz3fTCFxGz3gYC4zpz3n8
RGovz3yNhVZHz33fX5q/z3+8rV6nz4SX64UGz4pqm2ADz4uGC2zhz4/PG575z5BTW3LMz5FNsoMEz5QnulhNz6bzg1lkz6dFXIvS
z6m/+owEz6yNEpo5z68HyIMvz6+RaXdBz7EaD5tnz7RJCVPMz7mAxFEjz7vTmmWoz743HlFHz8BKQW93z8Bo1YIsz8EyHFaKz8nF
B4FJz9CPmZoEz9Hsp1QWz9cpYGLYz+RywYwlz+T7JH3nz+zNRYvlz/LbGlB+z/SjhZX2z/Vc2o4bz/bbXJDQz/hzb5wOz/1HcGYu
0AB422c70AHqS2gN0ALz8m3/0BcFn50h0BdnCITI0BldNI3U0Btl8ouU0B3j5WXX0B3kgZ3u0B/WSpZl0CBZaoQQ0CGN61oS0CRL
rJ6R0Ceqx5kU0Cf3J2la0Cs68Xlb0CvQoVac0DV8Aovj0DdLi4cR0DjhOWVS0Dj3UGrz0Dp9y2FR0Dqw7mox0D3LPWIH0EZ1um5m
0EaHJ4VP0E2GTVYR0E7dXljQ0FDf7Fr70FDu2E/50FPYm2gI0FSaxmYZ0FjJDpOi0FmzT3b50F1sYIxW0F1wJHQI0F1wW2Ae0GCi
D58s0GJSEX730GUNCFO80GUuCHXY0GipKZ1O0GvU419J0Gw0am0k0G5vtF180G84NI690HGX/30t0HHWj5jc0HiMc2P80HiSSJcF
0IN5nla+0IRvWoZt0IiApHu40KN5bpmn0KjSX3O90KrN13N20K95P4Lu0K/7y3iC0LzpRY9V0L+2KnnZ0L/AG5VG0MCcFJFR0MIk
y1Mm0MJSiJ130Mll53yM0MpIco7W0MykcI3o0M0Z4ZI60M3Y6npe0NJVGY0e0NgFqn9Y0Nw2dYgi0OACA2S60OJ73lnC0OrCuU/h
0PA5LZw90PQTXmiP0PqIpmMF0P5kcVWQ0P9GxWEd0QUyhVAa0Qk62php0Ql3PXIs0QrywW8Z0QtVZJ0Q0Qzwqnzx0RDVNJ9x0RVV
4I7t0RbcL2DE0RpQcZhq0RtKlYJm0RzTkII30SBoQV2k0SIFeW/e0SaBCoIr0SbzPG++0SdNAJy50SfHFHpa0Six5mUH0S4h4ZvD
0TYKR1ZI0UJbU15N0UQUhG/H0USEFmas0UYvpFeR0U0jNmRU0U3PL3Po0Vd6Z3z40V2GK4Xn0WA8aFsM0WZ9/k6x0W8Appat0W/W
oJvR0XRGzJdc0XfGtmKH0YV4312S0YY5KZ1z0Yc+6J1p0YkOHWxn0YopG1RY0YzkXIFQ0Y6QR5zf0Y+CFZe60Zf8x1eu0Zoz04/Z
0Z7sppmf0Z8KRl3b0aIKxVZ10aTC1n040aef0p3B0an7IGvg0bjg13eV0bm3N2mU0bnK6IW30bp4w2se0bqM93VH0bzGDHEH0b9n
eoZR0cQeLpp70cRiN4BZ0cZLOU/40ciX01Du0csQupb/0cwHrYsX0dSBwU7f0dc+soFr0drKu3MC0d5RhVI90ebJz4710e9f2Zj9
0fEz0GtU0fFxIk9e0fam9I4x0fqe63L30fqgaGQA0gL3213n0gNeMlhj0gR+w4L50gUDR32l0gsbDFDY0gygSlfP0hW7MV1Y0hnf
altM0h37f5IL0iOK6ZQS0ikBF2kK0jQ8bJ4F0j7TqWcV0kFv+Hi00kSau5S70kUms2bz0kUwxIVr0k9x84kX0liY3nIo0li+Y27U
0lmSEFld0l6wvlJL0mAHIY+r0mDxZ5f90mJefJ160mpyAF+20myUFnjD0nIj2IAU0nR0E4v20nZnno/e0nrct2vw0n430VhJ0oB4
P1Jc0oE6LZnt0oaSj3P50ocO6nLC0osGf5xX0o4Tg4Yw0pUpJWvA0pqm85SR0puPnJuy0p2sTpE50p6/p2oC0qf7oYkv0qvRA23a
0rVK2mhQ0rg/PYIM0rj902QU0rmSYoB60rpC71KF0sQ9Lntc0sYvjm/f0spD02aa0s4zEJqu0tBq3lLd0tFy3WWC0tTZsVaC0tWM
sYiH0t6spp0z0uFoPFpT0uMsc4mE0uSCqopU0ubeWVxr0uj+sJ7N0u6r3XLW0u9JSE4F0vDQHnfr0vUTQosu0vfes4mm0vged5bq
0vt1m5mc0v8FbF0f0v+l2WCr0wZF7VA30we59pIl0wli5nxV0wzwEmoL0w0oV4yc0w+SB5080xAx1E570xQpn1OP0xqIqJCJ0yIe
VFgN0yLbTX260yS0gIiS0yUC5Gmd0ydvtHZW0ykAWIFd0y4cc4O90zUbhJXC0zW5VHsx0zdsdniK0zghEJmE0z9/WIaV0z/kaZe7
00Cw4Xt200PG4prH00aUHmA+003r6Yay01BvYZfx01CyVG8c01HWH3OD01INLmeG01ZbuU/p01dRzJZT014YyG83018PXHRr018S
pn3Y01/T0YGL02GgO4m302LGFpIO02RCb5BI02aJmWz402gmk57R03AWeHL703CrdZwb03MFloEp03oAI1hu03szR4U903/97Gq3
04H9n1KP04qEQ33v04qQMHMa04sdoXpP049CvmB705Bj7l3s050Cg5Wv06wHlZvW069mkWpQ07Bx8nRY07F7knSZ07N06GRK07PF
kJlI07StYVvo07TSRXhu07XbQYeq07c32nqn07fJdYJO07uXqFqX08AUonGA08UWto2X08sugm9R083OKH3S09JQbmzm09UEHZbC
09ZUt4/V09aPIZ0G09kbD3u909ljBm8n092QwXpV0+Caxpm80+CsuXpg0+Y86o790+plrFNl0+27iVP10/BY71Lg0/MEP5Pq0/WC
GGot0/YlRpn60/sxRojQ1AG1smze1AJhL08Q1ASeGlrL1AgxnpIv1AlMtWGr1Ao4h29X1AzNzmbG1BFerX/N1BiJs27G1BicuIun
1BkE7mVi1BwnqVUG1B0Z1WCS1B4qRIyL1B9ehHAC1CSH5GHV1DLLTp181DkL1nva1DsaT2T41DwTEIAH1D0jml3d1D2raVEg1EAz
/YWu1EKG7XeD1ERvuH1k1EXlKZpg1FCdN2In1FS5/3LO1FkpXXkb1FllqoVC1Fp5T3DM1FslSGBI1GOdVmWu1GSvDntP1GeQBJO3
1GsqaXsE1HFeTE/Z1HUFzJMm1HXYH5NM1HbxN2L51Hh76pxv1Hi56IaG1HkRplYo1Hksa2ii1H40C57b1H+6UY7q1IBhy38Z1IKp
e5tL1INJrVLO1IcQ1Ji61Ienj3Ad1IiHKJqV1IsOcHLa1I2dKXIV1I/aLnEk1JAEwGkO1JkwQp041JmYWFPm1J6HwWvU1KD4JZst
1KVdcWHr1KY5Fmvp1KivFWml1Kub9FF21K8NF1cb1K+wslQz1LB3o4YC1LVQQ2Mx1LtA63s31L0ec2ll1MfyVpyQ1NBwEFow1NNF
wpoU1NN07mk11Nd8AFJK1NpeYZ881OHifV9W1OUoGIbc1Oj6qGo31OsgxXIZ1OzC5oDo1O7gxpWp1PPaQVhk1Pl4OH5C1QB5Glpg
1QEemJ+a1QPV7H4d1Qj+EX4J1Q0qHoa/1RJDZGl81RLGdYmt1RYGfVCY1RwZm3Up1R6AAWbd1TDsYp0N1TUvE2vV1TflBFOX1Tf+
OI0P1Tsysn6s1T3AwY961T812W7u1UcjfWxO1UltxFUX1U00HmXT1VMTqlxN1Vblu2PR1V3TM5TW1V7jZ4zo1WDxX3Cc1WE0WFTb
1WFIl1671WHnKHVk1WHtu4kl1WKrvlM11WK8iHAN1WZXOHe41Wa0I2511W2AGF/O1W2CD1wD1XCiDWQj1XL5TZni1XbuklDE1XdL
nJVX1YBjKYSg1YsFqHME1YxE/mFz1Y0OGlus1ZGhn4Pl1ZPzvHmt1ZTE24Fq1ZnsPXz71ZuRH2ZU1Z4GwHFc1Z+Mpmg91Z+8q1fo
1aM0XJ7T1aZkNWXm1aZslmqa1aloEmV+1ao1DF8N1bAioVlu1bGX1lAD1bPHaFl61bWIo2+L1bf+CWG21bpIOI/W1bucZYil1cSs
KWHh1ci7LFUP1couknru1cqObWP01c0hAInE1c1U7IB01dmcbHLq1dnUp09b1eFZG2ZP1ekviHgQ1etdxVHJ1evNK56W1e22dnZQ
1fBGSH191fC/JprX1fGtX18U1fI0FZnO1fK74XKe1fMXh4io1fxgamsc1gBSal261gNHrWVH1g8Df4t01hFr82Di1hPvY2XH1hjS
c2191iCatI9w1izrbVR21iz9424d1i5xi1yO1jHxWXJp1jZ5vGpN1jbGglAI1j1OOYrC1j7hjVPe1j9qY16J1kJ+dJUv1kjJ6Io7
1kw/13e01k1EWnRR1lKs7mVz1lToH3y71limkWGR1l0ZNXtm1l4bdIFU1l79aGUF1mPlM32i1mZkz26h1mjQ1Fpu1mysjU/W1m3g
2Yyd1nXpTYl11ndCa2B+1nd8gGu61nnlZFIq1nqkxpmF1oPoC3Pw1oPzZ5NX1oQVElhY1o/EhlIJ1pNwcpG41peV6JP61pi+BH+C
1qSn2oTE1q5vgV951rDJ4VPJ1rnHdHnp1roy+3DF1r15WodQ1r4RIY4v1r7fpHQk1sEJl5am1sIGonmk1sJzQnVR1sW0IVVN1sZb
goCI1srQXpPn1s32oHXA1tR8mWsO1tW/KWpS1tX4d4IY1temZJrc1txSS4qu1txfmoG91t5BNWi41t6fsnl61t7L9Y7e1uElb3oq
1uNDhWzf1uQsnF/+1uQ3cF1/1uepAl8D1uiMeXDm1uwW/4qx1u38dIP71u+G65rw1vLfzJ2Y1vOkEYwz1vTziJ031vVkGJQF1vct
OnHF1v9r0JhK1wIP83FG1wlWL5fB1w20ml1q1xQ+xGqA1xhwI5sI1xlIC1Mc1xqKaVtp1yGCd45Y1yP/w5UZ1yecroDy1yfMhHc3
1ykXhFCa1y78QW601zNGaoqK1zNW2V391zPo83mH1zboT16V1zkw5V8v1z5jJ4oo1z8JK14Z1z+b0YCU10E/vIb510SMC57010fp
a3XM10hxcWzr10xkM4bh100cXZpk115e8ll312AomovD12M5zpRt12RGMYCA12bWl2qv12gJMV8l12qwZVpC12uadHjQ12yIRVvA
12+UElAm12/+gn2g13EkFZUN13HfGWjI13Rui1n813TNU2R5132hhFzm14L9tVY11447D2Ny14/cBXbk15C5oYYQ15MGo5rP15mH
4lIR15q8WnXc16Hf520f16QdBX6016T4t5RL16Uldlgj16pgQW0/16vKb3rQ168LsJE716+gSn6J17Gm6lHv17PkBW9D17SHR59D
17Va8lHH17iXV1rq171CD5Gb18bz4GEL18pCEWFr18ysrZkm181GtJ5619Nh/5OT19VNnIHk19aKr1Vh19f0NY2u19tManip190h
C37419179HJL199ox42d1+Fwn2wW1+HRoFDv1+VatnMO1+pRjH1c1+uL12LZ1/AxeJwh1/AzW43v1/BptZhu1/boLpyv1/cmElrX
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
256hVlhE26DGuGRf26K0OWpV26TFImSh26ks6lYD2677rIRT26/zXocX27q9Mn29277ElV4828bN7FyD28qWroT628rLX4nH283q
I3EJ287WloFC28/Kj1Sh28/+9WPt29ICJIwN29IpyILV29V53nVK29fJOWHj29lLI1G82908KG/k2+VIw2rI2+h/LHd/2+84UVJ1
2/RKQpOw2/fZIJ3w2/1Um5uN2/4/7lcD3A0Es41L3A2EMXJl3BArzFv53BLX+26r3BQ8UlUJ3Blj034L3B1tTYwT3B59f2jl3CCA
plsa3CUtolwt3CbeAYIy3CcFqZn33Cg4XpYd3ClxcZ7H3DMta1EX3DnwQmwR3EQu842e3Emp0p4u3FBxGVjo3FF5g3XB3FUAYGM/
3Fv5n4jM3F2hpZXT3F/7tm8G3GAXJoh73GBNBXAh3GNj9olM3GyHY5xb3HJDnGeH3HJPvmiX3HJjQIpm3HSTSXaG3HUXOWRG3HX6
PXQd3HaZtWQz3HhFUJVN3HpI1Xko3HsM+GNx3IMddlEZ3IN2yYDD3IQXCVkd3Ixsk1Yz3I6tMFXp3JEvwJ523JR584pz3JS68npw
3JZlApDp3JyesJZr3J65IJXP3J9e118c3KDbz4Hs3LFZNlJk3LFfJ3HE3LLIW3gd3LSso2yE3Lllq5YD3Lu6o2yo3L7kcGZB3MCS
fYNO3MrWnFxQ3M8p+WzJ3NCd8GmL3Ncnf1xj3OLbe3Jh3OS0p10q3OXGYF1B3OfcLHyq3OoWvE9r3OyvR2Ks3O/3uo5p3PCSW2ZH
3PXxK4qY3Pi8i2hU3PxalHhP3QCjlXqZ3Qu8LZuG3Q2qP2WU3Q3t9F7D3Q+wWYeb3RLA41DI3RNRoF493RRb+Z2g3RbWE5ja3RmV
F11i3RodrIVn3Rv5hnoH3R4IiFy83R/N4Wbe3SMipl+M3SaKmV/Z3SkZqnBO3Smkvnq+3SylwmU23S1RcXqv3TK8fG4m3TTtSpYN
3TgrQIPY3T17Y4Jg3T3N1Yns3T7MW1Hc3T70s2DC3UFKyZJA3UGOBVeW3Ubkcnek3U7/l2Ll3U+lA4Fu3VZTH4zh3Vvdql6z3V18
GF093WINxpun3WUkD23H3WXPUZzw3WbicZ4s3Wc1SYKQ3W20NJAQ3W3rNJbd3XPUoZjq3XVZonxm3XgCZI/g3XpT+3SO3XwTEJ2N
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
42k1O2TQ42oof4mj424w223M42+3EZ2O43BGF4Wc43GuTE9R43p2XU6N44Bpt5Z/44byjHH444rGRFBd443UsZ2t447nsZZ345mn
G4vW45ykFoKj458eiX0e45/YQFPZ46dfcYZA46hpyVcr46jx4Y5t46pj3Gar46zdR4UU461zXnnx464G/Vyc464KtYQC47NABpOo
47dWLYvJ47eehG1Q474seYy1474txIUX48CXs56t48ULEHlx48Ub7ny248a6g33f481pEJij487so1H849CvHVGu49INR3hx49Kx
X4bJ49TcmV2q49h4b3bI49tFvoG549/F+18J4+Ij2J3O4+MxypwF4+NuOWMe4+OFbnD24+QKzWwQ4+RXUGAp4+XMSH9r4+sxBnJY
4+z83I7C4/VCLYYP4/vs8Zmu5AKLE40G5AY5R2fz5AwSKVQm5A33MVFv5BQDjGfq5BVzkWhC5BkIMVst5Bzzj3tB5B/PNVKw5CT4
54sz5CZQPFMs5Cd2BlCN5CeMNmsB5C138IR45C3JmJq15C/l7Hyg5DLXQ1Cv5DSrzGko5DkJDlan5D0iJ3au5EIhNYq35EJcYWmF
5EVhVHo45F2bAIrE5GHbgWTl5GIIpV0S5GSJ/4C75Ghegk5f5G2Dq1m35HAO9k7w5HE+MXhg5HI8fZHC5HOGBpy25HoZE1NW5Hpv
3ogC5Hqfdoue5H06aWIT5H+7JX+a5IG3cpRn5IH45GWx5IVKInaC5Iac5nvn5IuLyF+s5IyXm2GM5I50qoZj5I//tHjx5JNwcpI+
5JSLSYqS5JSXcnsz5JY49X2d5Jrk0Z8i5J9BsWZa5J+Mnnsg5KM0hFFg5KNqVG/05KcJM1b55KguvJNV5KndmHhX5K2ho4u35K8r
Ko3O5K9KvXXL5K9ysIXs5LJ8SobV5LQqfWvh5LkP7pdE5Lqci5ZV5LuYc3NT5Lwayo+o5Lw9Pn/J5MCF22l/5MIUImO+5MI5WYke
5MYx3GzC5MZo9lZm5Mec4J265Nd035Xg5NiiaE6c5Nk6B5xH5NwjOZZJ5OG+kn7s5OVwh3FP5OuFWZll5O4oe5pI5O+jt23Z5PAp
HZZg5P676H2Z5QIUd5ch5QNJhHSt5QOLpJTA5QlZWI8z5Q5NYZWM5RAoEZ1m5RFmamVD5RGu+nkq5Rc261s95SHfWGiS5SIdq4Dx
5Sbf6FDA5Sk0S4IU5Sw4Kn6M5TG+vpXA5TcJNWbF5TrDw5ne5UVrHH6c5UkQ5YdS5UkVin3+5UrCr5mq5UvBQG7i5U20T5o45VLF
tpTb5VTT2mNl5VvvTFcI5V/6LFh55WNJhVRj5WU8vI0c5WVhS2X/5WcAtJwk5XDM0Yna5XEFx2ql5XEeHXXF5XJKYVx55XMnm31e
5XbliFuw5X0wmo6R5X1OjIFl5X1++moz5YBw3VSC5YH6Ro1J5YKvy2pc5YYP6GQL5YZbXHxz5Yv0d1iP5Y5mnZPV5ZFQql0I5ZHJ
0Zn85ZMftmjs5ZZRyVxJ5ZmyModm5Z8ugGHc5Z9KfVUM5aI2iYDQ5aP5nXTM5aWqeml55agsJ53R5apFd1Px5a5NeVLa5a+zblLH
5bTROlTq5bXSDFQS5bdnW1lG5bjFO5qR5b1kiFxZ5cN9CFCG5cPojGoK5c2qR3Et5c7Zy5b55c+gHlyM5dFXMHc95dPI7o7G5dSr
eHrO5dS1j54k5d2jelux5eA/koUH5eJpF2oZ5eNvEnts5ePomFNr5eQAII5N5eiuA36y5ejgGXoW5e1V6Zpu5fQYG5f85fU/gILL
5fq244Sl5f4HK3SE5f7uKmM+5gS3AZN95gaWtF2R5gi+9GVM5g7yg4B75hIT1ZRM5hq5io1g5hxpplUO5h6n9ZOM5iQl6k/u5ik8
CVAE5iu0sZOE5i3QVpct5i8pbZi85jSnlJJZ5jTs63LA5jd+DY7o5jvMo4Mt5j4GXXaK5j+vXltC5kdgypuU5kg/qHYx5k3CQnIM
5lEV0YeK5lJUhWud5lRdo5K95lb1cV3r5llfuI075l05vJa25l4NoYPb5mVpVGz25maXjGO55mahdIf/5mgRO2WN5moljV/c5mrS
yWEF5mxyyZYM5nAqpHB75nMkvoHh5nOCRIfj5naWk5Lx5nc6oGRT5ni0WV895nm6/28z5n0CtFZX5oMqkIPf5oWqnXQ15ohwdWpC
5oi0i1zS5otIbmk+5o787Xmc5pFO/12A5pcVdYg+5qNxxWd15qOBYWXE5qZ/e27l5qnT1nsP5qy4YXHC5rJJNXZC5rLT3HXo5rRX
Rott5rUoioPt5raeoE/e5rgUvYCF5rvpSYeE5r483U4V5sbXm1IP5siOEI6r5stg0JAk5tJCxZFD5tN985C25tUGXIWq5tfETn2Q
5tkNpZcK5tpNf15u5txhL1Zl5t2HW4w55uB1mG3G5uDQvn275uE872Pe5uK9T1ba5uODypok5uOlk3cq5uVXhFFC5ub5OIuF5udZ
3HO65ulUzXbJ5u/PM1HQ5vA+lVY35vXU1JOX5viQ4Flg5vrjCZQ85v33uZWf5v6PZViy5wJUfYE25wJnQoGc5wNsJH2U5who3poP
5wnJM46t5womQmfp5wtFrVuK5xSMn3815xTLKZpN5xdvF57E5xgD2WjN5xh1iZJP5xwTJmh35yZZa56+5yvFaGn65zMcAJry5zUl
vXSS5zbAYGqd5zdtcJaB50Gz4pn050Hj6GaI50NZvZnB50XWRF7n50up9GFY51WRG4Lp515QmFd052HXYG/x52Hr+WXY52chIYK8
52jI5GbU5206fmLs528Ee5pr53H5MWWj53hz2m/h53+E4HLY54Ay8E/P54DfSnNB54FjcW4c54c2fE8L54hnzU/L54rzWGAU55go
goJp554ig1oO557LNFc+56JjInf856Kwepmp56PosHWK56uoHoc057OaL3Mh57oKiXzP57wFjpTV573rEosw58DTI30x58IbSHxf
58IyuWvd58LDp35b58Mx41ha58Q/RlUy58fpW4ph58pm54fa58tT6pS/581mC4FO583hnmBf587sGlas59Lk4Jw459Vg4YIe59kI
g59O59uM4p8a599rAk6o5+twD53q5+vk5G175+5AEo8S5+67qXUB5/DFs5mQ5/Kn9lcW5/MQ7mLk5/cByoDL6AW5on6F6AXSpYAe
6Ai/zFzd6AozroBq6A0otHfC6A22T3dg6A5vEYiW6BVqh4xT6BocT3Zu6B9vp5F46CJ8zlZP6CVGqYMM6CY4K4Rl6CftjpvG6DCe
slBb6DStq3rK6DYXEm6D6DkJmI6h6DoQbWMw6DqHsWn96DqipVZv6Du8P4pu6DzI0IBz6D2crn486EAj+WUl6ELlfYCa6EmPLFhH
6Ervk45J6E23UGn46FB5gJAi6FNISYv36FkgP1Gr6GTCm3fW6GWH65cc6GbpfZES6GmRoojR6GwkgFYL6HDdbWrQ6HQHVmJw6HT4
fIiI6Hfuf27Y6HoAYnGp6HvWW5Aq6H5wbZsN6IJcLFNQ6IRga4vA6I19W5jG6JAixpww6JOt61sK6JO/7oSW6JTdk3cY6JrTh4ia
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
9dXZ0Zuj9daG7YG+9ddEOHtZ9dpbLlAz9dp5KZTf9ds5pHVp9d1334Ms9eQwQJIV9eT1XnhR9ehFJILR9erTUp4t9euIfmSi9evk
8JjP9e45JE649fuzo5PR9gE1ZFkf9gPPHGZv9gjTn4BN9glLamEV9goAkZ8R9g43wYRo9g/1k5Ud9hDMG10z9hZdDonj9hluAZN3
9hn8LZjO9hrJNGVW9iORwYaP9iReklek9iRha2Ol9ib3uoFo9iim+Y/49i2kSnxN9i5x+k9v9jAOGJ4Z9jIYn3t69jN3pE9W9jaV
pW2X9jp0TGta9jtIlY+H9jz09mlb9j1r8GT19kQuY5zl9kmJgWPa9k8RIlx89lJnSX/j9lQT0YDe9lrBYo4Y9lzkeJJX9mgVnF4J
9mgx5oXq9mqadGld9mr+aHLP9mvF0mYe9mvVd2vm9myk8GWA9nBzP2m+9nRVv3Xy9oEGH5tM9oVnd3VA9oWGCI6i9osfI56d9o3d
FY649o/pj1Ji9pP4PZ8y9peO91kV9poeXZOt9pp/zZ689prnrpRz9px9pZxg9p0soItm9p2MF5Fa9qAu8lHB9qUDWHgP9qZjnVo4
9qxzKU8F9qz8y43R9q6OaWGV9rD8a5EZ9rhfUE5D9sC3WIua9sIE6Y9U9sPXnZgt9sTFx1o29scxOIDR9sdD8Fo39se2EW129sqU
8F3X9suGLln69s0K1Jwl9s/u0nI39tDfp1Om9tYz7pBt9thSKmMa9th0fW8C9trMUn6B9uEqkXNn9uWdFJLL9uxKyZOY9u4Jb2RF
9u5oHorg9vCxwZ7L9wFIHZAj9wW2j25L9weeoV2E9wgLeoX/9wuNSYuS9wxfC3hS9wzCVITY9w+Rs5QW9w+4O4LH9xEOh1Z+9xI0
D3sK9xR/dn1B9xv8fX/Z9xyYw3QX9yjUPlBO9yk0HnGK9ytYJoxK9y2OAIGz9zQ9IlAp9zSdApym9zU+d1AV9zVClGhi9zdKoZ5A
9zepSHho90vNlXkL9018IouJ91VYcZvd91bqlYJG91tUv41G915FFmZh92Rxn2ME92XSTWzP92oosljO92vnyXP+93C1/mTi93I0
zlCb93Zfc1tl93yJq19/936mmVxn94JPAWHK94y4i12l95LRfX1E95oeLGQH95wHbVQh950yoZJM96JRolZh96K623wv96NjJIX3
96mxAFpc96rdV3+p96tl0I9e96vXJH1r965y0E7n9651MouX96+m+3C597I9umKl97dAn1QH97dRSE6X97gGp2yp97izyHZ/978I
xGTe98j5YWCC98rKclAd981TQXfy984SYXPL99A4s59299LL7Vvx99MKrXHS99RqpHmC99jqplHt99qDMl4H99u/54EZ990jmIWF
9+GQRZHx9+U1xom89+duR3Wh9+h5H1d89+p+k19r9+zCuGTW9/CYDpDz+AY+fohg+Ad4qXVJ+AieqGHb+ApuJlog+A7OOZxV+BMb
lI+k+BY7CJ9n+Bek02j0+BiVd3ym+B1wZm5K+CCh5lgc+DzIO4nW+D4smm2Y+D8hSF14+EKWCYbq+EjoMH4P+EmzK1es+EtOTGQm
+FGldFL6+FR3u21m+F3KeJHd+GmMG5di+GnnyJQD+G49JYus+HqpZWiH+H3YLlLq+IMVk5Gz+IOQK4ht+IVXH2k/+IcZ82nJ+Ihx
6ogj+IuBRXEU+I9s82Sq+JH3AVvV+JTxo2su+JjBZJVe+JmCSWnK+Jq263fX+J/4K50K+KAVsVz5+KioGoiM+KtWbXiv+K2C8pFS
+K7C/481+LAqZVTQ+LS5ME+4+LVWP1bd+LYaCZGt+Lq9uGlw+LrbXX8g+L19ZFMX+L61Rk/G+MCW1nCk+MPE+nRo+MUU2lRZ+McR
RXUV+MqapnSo+MuuDGy6+MyjmZbI+M1nkmY/+NEEOpNT+Ni73ZpT+NslTIWr+Nu5ZpiF+N6dkovp+N6nPn89+Ob0u3Na+Ot1e2tq
+O8BHU+L+PIEaX3T+PNAJ4CL+PN3tm9m+PPwMGU3+PUT+G7g+PYvnGU1+PjMylQC+PvRoZU6+P+BJVeY+QJXQpJS+QxlM3/7+Q3L
1pO/+REQyH6i+RRqBJ+V+SB9XZLF+SGlsG+J+SLpI1+1+SMdsmn7+SdDXY0p+SjzHI3j+S8W8mun+TBFGFQQ+TSZfnTB+TS4f105
+TdUWmib+T3D8Vqg+UouQE8v+U0ZKFvD+VO56JYf+VUAS2E3+VYMbJ8U+VaBd4MN+WapLWdc+WjyF1um+WmzuGqQ+W0T4pyb+XK9
44i0+XQJhl0w+XnRB2Xi+XvVzVj6+YRG+nlI+YUrQ5tb+YZHeV8I+YwZTU+p+ZIOZJnr+ZLH2FjC+ZSuOU9M+Z3KOHyL+aCqUGCU
+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
//wfHli2//wtFpn/`;
});
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
define("extractors/cxsecret-font", ["require", "exports", "extractors/cxsecret-table"], function (require, exports, cxsecret_table_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.decodeBase64 = decodeBase64;
    exports.sha1Prefix32 = sha1Prefix32;
    exports.getGlyphTable = getGlyphTable;
    exports.lookupGlyphCodePoint = lookupGlyphCodePoint;
    exports.buildGlyphDecodeMap = buildGlyphDecodeMap;
    exports.replaceByCodePointMap = replaceByCodePointMap;
    /** 表内每条记录：4 字节哈希键 + 2 字节 BMP 码位 */
    const ENTRY_SIZE = 6;
    /** 认定一个内嵌字体是反爬字体所需的最少命中数，避免误伤普通内嵌字体 */
    const MIN_DECODE_HITS = 4;
    /* ------------------------------------------------------------------ base64 */
    const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    const BASE64_INDEX = (() => {
        const index = new Int16Array(128).fill(-1);
        for (let i = 0; i < BASE64_ALPHABET.length; i += 1)
            index[BASE64_ALPHABET.charCodeAt(i)] = i;
        return index;
    })();
    /** 自带 base64 解码，避免依赖宿主环境的 atob（油猴沙箱/Node 表现不一致） */
    function decodeBase64(input) {
        const clean = input.replace(/[^A-Za-z0-9+/]/g, '');
        const length = clean.length;
        const output = new Uint8Array(((length * 3) >> 2) + 3);
        let buffer = 0;
        let bits = 0;
        let offset = 0;
        for (let i = 0; i < length; i += 1) {
            const code = clean.charCodeAt(i);
            const value = code < 128 ? (BASE64_INDEX[code] ?? -1) : -1;
            if (value < 0)
                continue;
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
    function sha1Prefix32(bytes) {
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
            for (let i = 0; i < 16; i += 1)
                words[i] = view.getUint32(block + i * 4);
            for (let i = 16; i < 80; i += 1) {
                const mixed = (words[i - 3] ?? 0) ^ (words[i - 8] ?? 0) ^ (words[i - 14] ?? 0) ^ (words[i - 16] ?? 0);
                words[i] = (mixed << 1) | (mixed >>> 31);
            }
            let a = h0;
            let b = h1;
            let c = h2;
            let d = h3;
            let e = h4;
            for (let i = 0; i < 80; i += 1) {
                let f;
                let k;
                if (i < 20) {
                    f = (b & c) | (~b & d);
                    k = 0x5a827999;
                }
                else if (i < 40) {
                    f = b ^ c ^ d;
                    k = 0x6ed9eba1;
                }
                else if (i < 60) {
                    f = (b & c) | (b & d) | (c & d);
                    k = 0x8f1bbcdc;
                }
                else {
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
    let cachedTable;
    /** 懒解析内置解码表（163 KB base64 → 122 KB 二进制），只做一次 */
    function getGlyphTable() {
        if (cachedTable !== undefined)
            return cachedTable;
        try {
            const compact = cxsecret_table_1.CXSECRET_GLYPH_TABLE_B64.replace(/\s+/g, '');
            cachedTable = compact ? decodeBase64(compact) : null;
        }
        catch {
            cachedTable = null;
        }
        return cachedTable;
    }
    function readKey(table, offset) {
        const a = table[offset] ?? 0;
        const b = table[offset + 1] ?? 0;
        const c = table[offset + 2] ?? 0;
        const d = table[offset + 3] ?? 0;
        return ((a << 24) | (b << 16) | (c << 8) | d) >>> 0;
    }
    /** 表按哈希键升序排列，二分查找；返回真实汉字码位，未命中返回 0 */
    function lookupGlyphCodePoint(table, key) {
        let low = 0;
        let high = Math.floor(table.length / ENTRY_SIZE) - 1;
        while (low <= high) {
            const mid = (low + high) >> 1;
            const base = mid * ENTRY_SIZE;
            const current = readKey(table, base);
            if (current === key)
                return ((table[base + 4] ?? 0) << 8) | (table[base + 5] ?? 0);
            if (current < key)
                low = mid + 1;
            else
                high = mid - 1;
        }
        return 0;
    }
    /* ------------------------------------------------------------------ sfnt 解析 */
    function readTag(view, offset) {
        return String.fromCharCode(view.getUint8(offset), view.getUint8(offset + 1), view.getUint8(offset + 2), view.getUint8(offset + 3));
    }
    function readSfntDirectory(font) {
        const bytes = font instanceof Uint8Array ? font : new Uint8Array(font);
        if (bytes.byteLength < 12)
            return null;
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        // TTC（字体集合）取其第一个字体；页面内嵌字体一般是单个 TTF
        let base = 0;
        if (readTag(view, 0) === 'ttcf') {
            if (bytes.byteLength < 16)
                return null;
            base = view.getUint32(12);
        }
        if (base + 12 > bytes.byteLength)
            return null;
        const count = view.getUint16(base + 4);
        const tables = new Map();
        for (let i = 0; i < count; i += 1) {
            const entry = base + 12 + i * 16;
            if (entry + 16 > bytes.byteLength)
                break;
            tables.set(readTag(view, entry), [view.getUint32(entry + 8), view.getUint32(entry + 12)]);
        }
        return { view, bytes, tables };
    }
    function readLoca(view, offset, numGlyphs, format) {
        const count = numGlyphs + 1;
        const size = format === 0 ? 2 : 4;
        if (offset + count * size > view.byteLength)
            return null;
        const offsets = new Array(count);
        for (let i = 0; i < count; i += 1) {
            offsets[i] = format === 0 ? view.getUint16(offset + i * 2) * 2 : view.getUint32(offset + i * 4);
        }
        return offsets;
    }
    /** 读取 cmap → { 码位: 字形索引 }；支持 format 4（BMP）与 format 12 */
    function readCmap(view, offset) {
        const result = new Map();
        if (offset + 4 > view.byteLength)
            return result;
        const count = view.getUint16(offset + 2);
        const subtables = [];
        for (let i = 0; i < count; i += 1) {
            const entry = offset + 4 + i * 8;
            if (entry + 8 > view.byteLength)
                break;
            const platform = view.getUint16(entry);
            const encoding = view.getUint16(entry + 2);
            const rank = platform === 3 && encoding === 1 ? 0 : platform === 3 && encoding === 10 ? 1 : 2;
            subtables.push({ rank, offset: offset + view.getUint32(entry + 4) });
        }
        subtables.sort((left, right) => left.rank - right.rank);
        for (const subtable of subtables) {
            const start = subtable.offset;
            if (start + 2 > view.byteLength)
                continue;
            const format = view.getUint16(start);
            if (format === 4) {
                const segmentCount = view.getUint16(start + 6) >> 1;
                const endBase = start + 14;
                const startBase = endBase + segmentCount * 2 + 2;
                const deltaBase = startBase + segmentCount * 2;
                const rangeBase = deltaBase + segmentCount * 2;
                if (rangeBase + segmentCount * 2 > view.byteLength)
                    continue;
                for (let i = 0; i < segmentCount; i += 1) {
                    const end = view.getUint16(endBase + i * 2);
                    const first = view.getUint16(startBase + i * 2);
                    if (first === 0xffff)
                        continue;
                    const delta = view.getInt16(deltaBase + i * 2);
                    const rangeOffset = view.getUint16(rangeBase + i * 2);
                    for (let code = first; code <= end; code += 1) {
                        if (code === 0xffff)
                            break;
                        if (result.has(code))
                            continue;
                        let glyph;
                        if (rangeOffset === 0) {
                            glyph = (code + delta) & 0xffff;
                        }
                        else {
                            const pointer = rangeBase + i * 2 + rangeOffset + (code - first) * 2;
                            if (pointer + 2 > view.byteLength)
                                continue;
                            glyph = view.getUint16(pointer);
                            if (glyph !== 0)
                                glyph = (glyph + delta) & 0xffff;
                        }
                        if (glyph !== 0)
                            result.set(code, glyph);
                    }
                }
                continue;
            }
            if (format === 12) {
                const groupCount = view.getUint32(start + 12);
                for (let i = 0; i < groupCount; i += 1) {
                    const group = start + 16 + i * 12;
                    if (group + 12 > view.byteLength)
                        break;
                    const first = view.getUint32(group);
                    const last = view.getUint32(group + 4);
                    const glyph = view.getUint32(group + 8);
                    for (let code = first; code <= last; code += 1) {
                        if (result.has(code))
                            continue;
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
    function buildGlyphDecodeMap(font) {
        const result = new Map();
        const table = getGlyphTable();
        if (!table)
            return result;
        const sfnt = readSfntDirectory(font);
        if (!sfnt)
            return result;
        const { view, bytes, tables } = sfnt;
        const glyf = tables.get('glyf');
        const loca = tables.get('loca');
        const head = tables.get('head');
        const maxp = tables.get('maxp');
        const cmapTable = tables.get('cmap');
        // 无 glyf 说明是 CFF 轮廓，本机制不适用
        if (!glyf || !loca || !head || !maxp || !cmapTable)
            return result;
        if (head[0] + 52 > bytes.byteLength)
            return result;
        const numGlyphs = view.getUint16(maxp[0] + 4);
        const offsets = readLoca(view, loca[0], numGlyphs, view.getInt16(head[0] + 50));
        if (!offsets)
            return result;
        const cmap = readCmap(view, cmapTable[0]);
        if (cmap.size === 0)
            return result;
        const glyfOffset = glyf[0];
        let hits = 0;
        for (const [codePoint, glyph] of cmap) {
            if (glyph + 1 >= offsets.length)
                continue;
            const start = offsets[glyph] ?? 0;
            const end = offsets[glyph + 1] ?? 0;
            if (end <= start)
                continue;
            const decoded = lookupGlyphCodePoint(table, sha1Prefix32(bytes.subarray(glyfOffset + start, glyfOffset + end)));
            // 解码结果与码位相同说明该字形本就是它自己，不属于混淆，跳过
            if (decoded === 0 || decoded === codePoint)
                continue;
            result.set(codePoint, String.fromCharCode(decoded));
            hits += 1;
        }
        // 命中过少则视为误判，避免把普通内嵌字体当反爬字体处理
        return hits >= MIN_DECODE_HITS ? result : new Map();
    }
    /** 按映射逐字替换；无任何替换时原样返回，避免多余字符串拼接 */
    function replaceByCodePointMap(text, map) {
        if (!text || map.size === 0)
            return text;
        let output = '';
        let changed = false;
        for (const char of text) {
            const decoded = map.get(char.codePointAt(0) ?? 0);
            if (decoded === undefined) {
                output += char;
            }
            else {
                output += decoded;
                changed = true;
            }
        }
        return changed ? output : text;
    }
});
/**
 * font-cxsecret 字体反爬解码 —— DOM 接入层
 *
 * 职责：
 *   1. 从文档里找出内联的 @font-face 字体（学习通把字体以 base64 内联在 <style> 中）
 *   2. 用纯逻辑层算出「混淆码位 → 真实汉字」映射并缓存
 *   3. 按作用域把文本还原：只处理带 cxsecret 类名（或其祖先带该类名）的节点
 *
 * 为什么要按作用域过滤：混淆只作用于页面标记为 font-cxsecret 的子树，子树外的
 * 同码位文字属于正常内容，若一并替换会把它改坏。若页面里找不到任何该类名标记
 * （例如对方改名），则退化为整页替换——此时不替换只会输出乱码，替换更有利。
 */
define("extractors/cxsecret-decoder", ["require", "exports", "extractors/cxsecret-font"], function (require, exports, cxsecret_font_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.parentElementOf = parentElementOf;
    exports.getCxSecretDecoder = getCxSecretDecoder;
    exports.clearCxSecretDecoderCache = clearCxSecretDecoderCache;
    exports.isCxSecretScope = isCxSecretScope;
    exports.decodeCxSecretText = decodeCxSecretText;
    /** 与字体家族名无关的作用域判定：类名里含 cxsecret 即视为混淆子树 */
    const SCOPE_MARKER_SELECTOR = '[class*="cxsecret"]';
    const CXSECRET_FAMILY = /cxsecret/i;
    /** @font-face 块内的家族名与字体数据 */
    const FONT_FACE_PATTERN = /@font-face\s*\{([^}]*)\}/gi;
    const FAMILY_PATTERN = /font-family\s*:\s*(['"]?)([^;'"]+)\1/i;
    const SOURCE_URL_PATTERN = /url\(\s*(['"]?)(data:[^'")]+)\1\s*\)/i;
    const BASE64_PAYLOAD_PATTERN = /;base64,([A-Za-z0-9+/=\s]+)$/i;
    const decoderCache = new WeakMap();
    const scopeCache = new WeakMap();
    function resolveDocument(doc) {
        if (doc)
            return doc;
        return typeof document === 'undefined' ? null : document;
    }
    /** 取父元素：用 parentNode + nodeType 判定，比 parentElement 在各 DOM 实现上更可靠 */
    function parentElementOf(node) {
        const parent = node?.parentNode ?? null;
        return parent && parent.nodeType === Node.ELEMENT_NODE ? parent : null;
    }
    function tryQuery(root, selector) {
        try {
            return Array.from(root.querySelectorAll(selector));
        }
        catch {
            return [];
        }
    }
    function parseFontFaceBlock(block) {
        const family = block.match(FAMILY_PATTERN)?.[2]?.trim() ?? '';
        const url = block.match(SOURCE_URL_PATTERN)?.[2] ?? '';
        const base64 = url.match(BASE64_PAYLOAD_PATTERN)?.[1]?.replace(/\s+/g, '') ?? '';
        if (!base64)
            return null;
        return { family, base64 };
    }
    /** 收集文档里所有内联 data: 字体的 @font-face */
    function collectFontSources(doc) {
        const sources = [];
        const push = (css) => {
            if (!css || !css.includes('@font-face'))
                return;
            const pattern = new RegExp(FONT_FACE_PATTERN.source, 'gi');
            for (const match of css.matchAll(pattern)) {
                const source = parseFontFaceBlock(match[1] ?? '');
                if (source)
                    sources.push(source);
            }
        };
        for (const style of tryQuery(doc, 'style'))
            push(style.textContent ?? '');
        // 同源外链样式表：cssRules 不可读时（跨域）静默跳过
        try {
            for (const sheet of Array.from(doc.styleSheets ?? [])) {
                let rules;
                try {
                    rules = sheet.cssRules;
                }
                catch {
                    continue;
                }
                for (const rule of Array.from(rules ?? [])) {
                    const cssText = rule.cssText ?? '';
                    if (cssText.includes('@font-face'))
                        push(cssText);
                }
            }
        }
        catch {
            // styleSheets 在部分文档上不可用，忽略
        }
        return sources;
    }
    function buildMapFromSources(sources) {
        const map = new Map();
        for (const source of sources) {
            try {
                const partial = (0, cxsecret_font_1.buildGlyphDecodeMap)((0, cxsecret_font_1.decodeBase64)(source.base64));
                for (const [codePoint, char] of partial)
                    map.set(codePoint, char);
            }
            catch {
                // 单个字体解析失败不影响其他候选
            }
        }
        return map;
    }
    function createDecoder(doc) {
        const sources = collectFontSources(doc);
        if (sources.length === 0)
            return null;
        // 先按家族名筛选，命中不了再退回尝试全部内联字体
        const preferred = sources.filter((source) => CXSECRET_FAMILY.test(source.family));
        let map = preferred.length > 0 ? buildMapFromSources(preferred) : new Map();
        if (map.size === 0)
            map = buildMapFromSources(sources);
        if (map.size === 0)
            return null;
        return {
            size: map.size,
            scoped: tryQuery(doc, SCOPE_MARKER_SELECTOR).length > 0,
            decode: (text) => (0, cxsecret_font_1.replaceByCodePointMap)(text, map),
        };
    }
    /** 取文档的解码器；无内联反爬字体时返回 null */
    function getCxSecretDecoder(doc) {
        const target = resolveDocument(doc);
        if (!target)
            return null;
        if (decoderCache.has(target))
            return decoderCache.get(target) ?? null;
        let decoder = null;
        try {
            decoder = createDecoder(target);
        }
        catch {
            decoder = null;
        }
        decoderCache.set(target, decoder);
        return decoder;
    }
    /** 清空缓存：同一文档内重载了内容（字体可能变化）后调用 */
    function clearCxSecretDecoderCache(doc) {
        const target = resolveDocument(doc);
        if (target)
            decoderCache.delete(target);
    }
    /**
     * 判定元素是否处于混淆作用域内（自身或任一祖先带 cxsecret 类名）。
     * 结果按元素缓存，避免逐文本节点重复向上遍历。
     */
    function isCxSecretScope(element) {
        if (!element)
            return false;
        const cached = scopeCache.get(element);
        if (cached !== undefined)
            return cached;
        let result = false;
        let node = element;
        while (node) {
            const className = node.getAttribute('class') ?? '';
            if (className && CXSECRET_FAMILY.test(className)) {
                result = true;
                break;
            }
            node = parentElementOf(node);
        }
        scopeCache.set(element, result);
        return result;
    }
    /**
     * 还原元素文本里的混淆汉字。
     * @param text  已读出的原始文本
     * @param scope 文本所属元素（用于判定作用域）
     */
    function decodeCxSecretText(text, scope, doc) {
        if (!text)
            return text;
        const decoder = getCxSecretDecoder(doc ?? scope?.ownerDocument ?? null);
        if (!decoder)
            return text;
        if (decoder.scoped && !isCxSecretScope(scope))
            return text;
        return decoder.decode(text);
    }
});
define("extractors/rich-content", ["require", "exports", "utils/text", "extractors/cxsecret-decoder"], function (require, exports, text_1, cxsecret_decoder_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.normalizeRichContent = normalizeRichContent;
    exports.resolveImageUrl = resolveImageUrl;
    exports.extractBackgroundImages = extractBackgroundImages;
    exports.extractRichContent = extractRichContent;
    exports.richContentToText = richContentToText;
    exports.isRichContentEmpty = isRichContentEmpty;
    exports.stripOptionPrefix = stripOptionPrefix;
    exports.stripQuestionPrefix = stripQuestionPrefix;
    exports.stripAnswerLabel = stripAnswerLabel;
    exports.joinRichContents = joinRichContents;
    exports.textContent = textContent;
    const BLOCK_TAGS = new Set([
        'P',
        'DIV',
        'LI',
        'DD',
        'DT',
        'TR',
        'TABLE',
        'SECTION',
        'ARTICLE',
        'BLOCKQUOTE',
        'H1',
        'H2',
        'H3',
        'H4',
        'H5',
        'H6',
    ]);
    function sameTextStyle(left, right) {
        return (Boolean(left.bold) === Boolean(right.bold) &&
            Boolean(left.italic) === Boolean(right.italic) &&
            Boolean(left.subScript) === Boolean(right.subScript) &&
            Boolean(left.superScript) === Boolean(right.superScript));
    }
    function normalizeRichContent(parts) {
        const output = [];
        for (const part of parts) {
            if (part.type === 'text') {
                const text = part.text.replace(/\u00a0/g, ' ').replace(/[\t\r\f ]+/g, ' ');
                if (!text)
                    continue;
                const previous = output[output.length - 1];
                if (previous?.type === 'text' && sameTextStyle(previous, part)) {
                    output[output.length - 1] = { ...previous, text: previous.text + text };
                }
                else {
                    output.push({ ...part, text });
                }
                continue;
            }
            if (part.type === 'image') {
                if (part.url)
                    output.push(part);
                continue;
            }
            if (output[output.length - 1]?.type !== 'break')
                output.push(part);
        }
        // 修剪首尾的换行与纯空白文本：尾部空白会挡住相邻换行，导致 Word 导出时题目与选项之间出现多余空行
        const isTrimmable = (part) => {
            if (!part)
                return false;
            return part.type === 'break' || (part.type === 'text' && !part.text.trim());
        };
        while (isTrimmable(output[0]))
            output.shift();
        while (isTrimmable(output[output.length - 1]))
            output.pop();
        return output;
    }
    function resolveUrl(value, baseUrl) {
        if (!value || value === 'about:blank')
            return '';
        try {
            return new URL(value, baseUrl).href;
        }
        catch {
            return value;
        }
    }
    function resolveImageUrl(image) {
        const attributes = [
            'data-original',
            'data-src',
            'data-lazy-src',
            'origin-src',
            'src',
            'fileid',
        ];
        const baseUrl = image.ownerDocument.baseURI;
        if (image.currentSrc)
            return resolveUrl(image.currentSrc, baseUrl);
        for (const attribute of attributes) {
            const value = image.getAttribute(attribute)?.trim();
            if (value)
                return resolveUrl(value, baseUrl);
        }
        return resolveUrl(image.src, baseUrl);
    }
    function extractBackgroundImages(element) {
        // 不能按 instanceof HTMLElement 判定：顶层脚本跨层读取同源 iframe 文档时，
        // 节点的构造器属于那个文档所在的 window，instanceof 会一律判 false；
        // 改为按能力判断——能取到 style 的元素才可能有背景图
        const style = element.style;
        if (!style)
            return [];
        let background = style.backgroundImage;
        try {
            if (!background || background === 'none') {
                background = element.ownerDocument.defaultView?.getComputedStyle(element).backgroundImage ?? '';
            }
        }
        catch {
            // Computed styles may be unavailable for detached or cross-origin documents.
        }
        if (!background || background === 'none')
            return [];
        const images = [];
        const pattern = /url\(\s*(["']?)(.*?)\1\s*\)/giu;
        for (const match of background.matchAll(pattern)) {
            const value = match[2]?.trim();
            if (!value)
                continue;
            const url = resolveUrl(value, element.ownerDocument.baseURI);
            if (url)
                images.push({ type: 'image', url, alt: element.getAttribute('aria-label') ?? '' });
        }
        return images;
    }
    const EMPTY_STYLE = {
        bold: false,
        italic: false,
        subScript: false,
        superScript: false,
    };
    /**
     * 节点是否为元素。
     *
     * 这里不能用 `node instanceof Element`：脚本在顶层窗口运行时会把同源 iframe 的文档一并
     * 遍历（`collectAccessibleDocuments`），而那些节点的构造器属于它们自己的 window，
     * `instanceof` 会一律判 false —— 后果是整棵子树都读不出内容。题干若恰好落在容器的
     * 直接文本节点上还能读到，选项文字却都嵌在 `<a>` / `<label>` / `<span>` 里，就会一条都
     * 读不到，表现为「提取到题干、提取不到选项」（且题干里嵌套的题型标记会一起丢失）。
     * nodeType 是各文档统一的常量，与节点属于哪个 window 无关，用它判定即可。
     */
    function isElementNode(node) {
        return node.nodeType === Node.ELEMENT_NODE;
    }
    function extractRichContent(element) {
        if (!element)
            return [];
        const parts = [];
        const walk = (node, style) => {
            if (node.nodeType === Node.TEXT_NODE) {
                // 学习通对题干/选项做了字体反爬，取到的是混淆码位，这里按作用域还原
                parts.push({
                    type: 'text',
                    text: (0, cxsecret_decoder_1.decodeCxSecretText)(node.nodeValue ?? '', (0, cxsecret_decoder_1.parentElementOf)(node)),
                    ...style,
                });
                return;
            }
            if (!isElementNode(node))
                return;
            const tag = node.tagName.toUpperCase();
            if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE'].includes(tag))
                return;
            if (tag === 'IMG') {
                const url = resolveImageUrl(node);
                if (url)
                    parts.push({ type: 'image', url, alt: node.getAttribute('alt') ?? '' });
                return;
            }
            if (tag === 'BR') {
                parts.push({ type: 'break' });
                return;
            }
            const childStyle = {
                bold: style.bold || tag === 'STRONG' || tag === 'B',
                italic: style.italic || tag === 'EM' || tag === 'I',
                subScript: style.subScript || tag === 'SUB',
                superScript: style.superScript || tag === 'SUP',
            };
            const lengthBefore = parts.length;
            node.childNodes.forEach((child) => walk(child, childStyle));
            if (BLOCK_TAGS.has(tag) && parts.length > lengthBefore)
                parts.push({ type: 'break' });
        };
        element.childNodes.forEach((child) => walk(child, EMPTY_STYLE));
        const existingImageUrls = new Set(parts.filter((part) => part.type === 'image').map((part) => part.url));
        for (const image of extractBackgroundImages(element)) {
            if (!existingImageUrls.has(image.url))
                parts.push(image);
        }
        let normalized = normalizeRichContent(parts);
        if (normalized.length === 0) {
            const fallback = [
                element.getAttribute('aria-label'),
                element.getAttribute('title'),
                element.getAttribute('data-content'),
                element.getAttribute('data-text'),
                element.getAttribute('data-value'),
            ].find((value) => Boolean(value?.trim()));
            // 选项文字在真实页面上同时存在于 aria-label，同样需要还原
            if (fallback)
                normalized = [{ type: 'text', text: (0, cxsecret_decoder_1.decodeCxSecretText)(fallback, element) }];
        }
        return normalized;
    }
    function richContentToText(content, imageFormatter = () => '') {
        let output = '';
        for (const part of content) {
            if (part.type === 'text')
                output += part.text;
            else if (part.type === 'image')
                output += imageFormatter(part.url, part.alt);
            else
                output += '\n';
        }
        return (0, text_1.normalizeWhitespace)(output);
    }
    function isRichContentEmpty(content) {
        return !content.some((part) => {
            return part.type === 'image' || (part.type === 'text' && Boolean(part.text.trim()));
        });
    }
    function cloneContent(content) {
        return content.map((part) => ({ ...part }));
    }
    function stripOptionPrefix(content) {
        const output = cloneContent(content);
        let key = '';
        for (let index = 0; index < output.length; index += 1) {
            const part = output[index];
            if (part?.type !== 'text' || !part.text.trim())
                continue;
            const match = part.text.match(/^\s*([A-Z])\s*[.、．:：]?\s*/iu);
            if (match?.[1]) {
                key = match[1].toUpperCase();
                output[index] = { ...part, text: part.text.slice(match[0].length) };
            }
            break;
        }
        return { key, content: normalizeRichContent(output) };
    }
    function stripQuestionPrefix(content) {
        const output = cloneContent(content);
        let numberHandled = false;
        let typeHandled = false;
        for (let index = 0; index < output.length; index += 1) {
            const part = output[index];
            if (part?.type !== 'text')
                continue;
            let text = part.text.replace(/\s*[（(]\s*\d+(?:\.\d+)?\s*分?\s*[）)]\s*$/u, '');
            if (!numberHandled) {
                text = text.replace(/^\s*\d+\s*[.、．]\s*/u, '');
                if (text.trim())
                    numberHandled = true;
            }
            if (!typeHandled) {
                text = text.replace(/^\s*(?:[（(【[]\s*)?(?:单选题|单项选择题|多选题|多项选择题|填空题|判断题|简答题|论述题|问答题|选择题)(?:\s*[）)】\]])?\s*/u, '');
                if (text.trim())
                    typeHandled = true;
            }
            output[index] = { ...part, text };
        }
        return normalizeRichContent(output);
    }
    function stripAnswerLabel(content) {
        const output = cloneContent(content);
        for (let index = 0; index < output.length; index += 1) {
            const part = output[index];
            if (part?.type !== 'text' || !part.text.trim())
                continue;
            output[index] = {
                ...part,
                text: part.text.replace(/^\s*(?:正确答案|参考答案|答案|我的答案|你的答案|学生答案|答案解析|解析)\s*[:：]?\s*/u, ''),
            };
            break;
        }
        return normalizeRichContent(output);
    }
    function joinRichContents(contents, separator = '；') {
        const output = [];
        for (const content of contents) {
            if (isRichContentEmpty(content))
                continue;
            if (output.length > 0)
                output.push({ type: 'text', text: separator });
            output.push(...content.map((part) => ({ ...part })));
        }
        return normalizeRichContent(output);
    }
    function textContent(value) {
        const normalized = value.trim();
        return normalized ? [{ type: 'text', text: normalized }] : [];
    }
});
define("extractors/answer-comparison", ["require", "exports", "utils/text", "extractors/rich-content"], function (require, exports, text_2, rich_content_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.compareAnswers = compareAnswers;
    exports.hasExplicitWrongMarker = hasExplicitWrongMarker;
    function compareAnswers(type, userAnswer, correctAnswer) {
        if (type === 'short-answer')
            return false;
        const user = (0, text_2.normalizeAnswer)((0, rich_content_1.richContentToText)(userAnswer));
        const correct = (0, text_2.normalizeAnswer)((0, rich_content_1.richContentToText)(correctAnswer));
        if (!user || !correct)
            return false;
        if (type === 'multiple-choice') {
            return [...user].sort().join('') !== [...correct].sort().join('');
        }
        return user !== correct;
    }
    function hasExplicitWrongMarker(container) {
        return Boolean(container.querySelector([
            '.colorRed',
            '.wrong',
            '.answer-wrong',
            '.is-wrong',
            '[data-correct="false"]',
            '[data-result="wrong"]',
            // 章节测验「已完成 / 已批阅」视图不输出「正确答案」文本，只在每题下用图标表示批阅结果：
            // span.marking_dui 批阅正确、span.marking_cuo 批阅错误、span.marking_bandui 部分正确。
            // 这类页面没有可比较的答案文本，只能靠该标记判定错题，否则错题汇总会整页漏掉。
            // 部分正确同样需要复习，因此一并计入错题。
            '.marking_cuo',
            '.marking_bandui',
        ].join(',')));
    }
});
define("extractors/option-parser", ["require", "exports", "utils/text", "extractors/rich-content"], function (require, exports, text_3, rich_content_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.FALLBACK_OPTION_SELECTORS = void 0;
    exports.parseLegacyOption = parseLegacyOption;
    exports.parseAnswerBackgroundOption = parseAnswerBackgroundOption;
    exports.parseOptions = parseOptions;
    exports.optionKeyFromContainer = optionKeyFromContainer;
    exports.answersFromCheckedInputs = answersFromCheckedInputs;
    function parseLegacyOption(element, index) {
        const parsed = (0, rich_content_2.stripOptionPrefix)((0, rich_content_2.extractRichContent)(element));
        const key = parsed.key || String.fromCharCode(65 + index);
        if ((0, rich_content_2.isRichContentEmpty)(parsed.content))
            return null;
        return { key, content: parsed.content };
    }
    function parseAnswerBackgroundOption(element, index) {
        const keyElement = element.querySelector('.num_option, .option-letter, .option-index');
        const contentElement = element.querySelector('.answer_p, .option-content, .option-text') ?? element;
        const rawKey = (0, text_3.normalizeInlineWhitespace)(keyElement?.getAttribute('data') ??
            keyElement?.getAttribute('data-option') ??
            keyElement?.textContent ??
            String.fromCharCode(65 + index));
        const key = rawKey.match(/[A-Z]/iu)?.[0]?.toUpperCase() ?? '';
        const parsed = (0, rich_content_2.stripOptionPrefix)((0, rich_content_2.extractRichContent)(contentElement));
        if ((0, rich_content_2.isRichContentEmpty)(parsed.content))
            return null;
        return { key: key || parsed.key || String.fromCharCode(65 + index), content: parsed.content };
    }
    /**
     * 选项容器兜底选择器：在各提取器自带的候选之后追加。
     *
     * 学习通 2026 版答题页的每个选项是 `<li class="font-cxsecret before-after" role="radio|checkbox"
     * onclick="addChoice(this)">`，选项文字在 `a.after` 里。类名随版本变化，但 `role` 与
     * `onclick` 这两个语义标记极稳定，因此单独放在最后兜底：正常页面走自带候选、行为不变，
     * 只有在类名全部失效（表现为「只提取到题干、提取不到选项」）时才由这里接手。
     */
    exports.FALLBACK_OPTION_SELECTORS = [
        '.Zy_ulTop li',
        '.qtDetail li',
        'li[onclick*="addChoice"]',
        'li[role="radio"]',
        'li[role="checkbox"]',
        // 已批阅视图的选项是 <li class="clearfix" role="option"><i class="fl">A、</i><a class="fl">…</a></li>，
        // 该类名同样随版本变化，只有 role="option" 稳定；它也是页面里最常见的通用 role，
        // 因此放在最末，仅当前面所有候选都落空时才接手。
        'li[role="option"]',
    ];
    function parseOptions(container, selectors) {
        for (const selector of [...selectors, ...exports.FALLBACK_OPTION_SELECTORS]) {
            const elements = Array.from(container.querySelectorAll(selector));
            if (elements.length === 0)
                continue;
            const options = elements
                .map((element, index) => {
                return element.matches('.answerBg, [class*="answerBg"]')
                    ? parseAnswerBackgroundOption(element, index)
                    : parseLegacyOption(element, index);
            })
                .filter((option) => option !== null);
            if (options.length > 0)
                return options.sort((left, right) => left.key.localeCompare(right.key));
        }
        return [];
    }
    function optionKeyFromContainer(container, index) {
        const direct = [
            container.getAttribute('data'),
            container.getAttribute('data-option'),
            container.getAttribute('data-key'),
            container.querySelector('.num_option, .option-letter, .mark_letter')?.textContent,
        ].find((value) => Boolean(value?.trim()));
        if (direct) {
            const normalized = (0, text_3.normalizeInlineWhitespace)(direct).match(/[A-Z]/iu)?.[0];
            if (normalized)
                return normalized.toUpperCase();
        }
        return String.fromCharCode(65 + index);
    }
    function answersFromCheckedInputs(container) {
        const selected = Array.from(container.querySelectorAll('input[type="radio"]:checked, input[type="checkbox"]:checked'));
        if (selected.length === 0)
            return [];
        const keys = selected.map((input, index) => {
            const optionContainer = input.closest('.answerBg, li, .option, label') ?? input;
            return optionKeyFromContainer(optionContainer, index);
        });
        return [{ type: 'text', text: keys.join('') }];
    }
});
define("extractors/common-answer", ["require", "exports", "utils/dom", "extractors/rich-content", "extractors/option-parser"], function (require, exports, dom_1, rich_content_3, option_parser_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.extractCorrectAnswer = extractCorrectAnswer;
    exports.extractUserAnswer = extractUserAnswer;
    exports.extractAnalysis = extractAnalysis;
    exports.inferCorrectAnswerFromGrading = inferCorrectAnswerFromGrading;
    const CORRECT_SELECTORS = [
        '.mark_answer .mark_key .colorGreen .stuAnswerContent',
        '.mark_answer .mark_key .colorGreen',
        '.newAnswerBx .correctAnswerBx .answerCon',
        '.newAnswerBx .correctAnswer .answerCon',
        '.correctAnswerBx .answerCon',
        '.correctAnswerBx .correctAnswer',
        '.correctAnswerContent',
        '.correct-answer',
        '.rightAnswer',
        '[data-role="correct-answer"]',
    ];
    const USER_SELECTORS = [
        '.mark_answer .mark_key .colorDeep .stuAnswerContent',
        '.mark_answer .mark_key .colorDeep',
        '.newAnswerBx .myAnswerBx .answerCon',
        // 章节测验「已批阅」视图一次作答会渲染多组答案，此时用 myAllAnswerBx 包裹
        '.newAnswerBx .myAllAnswerBx .myAnswerBx',
        '.myAnswerBx .answerCon',
        '.myAnswer .answerCon',
        '.my-answer',
        '[data-role="user-answer"]',
    ];
    const ANALYSIS_SELECTORS = [
        '.newAnswerBx .answerKeyBx .answerCon',
        '.answerKeyBx .answerCon',
        '.newAnswerBx .analysisBx .answerCon',
        '.analysisBx .answerCon',
        '.mark_answer .mark_analysis',
        '.mark_answer .analysis',
        '.answerAnalysis',
        '.analysisContent',
        '.answer-analysis',
        '[data-role="analysis"]',
    ];
    function extractFirstContent(container, selectors) {
        const element = (0, dom_1.firstMatch)(container, selectors);
        return (0, rich_content_3.stripAnswerLabel)((0, rich_content_3.extractRichContent)(element));
    }
    function extractFillAnswer(container, selector) {
        const elements = (0, dom_1.allMatches)(container, [selector]);
        const contents = elements.map((element) => (0, rich_content_3.stripAnswerLabel)((0, rich_content_3.extractRichContent)(element)));
        return (0, rich_content_3.joinRichContents)(contents);
    }
    function extractCorrectAnswer(container) {
        const fill = extractFillAnswer(container, '.mark_answer .mark_fill.colorGreen dd');
        if (!(0, rich_content_3.isRichContentEmpty)(fill))
            return fill;
        return extractFirstContent(container, CORRECT_SELECTORS);
    }
    function extractUserAnswer(container) {
        const fill = extractFillAnswer(container, '.mark_answer .mark_fill .colorDeep, .mark_answer .mark_fill dd.colorDeep');
        if (!(0, rich_content_3.isRichContentEmpty)(fill))
            return fill;
        const direct = extractFirstContent(container, USER_SELECTORS);
        if (!(0, rich_content_3.isRichContentEmpty)(direct))
            return direct;
        const checked = (0, option_parser_1.answersFromCheckedInputs)(container);
        if (!(0, rich_content_3.isRichContentEmpty)(checked))
            return checked;
        const textInputs = Array.from(container.querySelectorAll('input[type="text"], input:not([type]), .blankInput, [data-role="blank-input"]'))
            .map((input) => input.value.trim())
            .filter(Boolean);
        if (textInputs.length > 0)
            return (0, rich_content_3.textContent)(textInputs.join('；'));
        const textarea = container.querySelector('textarea');
        if (textarea?.value.trim())
            return (0, rich_content_3.textContent)(textarea.value);
        const editable = container.querySelector('[contenteditable="true"]');
        if (editable?.innerText.trim())
            return (0, rich_content_3.textContent)(editable.innerText);
        return [];
    }
    function extractAnalysis(container) {
        return extractFirstContent(container, ANALYSIS_SELECTORS);
    }
    /**
     * 从「教师批阅结果」推断正确答案。
     *
     * 章节测验「已完成 / 已批阅」视图只在每题下用图标给出批阅结果（`.marking_dui` 正确 /
     * `.marking_cuo` 错误 / `.marking_bandui` 部分正确），**不输出「正确答案」文本**。
     * 批阅正确说明该题我的答案与标准答案一致，此时用我的答案回填，否则导出的「答案汇总」
     * 会整页变成「（未找到答案）」，期末复习没有意义。
     *
     * 只在拿不到正确答案文本时介入（调用方负责判断），页面本身有正确答案时行为不变；
     * 批阅错误 / 部分正确时不回填，避免把错的答案当成标准答案。
     */
    function inferCorrectAnswerFromGrading(container, userAnswer) {
        if ((0, rich_content_3.isRichContentEmpty)(userAnswer))
            return [];
        if (!container.querySelector('.marking_dui'))
            return [];
        return userAnswer;
    }
});
define("utils/hash", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.stableHash = stableHash;
    function stableHash(value) {
        let hash = 0x811c9dc5;
        for (let index = 0; index < value.length; index += 1) {
            hash ^= value.charCodeAt(index);
            hash = Math.imul(hash, 0x01000193);
        }
        return (hash >>> 0).toString(36);
    }
});
define("extractors/question-factory", ["require", "exports", "utils/hash", "extractors/answer-comparison", "extractors/rich-content"], function (require, exports, hash_1, answer_comparison_1, rich_content_4) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.createQuestion = createQuestion;
    function createQuestion(draft) {
        if ((0, rich_content_4.isRichContentEmpty)(draft.stem))
            return null;
        const options = draft.options ?? [];
        const correctAnswer = draft.correctAnswer ?? [];
        const userAnswer = draft.userAnswer ?? [];
        const analysis = draft.analysis ?? [];
        const fingerprint = [
            draft.type,
            draft.number ?? '',
            (0, rich_content_4.richContentToText)(draft.stem, (url) => url),
            ...options.map((option) => `${option.key}:${(0, rich_content_4.richContentToText)(option.content, (url) => url)}`),
        ].join('|');
        return {
            id: `${draft.source.extractor}-${(0, hash_1.stableHash)(fingerprint)}`,
            number: draft.number,
            type: draft.type,
            typeMeta: draft.typeMeta,
            stem: draft.stem,
            options,
            correctAnswer,
            userAnswer,
            analysis,
            isWrong: Boolean(draft.explicitWrong) || (0, answer_comparison_1.compareAnswers)(draft.type, userAnswer, correctAnswer),
            source: draft.source,
            chapterId: draft.chapterId,
            chapterTitle: draft.chapterTitle,
        };
    }
});
define("extractors/question-stem", ["require", "exports", "utils/dom", "extractors/rich-content"], function (require, exports, dom_2, rich_content_5) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.STEM_TYPE_MARKER_SELECTORS = exports.STEM_CONTAINER_SELECTORS = void 0;
    exports.resolveQuestionStem = resolveQuestionStem;
    /**
     * 题干容器候选选择器：按「当前版本 → 历史版本」排列，命中即用。
     *
     * 学习通改版导致题干丢失时，优先在这里追加新类名，无需改动任何提取器的流程代码。
     * 列表内容为四个提取器原有选择器的并集，保证既有页面行为不变。
     */
    exports.STEM_CONTAINER_SELECTORS = [
        '[data-role="stem"]',
        '.question-stem',
        '.questionStem',
        '.question-content',
        '.qtContent',
        '.mark_name',
        '.fontLabel', // 2026 版答题页：.Zy_TItle > div.fontLabel
        '.subject-title',
        '.subject',
    ];
    /**
     * 题型标记选择器：题干与「【单选题】」同处一个元素，是超星最稳定的结构特征。
     * 用于「已知类名全部失效」时反推题干容器。
     */
    exports.STEM_TYPE_MARKER_SELECTORS = [
        '.newZy_TItle',
        '[class*="TestType"]',
        '[class*="ztType"]',
    ];
    /**
     * 非题干节点：选项列表、作答控件、答案区、题型标记。
     * 仅在「剔除式兜底」路径使用（容器本身混有选项时才需要裁剪）。
     */
    const STEM_EXCLUDE_SELECTOR = [
        'script',
        'style',
        'ul',
        'ol',
        'select',
        'textarea',
        '.num_option',
        '.mark_answer',
        '.ans-job-icon',
        '.newZy_TItle',
        '[class*="TestType"]',
        '[class*="ztType"]',
    ].join(',');
    /** 题号候选元素：形如 <i class="fl">1</i> */
    const INDEX_CANDIDATE_SELECTOR = 'i,em,.fl,[class*="index"],[class*="number"]';
    /** 纯题号文本：不含「.」「、」等分隔符，避免误伤题干中自带序号的句子 */
    const BARE_INDEX = /^\s*\d{1,3}\s*$/u;
    const EMPTY_STEM = { content: [], rawText: '' };
    /** 元素是否含有可作题干的内容（文本或图片） */
    function hasStemContent(element) {
        if (element.querySelector('img') !== null)
            return true;
        return Boolean((element.textContent ?? '').replace(/\u00a0/g, ' ').trim());
    }
    /** 复制节点并剔除题号、题型标记、选项列表等非题干节点 */
    function pruneStemElement(element) {
        const clone = element.cloneNode(true);
        clone.querySelectorAll(STEM_EXCLUDE_SELECTOR).forEach((node) => node.remove());
        clone.querySelectorAll(INDEX_CANDIDATE_SELECTOR).forEach((node) => {
            if (BARE_INDEX.test(node.textContent ?? ''))
                node.remove();
        });
        return clone;
    }
    /** 读取容器内容并剥离题号/题型前缀；prune 为真时先裁剪非题干节点 */
    function readStem(element, prune) {
        const source = prune ? pruneStemElement(element) : element;
        return {
            content: (0, rich_content_5.stripQuestionPrefix)((0, rich_content_5.extractRichContent)(source)),
            rawText: (0, dom_2.textOf)(source),
        };
    }
    /**
     * 分层定位题干，任一层取到非空内容即返回：
     *
     * 1. 已知题干类名（当前版本 → 历史版本）
     * 2. 题型标记的父元素（题干容器改名也不影响这条结构特征）
     * 3. 层层兜底：剔除题号与选项后，取标题容器 / 整题容器的剩余内容
     *
     * 题干为空会让 createQuestion 判该题无效并整题丢弃，因此这里必须尽力给出结果。
     */
    function resolveQuestionStem(container, title) {
        const scope = title ?? container;
        // 第一层：已知题干容器。不裁剪，以保留懒加载图片已解析出的 currentSrc
        const known = (0, dom_2.firstMatch)(scope, exports.STEM_CONTAINER_SELECTORS) ??
            (0, dom_2.firstMatch)(container, exports.STEM_CONTAINER_SELECTORS);
        if (known && hasStemContent(known)) {
            const stem = readStem(known, false);
            if (!(0, rich_content_5.isRichContentEmpty)(stem.content))
                return stem;
        }
        // 第二层：题型标记的父元素 —— 题干与「【单选题】」同级
        const marker = (0, dom_2.firstMatch)(scope, exports.STEM_TYPE_MARKER_SELECTORS) ??
            (0, dom_2.firstMatch)(container, exports.STEM_TYPE_MARKER_SELECTORS);
        const markerParent = marker?.parentElement ?? null;
        if (markerParent && markerParent !== scope) {
            const stem = readStem(markerParent, true);
            if (!(0, rich_content_5.isRichContentEmpty)(stem.content))
                return stem;
        }
        // 第三层：标题容器本身，剔除题号、选项列表与答案区后的剩余内容
        const holders = scope === container ? [scope] : [scope, container];
        for (const holder of holders) {
            if (!hasStemContent(holder))
                continue;
            const stem = readStem(holder, true);
            if (!(0, rich_content_5.isRichContentEmpty)(stem.content))
                return stem;
        }
        return EMPTY_STEM;
    }
});
define("extractors/question-type", ["require", "exports", "utils/text"], function (require, exports, text_4) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.detectQuestionType = detectQuestionType;
    exports.inferQuestionType = inferQuestionType;
    const LEGACY_NUMERIC_TYPES = {
        '0': 'single-choice',
        '1': 'multiple-choice',
        '2': 'fill-blank',
        '3': 'true-false',
        '4': 'short-answer',
    };
    const TYPE_PATTERNS = [
        [/(?:多选(?:题)?|多项选择(?:题)?|不定项选择(?:题)?|multiple\s*choice)/iu, 'multiple-choice'],
        [/(?:单选(?:题)?|单项选择(?:题)?|选择题|single\s*choice)/iu, 'single-choice'],
        [/(?:填空(?:题)?|完形填空|fill(?:ing)?\s*(?:in\s*)?blank)/iu, 'fill-blank'],
        [/(?:判断(?:题)?|是非(?:题)?|true\s*or\s*false)/iu, 'true-false'],
        [
            /(?:简答(?:题)?|问答(?:题)?|论述(?:题)?|计算(?:题)?|名词解释|材料分析题|主观题|essay|short\s*answer)/iu,
            'short-answer',
        ],
    ];
    function detectQuestionType(...values) {
        const text = (0, text_4.normalizeInlineWhitespace)(values.filter(Boolean).join(' '));
        if (!text)
            return null;
        const numericType = LEGACY_NUMERIC_TYPES[text];
        if (numericType)
            return numericType;
        for (const [pattern, type] of TYPE_PATTERNS) {
            if (pattern.test(text))
                return type;
        }
        return null;
    }
    function inferQuestionType(container) {
        const explicit = detectQuestionType(container.getAttribute('typeName'), container.getAttribute('typename'), container.getAttribute('data-type-name'), container.getAttribute('data-question-type'), container.getAttribute('data-type'), container.querySelector('.type_tit, .newTestType, .question-type, [class*="typeName"]')
            ?.textContent);
        if (explicit)
            return explicit;
        if (container.querySelector('input[type="checkbox"]'))
            return 'multiple-choice';
        if (container.querySelector('input[type="radio"]'))
            return 'single-choice';
        if (container.querySelector('textarea, [contenteditable="true"]'))
            return 'short-answer';
        const optionTexts = Array.from(container.querySelectorAll('.answerBg, .option, .Zy_ulTop li, .mark_letter li')).map((option) => (0, text_4.normalizeInlineWhitespace)(option.textContent ?? ''));
        if (optionTexts.length === 2 &&
            optionTexts.every((value) => /(?:正确|错误|对|错|true|false|√|×)/iu.test(value))) {
            return 'true-false';
        }
        return null;
    }
});
define("extractors/generic-exercise-extractor", ["require", "exports", "utils/dom", "utils/array", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-stem", "extractors/question-type"], function (require, exports, dom_3, array_1, answer_comparison_2, common_answer_1, option_parser_2, question_factory_1, question_stem_1, question_type_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.GenericExerciseExtractor = void 0;
    const CONTAINER_SELECTORS = [
        '[data-question-id]',
        '.question-item',
        '.subject-item',
        '.exercise-question',
        '.TiMu',
        '.questionLi',
    ];
    class GenericExerciseExtractor {
        constructor() {
            this.id = 'generic-exercise';
            this.confidence = 40;
        }
        supports(root) {
            return (0, dom_3.allMatches)(root, CONTAINER_SELECTORS).length > 0;
        }
        extract(context) {
            const containers = (0, array_1.uniqueBy)((0, dom_3.allMatches)(context.root, CONTAINER_SELECTORS), (element) => {
                return element.getAttribute('data-question-id') ?? `${element.tagName}:${(0, dom_3.textOf)(element).slice(0, 80)}`;
            });
            const questions = [];
            for (const container of containers) {
                const type = (0, question_type_1.detectQuestionType)(container.getAttribute('typeName'), container.getAttribute('data-question-type'), (0, dom_3.textOf)(container.querySelector('.question-type, .type_tit, .newTestType'))) ?? (0, question_type_1.inferQuestionType)(container);
                if (!type)
                    continue;
                const stem = (0, question_stem_1.resolveQuestionStem)(container);
                const question = (0, question_factory_1.createQuestion)({
                    number: (0, dom_3.parseLeadingNumber)(stem.rawText),
                    type,
                    typeMeta: (0, dom_3.textOf)(container.querySelector('.question-type, .colorShallow')) || undefined,
                    stem: stem.content,
                    options: (0, option_parser_2.parseOptions)(container, [
                        '.answerBg',
                        '.option-list > li',
                        '.options > li',
                        '.mark_letter > li',
                        '.Zy_ulTop > li',
                        'label.option',
                    ]),
                    correctAnswer: (0, common_answer_1.extractCorrectAnswer)(container),
                    userAnswer: (0, common_answer_1.extractUserAnswer)(container),
                    analysis: (0, common_answer_1.extractAnalysis)(container),
                    explicitWrong: (0, answer_comparison_2.hasExplicitWrongMarker)(container),
                    source: {
                        extractor: this.id,
                        pageUrl: context.pageUrl,
                        selector: CONTAINER_SELECTORS.join(', '),
                    },
                });
                if (question)
                    questions.push(question);
            }
            return questions;
        }
    }
    exports.GenericExerciseExtractor = GenericExerciseExtractor;
});
define("extractors/mark-item-extractor", ["require", "exports", "utils/dom", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-stem", "extractors/question-type"], function (require, exports, dom_4, answer_comparison_3, common_answer_2, option_parser_3, question_factory_2, question_stem_2, question_type_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.MarkItemExtractor = void 0;
    class MarkItemExtractor {
        constructor() {
            this.id = 'mark-item';
            this.confidence = 100;
        }
        supports(root) {
            return root.querySelector('.mark_item .questionLi') !== null;
        }
        extract(context) {
            const questions = [];
            context.root.querySelectorAll('.mark_item').forEach((section) => {
                const sectionType = (0, question_type_2.detectQuestionType)((0, dom_4.textOf)(section.querySelector('.type_tit')));
                if (!sectionType)
                    return;
                section.querySelectorAll('.questionLi').forEach((container) => {
                    const type = (0, question_type_2.detectQuestionType)(container.getAttribute('typeName'), container.getAttribute('data-question-type'), (0, dom_4.textOf)(container.querySelector('.colorShallow'))) ?? sectionType;
                    const stem = (0, question_stem_2.resolveQuestionStem)(container);
                    const question = (0, question_factory_2.createQuestion)({
                        number: (0, dom_4.parseLeadingNumber)(stem.rawText),
                        type,
                        typeMeta: (0, dom_4.textOf)(container.querySelector('.colorShallow')) || undefined,
                        stem: stem.content,
                        options: (0, option_parser_3.parseOptions)(container, [
                            '.mark_letter > li',
                            '.mark_letter li',
                            '.answerBg',
                        ]),
                        correctAnswer: (0, common_answer_2.extractCorrectAnswer)(container),
                        userAnswer: (0, common_answer_2.extractUserAnswer)(container),
                        analysis: (0, common_answer_2.extractAnalysis)(container),
                        explicitWrong: (0, answer_comparison_3.hasExplicitWrongMarker)(container),
                        source: {
                            extractor: this.id,
                            pageUrl: context.pageUrl,
                            selector: '.mark_item .questionLi',
                        },
                    });
                    if (question)
                        questions.push(question);
                });
            });
            return questions;
        }
    }
    exports.MarkItemExtractor = MarkItemExtractor;
});
define("extractors/page-context", ["require", "exports", "utils/text"], function (require, exports, text_5) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.resolvePrevTitle = resolvePrevTitle;
    exports.resolvePageTitle = resolvePageTitle;
    exports.resolvePageUrl = resolvePageUrl;
    const TITLE_SELECTORS = [
        '.mark_title',
        '.newTestTitle',
        '.TestTitle_name',
        '.testTitle',
        '.courseName',
        '.chapterText',
        '.chapter-name',
        'h1',
    ];
    // 学生学习页面：章节/测验标题所在的容器，优先作为提取标题来源。
    // 真实 DOM 用类名（<div class="prev_title" title="毛泽东思想的主要内容">），
    // 这里同时兼容 id 形式，两者取先命中者。
    const PREV_TITLE_SELECTORS = ['#prev_title', '.prev_title'];
    /**
     * 学习通学生学习页面（章节页外层壳）上的章节标题。
     *
     * 该节点位于外层壳文档，而题目在嵌套的知识卡片 / 答题 iframe 里，嵌套文档只能解析出
     * 自己的标题（如「章节测验 待完成」），因此需要由外层壳提供章节名。无此节点时返回空串。
     */
    function resolvePrevTitle(root) {
        for (const selector of PREV_TITLE_SELECTORS) {
            const element = root.querySelector(selector);
            if (!element)
                continue;
            const fromAttribute = (0, text_5.normalizeInlineWhitespace)(element.getAttribute('title') ?? '');
            if (fromAttribute)
                return fromAttribute;
            const fromText = (0, text_5.normalizeInlineWhitespace)(element.textContent ?? '');
            if (fromText)
                return fromText;
        }
        return '';
    }
    function resolvePageTitle(root) {
        // 学生学习页面：优先使用 #prev_title / .prev_title 的 title 属性，回退到其正文
        const prevTitle = resolvePrevTitle(root);
        if (prevTitle)
            return prevTitle;
        for (const selector of TITLE_SELECTORS) {
            const value = (0, text_5.normalizeInlineWhitespace)(root.querySelector(selector)?.textContent ?? '');
            if (value)
                return value;
        }
        return (0, text_5.normalizeInlineWhitespace)(root.title)
            .replace(/[-_|]\s*(?:超星学习通|学习通).*$/u, '')
            .trim() || '学习通题目';
    }
    function resolvePageUrl(root) {
        try {
            return root.location.href;
        }
        catch {
            return window.location.href;
        }
    }
});
define("extractors/question-li-extractor", ["require", "exports", "utils/dom", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-stem", "extractors/question-type"], function (require, exports, dom_5, answer_comparison_4, common_answer_3, option_parser_4, question_factory_3, question_stem_3, question_type_3) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.QuestionLiExtractor = void 0;
    class QuestionLiExtractor {
        constructor() {
            this.id = 'question-li';
            this.confidence = 90;
        }
        supports(root) {
            return root.querySelector('.questionLi') !== null;
        }
        extract(context) {
            const questions = [];
            context.root.querySelectorAll('.questionLi').forEach((container) => {
                const type = (0, question_type_3.detectQuestionType)(container.getAttribute('typeName'), container.getAttribute('typename'), container.getAttribute('data-question-type'), (0, dom_5.textOf)(container.querySelector('.type_tit, .question-type, .colorShallow'))) ?? (0, question_type_3.inferQuestionType)(container);
                if (!type)
                    return;
                const stem = (0, question_stem_3.resolveQuestionStem)(container);
                const question = (0, question_factory_3.createQuestion)({
                    number: (0, dom_5.parseLeadingNumber)(stem.rawText),
                    type,
                    typeMeta: (0, dom_5.textOf)(container.querySelector('.colorShallow, .question-type')) || undefined,
                    stem: stem.content,
                    options: (0, option_parser_4.parseOptions)(container, [
                        '.answerBg',
                        '.mark_letter > li',
                        '.option-list > li',
                        '.options > li',
                    ]),
                    correctAnswer: (0, common_answer_3.extractCorrectAnswer)(container),
                    userAnswer: (0, common_answer_3.extractUserAnswer)(container),
                    analysis: (0, common_answer_3.extractAnalysis)(container),
                    explicitWrong: (0, answer_comparison_4.hasExplicitWrongMarker)(container),
                    source: {
                        extractor: this.id,
                        pageUrl: context.pageUrl,
                        selector: '.questionLi',
                    },
                });
                if (question)
                    questions.push(question);
            });
            return questions;
        }
    }
    exports.QuestionLiExtractor = QuestionLiExtractor;
});
define("extractors/timu-extractor", ["require", "exports", "utils/dom", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-stem", "extractors/question-type", "extractors/rich-content"], function (require, exports, dom_6, answer_comparison_5, common_answer_4, option_parser_5, question_factory_4, question_stem_4, question_type_4, rich_content_6) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.TiMuExtractor = exports.TIMU_OPTION_SELECTORS = void 0;
    /**
     * 选项容器候选（当前版本 → 历史版本），命中即用；解析与诊断报告共用同一份，
     * 新增候选只改这里，不要在两处各写一套。
     */
    exports.TIMU_OPTION_SELECTORS = [
        '.Zy_ulTop.qtDetail > li',
        '.Zy_ulTop > li',
        '.answerBg',
        '.option-list > li',
    ];
    /** 「单题块」候选：选项被挪出题目容器时，通常仍留在这几个块里 */
    const QUESTION_BLOCK_SELECTOR = '.singleQuesId, .questionLi, .mark_item';
    /** 页面给每道题标的 qid（形如 405842140）：跨容器找选项时用它精确定位，不会串题 */
    function resolveQuestionId(container) {
        const block = container.closest(QUESTION_BLOCK_SELECTOR);
        const candidates = [
            container.getAttribute('qid'),
            container.getAttribute('data-qid'),
            block?.getAttribute('data'),
            block?.getAttribute('id'),
            container.getAttribute('id'),
            container.querySelector('input[id^="answer"]')?.getAttribute('id'),
            container.querySelector('li[qid]')?.getAttribute('qid'),
        ];
        for (const raw of candidates) {
            // 要求至少 4 位数字，避免把 <i class="fl">1</i> 这类题号当成 qid
            const matched = (raw ?? '').match(/\d{4,}/u);
            if (matched)
                return matched[0];
        }
        return '';
    }
    /** 按 qid 取该题的选项节点（只认页面自己标的 qid） */
    function parseOptionsByQuestionId(scope, qid) {
        for (const selector of [`li[qid="${qid}"]`, `li[data-qid="${qid}"]`]) {
            let elements;
            try {
                elements = Array.from(scope.querySelectorAll(selector));
            }
            catch {
                continue;
            }
            if (elements.length === 0)
                continue;
            const options = elements
                .map((element, index) => (0, option_parser_5.parseLegacyOption)(element, index))
                .filter((option) => option !== null);
            if (options.length > 0)
                return options.sort((left, right) => left.key.localeCompare(right.key));
        }
        return [];
    }
    /**
     * 选项解析：先取题目容器内的候选，容器内一个都取不到时向外层逐级兜底。
     *
     * 已批阅视图（章节测验「已完成」）的部分版本把选项放在题目容器之外（同级表单 /
     * 外层单题块里），此时容器内所有候选都会落空，表现为「题目都提取到了、选项一个都没有」
     * ——即使选项明明显示在页面上。兜底只在容器内确实取不到时触发，命中即用，正常页面不受影响。
     */
    function resolveOptions(container) {
        const direct = (0, option_parser_5.parseOptions)(container, exports.TIMU_OPTION_SELECTORS);
        if (direct.length > 0)
            return direct;
        const block = container.closest(QUESTION_BLOCK_SELECTOR);
        for (const scope of [block, container.parentElement]) {
            if (!scope || scope === container)
                continue;
            // 只有这一层确实只装了一题时才用类名候选，避免把邻题的选项串进来
            if (scope.querySelectorAll('.TiMu').length > 1)
                continue;
            const options = (0, option_parser_5.parseOptions)(scope, exports.TIMU_OPTION_SELECTORS);
            if (options.length > 0)
                return options;
        }
        const qid = resolveQuestionId(container);
        if (qid) {
            for (const scope of [block, container.parentElement, container.ownerDocument]) {
                if (!scope)
                    continue;
                const options = parseOptionsByQuestionId(scope, qid);
                if (options.length > 0)
                    return options;
            }
        }
        return [];
    }
    /**
     * 取题目分组标题（`h3.newTestType`）声明的题型。
     *
     * 答题页把分组标题放在 `.aiArea` 的后代里，而已批阅视图（章节测验「已完成」）把它放在
     * `.aiArea` 之前的兄弟节点上，只查后代会永远落空。落空的后果是：题干里没有
     * 「【单选题】」这类标记时（例如填空题）整题被丢弃，因此这里补上兄弟节点查找。
     */
    function resolveGroupType(area) {
        const explicit = (0, question_type_4.detectQuestionType)((0, dom_6.textOf)(area.querySelector('.newTestType')), area.getAttribute('data-question-type'));
        if (explicit)
            return explicit;
        // 向上找最近的一个分组标题：标题与它后面的若干 .aiArea 是同级兄弟
        for (let sibling = area.previousElementSibling; sibling; sibling = sibling.previousElementSibling) {
            if (sibling.matches('.newTestType'))
                return (0, question_type_4.detectQuestionType)((0, dom_6.textOf)(sibling));
        }
        return null;
    }
    class TiMuExtractor {
        constructor() {
            this.id = 'timu';
            this.confidence = 110;
        }
        supports(root) {
            return root.querySelector('.TiMu.newTiMu, #ZyBottom .TiMu') !== null;
        }
        extract(context) {
            const questions = [];
            const areas = Array.from(context.root.querySelectorAll('#ZyBottom .aiArea'));
            if (areas.length === 0) {
                context.root.querySelectorAll('.TiMu.newTiMu').forEach((container) => {
                    const question = this.extractQuestion(container, null, context);
                    if (question)
                        questions.push(question);
                });
                return questions;
            }
            let currentType = null;
            for (const area of areas) {
                currentType = resolveGroupType(area) ?? currentType;
                area.querySelectorAll('.TiMu.newTiMu, .TiMu').forEach((container) => {
                    const question = this.extractQuestion(container, currentType, context);
                    if (question)
                        questions.push(question);
                });
            }
            return questions;
        }
        extractQuestion(container, inheritedType, context) {
            const title = container.querySelector('.Zy_TItle, .question-title') ?? container;
            const type = (0, question_type_4.detectQuestionType)((0, dom_6.textOf)(title.querySelector('.newZy_TItle')), container.getAttribute('typeName'), container.getAttribute('data-question-type')) ??
                inheritedType ??
                (0, question_type_4.inferQuestionType)(container);
            if (!type)
                return null;
            // 题干经分层定位，避免学习通改版换掉题干类名后整题被丢弃
            const stem = (0, question_stem_4.resolveQuestionStem)(container, title);
            // 「已批阅」视图没有正确答案文本，只有批阅结果图标，拿不到答案时按批阅结果推断
            const userAnswer = (0, common_answer_4.extractUserAnswer)(container);
            const correctAnswer = (0, common_answer_4.extractCorrectAnswer)(container);
            return (0, question_factory_4.createQuestion)({
                number: (0, dom_6.parseLeadingNumber)((0, dom_6.textOf)(title.querySelector('i.fl')) || stem.rawText),
                type,
                typeMeta: (0, dom_6.textOf)(title.querySelector('.newZy_TItle')) || undefined,
                stem: stem.content,
                options: resolveOptions(container),
                correctAnswer: (0, rich_content_6.isRichContentEmpty)(correctAnswer)
                    ? (0, common_answer_4.inferCorrectAnswerFromGrading)(container, userAnswer)
                    : correctAnswer,
                userAnswer,
                analysis: (0, common_answer_4.extractAnalysis)(container),
                explicitWrong: (0, answer_comparison_5.hasExplicitWrongMarker)(container),
                source: {
                    extractor: this.id,
                    pageUrl: context.pageUrl,
                    selector: '.TiMu.newTiMu',
                },
            });
        }
    }
    exports.TiMuExtractor = TiMuExtractor;
});
define("extractors/composite-extractor", ["require", "exports", "domain/question", "utils/array", "extractors/generic-exercise-extractor", "extractors/mark-item-extractor", "extractors/page-context", "extractors/question-li-extractor", "extractors/timu-extractor"], function (require, exports, question_1, array_2, generic_exercise_extractor_1, mark_item_extractor_1, page_context_1, question_li_extractor_1, timu_extractor_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.CompositeExtractor = void 0;
    class CompositeExtractor {
        constructor(extractors = [
            new timu_extractor_1.TiMuExtractor(),
            new mark_item_extractor_1.MarkItemExtractor(),
            new question_li_extractor_1.QuestionLiExtractor(),
            new generic_exercise_extractor_1.GenericExerciseExtractor(),
        ]) {
            this.extractors = extractors;
        }
        extract(root) {
            const title = (0, page_context_1.resolvePageTitle)(root);
            const pageUrl = (0, page_context_1.resolvePageUrl)(root);
            const candidates = [];
            for (const extractor of this.extractors) {
                if (!extractor.supports(root))
                    continue;
                const questions = (0, array_2.uniqueBy)(extractor.extract({ root, title, pageUrl }), (question) => {
                    return question.id;
                });
                if (questions.length > 0) {
                    candidates.push({ extractor: extractor.id, confidence: extractor.confidence, questions });
                }
            }
            const best = candidates.sort((left, right) => {
                const confidenceDifference = right.confidence - left.confidence;
                return confidenceDifference || right.questions.length - left.questions.length;
            })[0];
            if (!best)
                return null;
            return this.buildResult(title, pageUrl, best.extractor, best.questions);
        }
        buildResult(title, sourceUrl, extractor, questions) {
            return {
                title,
                questions,
                typeOrder: (0, question_1.deriveTypeOrder)(questions),
                statistics: (0, question_1.buildStatistics)(questions),
                sourceUrl,
                extractor,
                extractedAt: new Date().toISOString(),
            };
        }
    }
    exports.CompositeExtractor = CompositeExtractor;
});
define("extractors/dom-diagnostics", ["require", "exports", "utils/dom", "extractors/composite-extractor", "extractors/cxsecret-decoder", "extractors/option-parser", "extractors/question-stem", "extractors/rich-content", "extractors/timu-extractor"], function (require, exports, dom_7, composite_extractor_1, cxsecret_decoder_2, option_parser_6, question_stem_5, rich_content_7, timu_extractor_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.buildExtractionDiagnostics = buildExtractionDiagnostics;
    // 报告总长度上限与各块截断长度：只保留判断结构所需的信息
    const MAX_REPORT_CHARS = 9000;
    const MAX_QUESTIONS_PER_DOC = 6;
    const MAX_OUTLINE_DEPTH = 3;
    /** 一个文档里可能承载题目的容器，与各提取器的 supports() 保持一致 */
    const QUESTION_CONTAINER_SELECTORS = [
        '.TiMu.newTiMu',
        '.TiMu',
        '.questionLi',
        '.mark_item',
        '.answerBg',
    ];
    /** 题目容器的外层结构：选项若不在题目容器里，最可能落在这一层 */
    const QUESTION_SCOPE_SELECTORS = ['.singleQuesId', '.questionLi', '.mark_item'];
    /** 需要选项才算完整的题型（此处就地判断，避免反向依赖应用层） */
    const CHOICE_TYPES = new Set(['single-choice', 'multiple-choice']);
    function safeQuery(root, selector) {
        try {
            return root.querySelector(selector);
        }
        catch {
            return null;
        }
    }
    function safeQueryAllCount(root, selector) {
        try {
            return root.querySelectorAll(selector).length;
        }
        catch {
            return -1;
        }
    }
    /** 空白折叠 + 截断，便于把结构压进一行 */
    function clip(value, limit) {
        const text = value.replace(/\s+/gu, ' ').trim();
        return text.length <= limit ? text : `${text.slice(0, limit)}…(+${text.length - limit}字)`;
    }
    function describeElement(element) {
        const tag = element.tagName.toLowerCase();
        const id = element.getAttribute('id');
        const className = (element.getAttribute('class') ?? '').trim().replace(/\s+/gu, '.');
        return `${tag}${id ? `#${id}` : ''}${className ? `.${className}` : ''}`;
    }
    /**
     * 元素文本：走与提取相同的富文本管道（含 cxsecret 还原）。
     * 直接读 textContent 拿到的是字体反爬的混淆码位，报告里会满屏乱码、无法核对。
     */
    function readableText(element) {
        if (!element)
            return '';
        return (0, rich_content_7.richContentToText)((0, rich_content_7.extractRichContent)(element));
    }
    /** 元素结构速览：标签.类名 > 子元素；深度与节点数都受限，足以看出选项容器在不在 */
    function outline(node, depth, budget) {
        if (budget.left <= 0)
            return '…';
        budget.left -= 1;
        const name = describeElement(node);
        const children = Array.from(node.children);
        if (children.length === 0)
            return name;
        if (depth <= 1)
            return `${name}(+${children.length} 个子元素)`;
        const parts = [];
        for (const child of children) {
            if (budget.left <= 0) {
                parts.push('…');
                break;
            }
            parts.push(outline(child, depth - 1, budget));
        }
        return `${name} > [${parts.join(', ')}]`;
    }
    /** 题目的 qid：答题页放在 .singleQuesId 的 id/属性上，用于跨容器找选项 */
    function resolveQuestionId(container) {
        const scope = container.closest('.singleQuesId') ?? container;
        const raw = [scope.getAttribute('id'), scope.getAttribute('data'), container.getAttribute('data')]
            .filter(Boolean)
            .join(' ');
        return raw.match(/\d{4,}/u)?.[0] ?? '(无)';
    }
    /** 逐候选选择器统计容器内的选项节点数，命中即说明该候选可用 */
    function describeOptionCounts(container) {
        return [...timu_extractor_2.TIMU_OPTION_SELECTORS, ...option_parser_6.FALLBACK_OPTION_SELECTORS]
            .map((selector) => `${selector}=${safeQueryAllCount(container, selector)}`)
            .join(' ');
    }
    /** 选项节点总数（容器内任一候选命中即为该题选项） */
    function optionNodeCount(container) {
        for (const selector of [...timu_extractor_2.TIMU_OPTION_SELECTORS, ...option_parser_6.FALLBACK_OPTION_SELECTORS]) {
            const count = safeQueryAllCount(container, selector);
            if (count > 0)
                return count;
        }
        return 0;
    }
    /**
     * 容器内实际解析出的选项条数（走与提取完全相同的管道）。
     *
     * 与「候选命中数」对照即可把成因分成两类：命中数为 0 是选择器没命中；
     * 命中数大于 0、解析数为 0，则是节点找到了但读不出文字（历史真机现场即此类）。
     */
    function parsedOptionCount(container) {
        try {
            return (0, option_parser_6.parseOptions)(container, timu_extractor_2.TIMU_OPTION_SELECTORS).length;
        }
        catch {
            return -1;
        }
    }
    /** 单题的诊断块：容器、qid、题干来源、选项计数与结构速览 */
    function describeQuestion(container, index) {
        const matched = QUESTION_CONTAINER_SELECTORS.find((selector) => {
            try {
                return container.matches(selector);
            }
            catch {
                return false;
            }
        }) ?? '(未匹配已知容器)';
        const title = safeQuery(container, '.Zy_TItle, .question-title') ?? container;
        const stemElement = (0, dom_7.firstMatch)(title, question_stem_5.STEM_CONTAINER_SELECTORS) ?? (0, dom_7.firstMatch)(container, question_stem_5.STEM_CONTAINER_SELECTORS);
        const stemSelector = question_stem_5.STEM_CONTAINER_SELECTORS.find((selector) => safeQuery(title, selector) !== null) ??
            question_stem_5.STEM_CONTAINER_SELECTORS.find((selector) => safeQuery(container, selector) !== null) ??
            '(未命中，走兜底)';
        const qid = resolveQuestionId(container);
        const options = optionNodeCount(container);
        const parsed = parsedOptionCount(container);
        const head = `  ${index + 1}) 容器=${matched} qid=${qid} 题干来源=${stemSelector} ` +
            `题干=${clip(readableText(stemElement), 60) || '(空)'}`;
        const counts = `     容器内选项计数: ${describeOptionCounts(container)} 解析出选项=${parsed} 条`;
        // 节点命中且能解析出选项：该题选项链路正常，无需展开结构
        if (options > 0 && parsed > 0) {
            return [head, counts];
        }
        // 其余情况展开结构。同时打印「原样 textContent」与「提取管道读出的文本」：
        // 原文有文字、管道为空，即说明文字确实在节点里、只是没被提取管道读到
        // （历史上这条差异来自跨 window 的构造器判定，见 rich-content.ts 的 isElementNode）。
        const lines = [head, counts];
        const items = Array.from(container.querySelectorAll('li')).slice(0, 8);
        const raw = items.map((item) => clip(item.textContent ?? '', 24) || '(空)');
        const piped = items.map((item) => clip(readableText(item), 24) || '(空)');
        lines.push(`     容器内 li 原文(${container.querySelectorAll('li').length} 个): ${raw.join(' | ') || '(无)'}`);
        lines.push(`     容器内 li 管道: ${piped.join(' | ') || '(无)'}`);
        for (const selector of QUESTION_SCOPE_SELECTORS) {
            const scope = container.closest(selector);
            if (!scope || scope === container)
                continue;
            lines.push(`     外层 ${selector}: 选项计数 ${describeOptionCounts(scope)}`);
            lines.push(`     外层结构: ${clip(outline(scope, MAX_OUTLINE_DEPTH, { left: 26 }), 700)}`);
        }
        if (qid !== '(无)') {
            const owner = container.ownerDocument;
            const byQid = safeQueryAllCount(owner, `li[qid="${qid}"]`);
            const byData = safeQueryAllCount(owner, `li[data-qid="${qid}"]`);
            lines.push(`     按 qid 全文档查找: li[qid]=${byQid} li[data-qid]=${byData}` +
                `（>0 说明选项在题目容器之外）`);
        }
        lines.push(`     容器结构: ${clip(outline(container, MAX_OUTLINE_DEPTH, { left: 26 }), 900)}`);
        return lines;
    }
    /** 单个文档的诊断块 */
    function describeDocument(root, depth, extractor) {
        const url = (() => {
            try {
                return root.location?.href ?? '(无 url)';
            }
            catch {
                return '(跨域)';
            }
        })();
        const decoder = (0, cxsecret_decoder_2.getCxSecretDecoder)(root);
        const containers = [];
        for (const selector of QUESTION_CONTAINER_SELECTORS) {
            const found = Array.from(root.querySelectorAll(selector));
            for (const element of found)
                if (!containers.includes(element))
                    containers.push(element);
        }
        const lines = [
            `[L${depth}] readyState=${root.readyState ?? '未知'} 题目容器=${QUESTION_CONTAINER_SELECTORS.map((selector) => `${selector}×${safeQueryAllCount(root, selector)}`).join(' ')}`,
            `     cxsecret: 作用域标记=${decoder ? (decoder.scoped ? '有' : '无') : '无字体'} 解码表=${decoder?.size ?? 0} 条`,
            `     url=${clip(url, 120)}`,
        ];
        const result = extractor.extract(root);
        if (result) {
            const missing = result.questions.filter((question) => CHOICE_TYPES.has(question.type) && question.options.length === 0).length;
            lines.push(`     链路自检: 提取器=${result.extractor} 题数=${result.questions.length} ` +
                `选项总数=${result.questions.reduce((total, question) => total + question.options.length, 0)} ` +
                `缺选项选择题=${missing}`);
        }
        else {
            lines.push('     链路自检: 该文档未提取到题目');
        }
        containers.slice(0, MAX_QUESTIONS_PER_DOC).forEach((container, index) => {
            lines.push(...describeQuestion(container, index));
        });
        if (containers.length > MAX_QUESTIONS_PER_DOC) {
            lines.push(`  （其余 ${containers.length - MAX_QUESTIONS_PER_DOC} 个题目容器已省略）`);
        }
        return lines;
    }
    /**
     * 生成当前页面的提取诊断报告。
     * 报告包含：各层文档（含同源 iframe）的就绪状态、题目容器数量、逐题的选项候选命中数、
     * 容器/外层结构速览、以及提取链路自检结果。
     */
    function buildExtractionDiagnostics(options = {}) {
        const root = options.root ?? (typeof document === 'undefined' ? null : document);
        const lines = [
            `=== 学习通题目导出 · 提取诊断 ${options.version ? `v${options.version} ` : ''}${new Date().toISOString()} ===`,
            '说明：把整段内容复制回传，即可定位「提取到题干、提取不到选项」的成因。',
            '读法：每题的「解析出选项」= 提取管道真正读出的选项条数；候选命中数>0 而解析数=0，',
            '说明节点找到了但读不出文字。「li 原文」是未处理的 textContent、「li 管道」是提取管道读到的',
            '文本，原文有而管道为空即属此类（字体反爬页面的原文会是混淆码位，以管道为准）。',
        ];
        if (!root) {
            lines.push('（当前环境没有 document，无法诊断）');
            return lines.join('\n');
        }
        const extractor = new composite_extractor_1.CompositeExtractor();
        for (const { document: current, depth } of (0, dom_7.collectAccessibleDocuments)(root)) {
            lines.push(...describeDocument(current, depth, extractor));
        }
        const report = lines.join('\n');
        if (report.length <= MAX_REPORT_CHARS)
            return report;
        return `${report.slice(0, MAX_REPORT_CHARS)}\n…（报告已截断）`;
    }
});
define("exporters/legacy-bridge", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.emptyLegacyResults = emptyLegacyResults;
    exports.hasRichContent = hasRichContent;
    exports.normalizeRichContent = normalizeRichContent;
    exports.richContentToText = richContentToText;
    exports.questionContent = questionContent;
    exports.optionContent = optionContent;
    exports.answerContent = answerContent;
    exports.escapeMarkdownAlt = escapeMarkdownAlt;
    exports.escapeMarkdownUrl = escapeMarkdownUrl;
    exports.formatRichForText = formatRichForText;
    exports.formatRichForMD = formatRichForMD;
    exports.richTextOnly = richTextOnly;
    exports.shuffleQuestions = shuffleQuestions;
    exports.toLegacyResult = toLegacyResult;
    exports.toLegacyChapter = toLegacyChapter;
    // QuestionType → 原版中文题型键映射
    const TYPE_KEY_MAP = {
        'single-choice': '单选',
        'multiple-choice': '多选',
        'fill-blank': '填空',
        'true-false': '判断',
        'short-answer': '简答',
    };
    // 创建原版结构的空题目集合（五种题型各一个空数组）
    function emptyLegacyResults() {
        return { '单选': [], '多选': [], '填空': [], '判断': [], '简答': [] };
    }
    // ==================== 原版富文本辅助函数（逐行移植） ====================
    // 判断富文本是否有实质内容（文本或图片）
    function hasRichContent(content) {
        return (content || []).some((part) => part.type === 'image' || (part.type === 'text' && part.text.trim().length > 0));
    }
    // 规范化富文本数组：合并相邻文本节点、去首尾空行、压缩空白
    function normalizeRichContent(parts) {
        const normalized = [];
        const pushText = (part) => {
            if (!part.text)
                return;
            const value = part.text.replace(/\u00a0/g, ' ').replace(/[ \t\r\f]+/g, ' ');
            if (!value)
                return;
            const last = normalized[normalized.length - 1];
            // 合并相邻文本，保留第一个 part 的格式属性
            if (last && last.type === 'text') {
                normalized[normalized.length - 1] = { ...last, text: last.text + value };
            }
            else {
                normalized.push({
                    type: 'text',
                    text: value,
                    bold: part.bold,
                    italic: part.italic,
                    subScript: part.subScript,
                    superScript: part.superScript,
                });
            }
        };
        const pushBreak = () => {
            const last = normalized[normalized.length - 1];
            if (!last || last.type !== 'break')
                normalized.push({ type: 'break' });
        };
        for (const part of parts || []) {
            if (!part)
                continue;
            if (part.type === 'text') {
                pushText(part);
            }
            else if (part.type === 'image' && part.url) {
                normalized.push(part);
            }
            else if (part.type === 'break') {
                pushBreak();
            }
        }
        // 修剪首尾的换行与纯空白文本：尾部空白会挡住相邻换行，导致 Word 导出时题目与选项之间出现多余空行
        const isTrimmable = (part) => {
            if (!part)
                return false;
            return part.type === 'break' || (part.type === 'text' && !part.text.trim());
        };
        while (isTrimmable(normalized[0]))
            normalized.shift();
        while (isTrimmable(normalized[normalized.length - 1]))
            normalized.pop();
        return normalized;
    }
    // 将富文本数组转为纯文本，图片通过回调格式化
    function richContentToText(content, imageFormatter) {
        let text = '';
        for (const part of content || []) {
            if (part.type === 'text')
                text += part.text;
            else if (part.type === 'image')
                text += imageFormatter ? imageFormatter(part.url, part) : '';
            else if (part.type === 'break')
                text += '\n';
        }
        return text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    }
    // 适配器：兼容新旧数据格式，统一返回富文本数组
    function questionContent(q) {
        if (hasRichContent(q.stemContent))
            return q.stemContent;
        const parts = q.stem ? [{ type: 'text', text: q.stem }] : [];
        if (q.images && q.images.length) {
            q.images.forEach((url) => parts.push({ type: 'break' }, { type: 'image', url, alt: '' }));
        }
        return normalizeRichContent(parts);
    }
    function optionContent(opt) {
        if (hasRichContent(opt.content))
            return opt.content;
        return opt.text ? [{ type: 'text', text: opt.text }] : [];
    }
    function answerContent(q) {
        if (hasRichContent(q.correctAnswerContent))
            return q.correctAnswerContent;
        return q.correctAnswer ? [{ type: 'text', text: q.correctAnswer }] : [];
    }
    // Markdown URL 转义，防止特殊字符破坏图片语法
    function escapeMarkdownAlt(text) {
        return (text || '图片').replace(/[[\]\n\r]/g, ' ').trim() || '图片';
    }
    function escapeMarkdownUrl(url) {
        // 编码 URL 中的特殊字符，防止破坏 Markdown 图片语法，保留已编码部分
        return (url || '').replace(/[()\\]/g, (ch) => '%' + ch.charCodeAt(0).toString(16).toUpperCase());
    }
    // 富文本格式化：TXT/MD 分别处理
    function formatRichForText(content) {
        return richContentToText(content, (url) => `\n[图片: ${url}]\n`);
    }
    function formatRichForMD(content) {
        return richContentToText(content, (url, part) => `\n![${escapeMarkdownAlt(part.alt)}](${escapeMarkdownUrl(url)})\n`);
    }
    // 富文本转纯文本（忽略图片）
    function richTextOnly(content) {
        return richContentToText(content || [], () => '').replace(/\s+/g, ' ').trim();
    }
    // Fisher-Yates 洗牌算法，同类型题目内部打乱
    function shuffleQuestions(results, typeOrder) {
        const shuffled = emptyLegacyResults();
        for (const qtype of typeOrder) {
            const arr = [...(results[qtype] || [])];
            for (let i = arr.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                // Fisher-Yates 交换（noUncheckedIndexedAccess 下需显式判空）
                const left = arr[i];
                const right = arr[j];
                if (left && right) {
                    arr[i] = right;
                    arr[j] = left;
                }
            }
            shuffled[qtype] = arr;
        }
        return shuffled;
    }
    // ==================== 领域模型 → 原版结构转换 ====================
    // 将 TS 领域模型的单题转换为原版题目结构
    function toLegacyQuestion(question) {
        return {
            stem: formatRichForText(question.stem),
            stemContent: question.stem,
            typeMeta: question.typeMeta,
            options: question.options.map((option) => ({
                letter: option.key,
                text: richTextOnly(option.content),
                content: option.content,
            })),
            correctAnswer: richContentToText(question.correctAnswer, () => '')
                .replace(/\s+/g, ' ')
                .trim(),
            correctAnswerContent: question.correctAnswer,
            myAnswer: richTextOnly(question.userAnswer),
            isWrong: question.isWrong,
        };
    }
    // 将 TS 领域模型结果转换为原版提取数据结构（全部题目按题型合并）
    function toLegacyResult(result) {
        const results = emptyLegacyResults();
        for (const question of result.questions) {
            results[TYPE_KEY_MAP[question.type]].push(toLegacyQuestion(question));
        }
        const typeOrder = result.typeOrder.map((type) => TYPE_KEY_MAP[type]);
        return {
            title: result.title,
            results,
            typeOrder,
            wrongCount: result.statistics.wrong,
            hasMyAnswer: result.statistics.withUserAnswer > 0,
            hasCorrectAnswer: result.statistics.withCorrectAnswer > 0,
        };
    }
    // 将 TS 章节数据转换为原版分章节数据结构
    function toLegacyChapter(chapter) {
        const results = emptyLegacyResults();
        for (const question of chapter.questions) {
            results[TYPE_KEY_MAP[question.type]].push(toLegacyQuestion(question));
        }
        const typeOrder = chapter.typeOrder.map((type) => TYPE_KEY_MAP[type]);
        return {
            title: chapter.title,
            results,
            typeOrder,
            wrongCount: chapter.statistics.wrong,
            hasMyAnswer: chapter.statistics.withUserAnswer > 0,
            hasCorrectAnswer: chapter.statistics.withCorrectAnswer > 0,
        };
    }
});
define("exporters/text-formatter", ["require", "exports", "exporters/legacy-bridge"], function (require, exports, legacy_bridge_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.formatOutput = formatOutput;
    exports.formatAnswersTXT = formatAnswersTXT;
    exports.formatOutputWithAnswers = formatOutputWithAnswers;
    exports.formatWrongQuestionsTXT = formatWrongQuestionsTXT;
    // ==================== TXT 格式化（富文本） ====================
    function formatOutput(results, typeOrder) {
        const typeLabels = {
            单选: '单选题',
            多选: '多选题',
            填空: '填空题',
            判断: '判断题',
            简答: '简答题',
        };
        const typeNumbers = ['一', '二', '三', '四', '五', '六'];
        let output = '';
        let globalNum = 0;
        let sectionIdx = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            const label = typeLabels[qtype] ?? qtype;
            const num = typeNumbers[sectionIdx] || sectionIdx + 1;
            output += `${num}. ${label}（共${questions.length}题）\n`;
            for (const q of questions) {
                globalNum++;
                output += `${globalNum}. ${(0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.questionContent)(q))}\n`;
                if (q.options && q.options.length > 0) {
                    for (const opt of q.options) {
                        output += `${opt.letter}. ${(0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.optionContent)(opt))}\n`;
                    }
                }
                output += '\n';
            }
            sectionIdx++;
        }
        return output.trim();
    }
    function formatAnswersTXT(results, typeOrder) {
        const typeLabels = {
            单选: '单选题',
            多选: '多选题',
            填空: '填空题',
            判断: '判断题',
            简答: '简答题',
        };
        const typeNumbers = ['一', '二', '三', '四', '五', '六'];
        let output = '';
        let globalNum = 0;
        let sectionIdx = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            const num = typeNumbers[sectionIdx] || sectionIdx + 1;
            output += `${num}、${typeLabels[qtype] ?? qtype}\n\n`;
            sectionIdx++;
            for (const q of questions) {
                globalNum++;
                const answer = (0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.answerContent)(q)) || '（未找到答案）';
                if (qtype === '填空' && answer.includes('；')) {
                    const parts = answer.split('；').map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
                    output += `${globalNum}. \n`;
                    parts.forEach((part, i) => {
                        output += `(${i + 1}) ${part}\n`;
                    });
                    output += '\n';
                }
                else {
                    output += `${globalNum}. ${answer}\n\n`;
                }
            }
        }
        return output.trim();
    }
    function formatOutputWithAnswers(results, typeOrder) {
        let output = formatOutput(results, typeOrder);
        output += '\n\n\n';
        output += '========================================\n';
        output += '              答案汇总\n';
        output += '========================================\n\n';
        output += formatAnswersTXT(results, typeOrder);
        return output.trim();
    }
    function formatWrongQuestionsTXT(results, typeOrder) {
        let output = '';
        output += '\n\n\n';
        output += '========================================\n';
        output += '              错题汇总\n';
        output += '========================================\n\n';
        let globalNum = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            if (qtype === '简答') {
                globalNum += questions.length;
                continue;
            }
            for (const q of questions) {
                globalNum++;
                if (!q.isWrong)
                    continue;
                const typeLabel = qtype === '填空' ? '填空题' : '题目';
                output += `${globalNum}. (${typeLabel})${(0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.questionContent)(q))}\n`;
                output += `   我的答案: ${q.myAnswer || '无'}\n`;
                output += `   正确答案: ${(0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.answerContent)(q)) || '（未找到答案）'}\n\n`;
            }
        }
        return output.replace(/\n+$/, '');
    }
});
define("exporters/markdown-formatter", ["require", "exports", "exporters/legacy-bridge"], function (require, exports, legacy_bridge_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.formatOutputMD = formatOutputMD;
    exports.formatAnswersMD = formatAnswersMD;
    exports.formatOutputWithAnswersMD = formatOutputWithAnswersMD;
    exports.formatWrongQuestionsMD = formatWrongQuestionsMD;
    // ==================== Markdown 格式化（富文本） ====================
    function formatOutputMD(results, typeOrder) {
        const typeLabels = {
            单选: '单选题',
            多选: '多选题',
            填空: '填空题',
            判断: '判断题',
            简答: '简答题',
        };
        const typeNumbers = ['一', '二', '三', '四', '五', '六'];
        let output = '';
        let globalNum = 0;
        let sectionIdx = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            const label = typeLabels[qtype] ?? qtype;
            const num = typeNumbers[sectionIdx] || sectionIdx + 1;
            output += `### ${num}、${label}（共${questions.length}题）\n\n`;
            for (const q of questions) {
                globalNum++;
                output += `**${globalNum}.** ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.questionContent)(q))}\n\n`;
                if (q.options && q.options.length > 0) {
                    for (const opt of q.options) {
                        output += `- ${opt.letter}. ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.optionContent)(opt))}\n`;
                    }
                    output += '\n';
                }
                else {
                    output += '\n';
                }
            }
            sectionIdx++;
        }
        return output.trim();
    }
    function formatAnswersMD(results, typeOrder) {
        let output = '';
        let globalNum = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            const typeLabels = {
                单选: '单选题',
                多选: '多选题',
                填空: '填空题',
                判断: '判断题',
                简答: '简答题',
            };
            output += `**${typeLabels[qtype] ?? qtype}**\n\n`;
            for (const q of questions) {
                globalNum++;
                const answer = (0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.answerContent)(q)) || '（未找到答案）';
                if (qtype === '填空' && answer.includes('；')) {
                    const parts = answer.split('；').map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
                    output += `${globalNum}.  \n`;
                    parts.forEach((part, i) => {
                        output += `    (${i + 1}) ${part}  \n`;
                    });
                    output += '\n';
                }
                else {
                    output += `${globalNum}. ${answer}  \n`;
                }
            }
            output += '\n';
        }
        return output.trim();
    }
    function formatOutputWithAnswersMD(results, typeOrder) {
        let output = formatOutputMD(results, typeOrder);
        output += '\n\n---\n\n';
        output += '## 答案汇总\n\n';
        output += formatAnswersMD(results, typeOrder);
        return output.trim();
    }
    function formatWrongQuestionsMD(results, typeOrder) {
        let output = '\n\n\n---\n\n';
        output += '## 错题汇总\n\n';
        let globalNum = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            if (qtype === '简答') {
                globalNum += questions.length;
                continue;
            }
            for (const q of questions) {
                globalNum++;
                if (!q.isWrong)
                    continue;
                output += `**${globalNum}.** ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.questionContent)(q))}\n\n`;
                output += `- 我的答案: ${q.myAnswer || '无'}\n`;
                output += `- 正确答案: ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.answerContent)(q)) || '（未找到答案）'}\n\n`;
            }
        }
        return output.replace(/\n+$/, '');
    }
});
define("exporters/word-plan", ["require", "exports", "exporters/legacy-bridge"], function (require, exports, legacy_bridge_3) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.planWordDocuments = planWordDocuments;
    /** 去掉用户可能输入的扩展名，得到输出文件的基础名（与 TXT/MD 导出规则一致） */
    function baseName(result, options) {
        return (options.filename || result.title || '学习通题目').replace(/\.(txt|md|docx)$/, '');
    }
    /** 章节名可能含路径分隔符等文件名非法字符，替换后再落盘 */
    function safeChapterTitle(title, index) {
        return (title || `章节${index + 1}`).replace(/[\\/:*?"<>|]/g, '_');
    }
    function toSection(chapter, shuffle) {
        return {
            title: chapter.title,
            results: shuffle ? (0, legacy_bridge_3.shuffleQuestions)(chapter.results, chapter.typeOrder) : chapter.results,
            typeOrder: chapter.typeOrder,
        };
    }
    /**
     * 规划本次 Word 导出要产出哪些文件：
     * - 勾选「按章节拆分文件」且结果是多章节 → **一章一个 .docx**（与 TXT/Markdown 行为一致）；
     * - 否则只产出一个 .docx：**多章节时正文按「章节 → 题型」分节**，单章节时与原先完全一致。
     *
     * 题库导入（智能导入）格式不认章节标题，所以多章节时也不在单文件里插章节标题，
     * 只按需要拆文件——拆出来的每个文件仍然只有一章的题目。
     */
    function planWordDocuments(result, options) {
        const base = baseName(result, options);
        // 题库导入必带答案并禁用打乱，与 export-service 的既有规则保持一致
        const shuffle = !options.bankImport && options.shuffle;
        const chapters = result.chapters ?? [];
        // 分章节下载：每个章节单独一个文件
        if (options.splitByChapter && chapters.length > 1) {
            return chapters.map((chapter, index) => {
                const data = (0, legacy_bridge_3.toLegacyChapter)(chapter);
                const title = safeChapterTitle(data.title, index);
                return {
                    title: data.title || title,
                    sections: [toSection(data, shuffle)],
                    filename: `${base}_${index + 1}_${title}.docx`,
                };
            });
        }
        const legacy = (0, legacy_bridge_3.toLegacyResult)(result);
        // 多章节单文件：正文按章节分节，避免把各章节的同题型题目混在同一大题里
        const groupedByChapter = chapters.length > 1 && !options.bankImport;
        const sections = groupedByChapter
            ? chapters.map((chapter) => toSection((0, legacy_bridge_3.toLegacyChapter)(chapter), shuffle))
            : [
                {
                    title: legacy.title,
                    results: shuffle ? (0, legacy_bridge_3.shuffleQuestions)(legacy.results, legacy.typeOrder) : legacy.results,
                    typeOrder: legacy.typeOrder,
                },
            ];
        return [{ title: legacy.title, sections, filename: `${base}.docx` }];
    }
});
define("exporters/word-exporter", ["require", "exports", "exporters/legacy-bridge"], function (require, exports, legacy_bridge_4) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.WORD_MIME = void 0;
    exports.generateWordBlob = generateWordBlob;
    exports.WORD_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    // ==================== Word 文档生成（富文本） ====================
    async function fetchImageAsset(url) {
        if (!url)
            return null;
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 10000);
            const resp = await fetch(url, { mode: 'cors', signal: controller.signal });
            clearTimeout(timeoutId);
            if (!resp.ok)
                return null;
            const blob = await resp.blob();
            if (!blob.type.startsWith('image/'))
                return null;
            const data = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : null);
                reader.onerror = () => resolve(null);
                reader.readAsDataURL(blob);
            });
            if (!data)
                return null;
            const size = await new Promise((resolve) => {
                const img = new Image();
                img.onload = () => resolve({ width: img.naturalWidth || 300, height: img.naturalHeight || 200 });
                img.onerror = () => resolve({ width: 300, height: 200 });
                img.src = data;
            });
            const typeMap = {
                'image/png': 'png',
                'image/jpeg': 'jpg',
                'image/jpg': 'jpg',
                'image/gif': 'gif',
                'image/bmp': 'bmp',
            };
            const type = typeMap[blob.type] || 'png';
            return { data, type, width: size.width, height: size.height };
        }
        catch {
            return null;
        }
    }
    async function buildImageRun(url) {
        if (!url)
            return null;
        const base64 = await fetchImageAsset(url);
        if (base64) {
            // 按比例限制最大宽高，避免固定尺寸导致变形
            const maxWidth = 420;
            const maxHeight = 260;
            const scale = Math.min(maxWidth / base64.width, maxHeight / base64.height, 1);
            return new docx.ImageRun({
                data: base64,
                transformation: {
                    width: Math.round(base64.width * scale),
                    height: Math.round(base64.height * scale),
                },
                type: base64.type,
            });
        }
        // 图片加载失败，记录计数
        window.__xxt_failed_image_count = (window.__xxt_failed_image_count || 0) + 1;
        return null;
    }
    async function buildRichRuns(content, prefix = '') {
        const { TextRun } = docx;
        const runs = [];
        if (prefix)
            runs.push(new TextRun({ text: prefix, font: 'Microsoft YaHei', size: 22 }));
        for (const part of content || []) {
            if (part.type === 'text') {
                const normalized = part.text.replace(/\n+/g, ' ');
                if (normalized) {
                    runs.push(new TextRun({
                        text: normalized,
                        font: 'Microsoft YaHei',
                        size: 22,
                        bold: part.bold || false,
                        italics: part.italic || false,
                        subScript: part.subScript || false,
                        superScript: part.superScript || false,
                    }));
                }
            }
            else if (part.type === 'image') {
                if (runs.length > 0)
                    runs.push(new TextRun({ text: ' ', font: 'Microsoft YaHei', size: 22 }));
                const imgRun = await buildImageRun(part.url);
                if (imgRun)
                    runs.push(imgRun);
                runs.push(new TextRun({ text: ' ', font: 'Microsoft YaHei', size: 22 }));
            }
            else if (part.type === 'break') {
                runs.push(new TextRun({ text: '\n', break: 1, font: 'Microsoft YaHei', size: 22 }));
            }
        }
        return runs.length ? runs : [new TextRun({ text: prefix, font: 'Microsoft YaHei', size: 22 })];
    }
    // 将富文本内容构建为带段后间距的 Paragraph 数组（Word 导出用）
    async function buildRichParagraphs(content, prefix = '', spacing = 0) {
        const { Paragraph } = docx;
        const runs = await buildRichRuns(content, prefix);
        return [
            new Paragraph({
                children: runs,
                spacing: { after: spacing },
            }),
        ];
    }
    // 清洗答案文本：去除“正确答案:/我的答案:”等标签，避免组合出现
    function cleanAnswerText(text) {
        if (!text)
            return '';
        return text
            .replace(/\u00a0/g, ' ')
            .replace(/\r/g, '\n')
            .replace(/(?:正确答案|参考答案)[:：]?\s*/g, '')
            .replace(/(?:我的答案|你的答案|学生答案)[:：]?\s*/g, '')
            .replace(/^[：:\s]+/, '')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    }
    const FONT = 'Microsoft YaHei';
    // 题型大题头字号（半磅）：比章节标题小一号（章节标题 28 = 14pt → 题型 24 = 12pt），与正文 22 拉开层次
    const SIZE_TYPE_HEADER = 24;
    const COLOR = {
        title: '1F2937',
        muted: '6B7280',
        border: 'E5E7EB',
        answer: '00A870',
        wrong: 'DC2626',
        analysisBg: 'F7F9FC',
        type: '64748B',
    };
    // 题型名（不含序号）：序号由 buildTypeHeaders 按本章实际出现的题型顺序现场排
    const TYPE_NAMES = {
        单选: '单项选择题',
        多选: '多项选择题',
        填空: '填空题',
        判断: '判断题',
        简答: '简答题',
    };
    const CN_NUMERALS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
    /** 大题头的中文序号（超出表长时退化为阿拉伯数字） */
    function ordinalLabel(index) {
        return CN_NUMERALS[index] ?? String(index + 1);
    }
    /**
     * 本章的「题型 → 大题头」映射。
     *
     * 序号按本章**实际有题目**的题型从「一」重新排：此前用全局固定的题型序号，
     * 某章没有单选题时多选题就会顶着「二、多项选择题」的跳号。
     */
    function buildTypeHeaders(section) {
        const headers = new Map();
        let ordinal = 0;
        for (const qtype of section.typeOrder) {
            const questions = section.results[qtype];
            if (!questions || questions.length === 0)
                continue;
            headers.set(qtype, `${ordinalLabel(ordinal)}、${TYPE_NAMES[qtype] ?? qtype}`);
            ordinal += 1;
        }
        return headers;
    }
    const BANK_TYPE_LABELS = {
        单选: '【单选题】',
        多选: '【多选题】',
        填空: '【填空题】',
        判断: '【判断题】',
        简答: '【简答题】',
    };
    /** 题库导入格式的页面设置：A4 纵向，页边距比试卷略宽 */
    function bankImportSections(children) {
        const { convertMillimetersToTwip } = docx;
        return [
            {
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 },
                        margin: {
                            top: convertMillimetersToTwip(20),
                            bottom: convertMillimetersToTwip(20),
                            left: convertMillimetersToTwip(25),
                            right: convertMillimetersToTwip(25),
                        },
                    },
                },
                children,
            },
        ];
    }
    /** 试卷正文的章节大标题：多章节单文件时用来分节（除首章外先分页） */
    function chapterHeadingParagraph(text, pageBreak) {
        const { Paragraph, TextRun, AlignmentType, HeadingLevel, PageBreak } = docx;
        const parts = [];
        if (pageBreak)
            parts.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));
        parts.push(new Paragraph({
            children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: COLOR.title })],
            heading: HeadingLevel.HEADING_2,
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 200 },
        }));
        return parts;
    }
    /** 答案页 / 错题页里的章节小标题（不强制分页，与上级大标题区分） */
    function chapterLabelParagraph(text) {
        const { Paragraph, TextRun } = docx;
        return new Paragraph({
            children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: COLOR.title })],
            spacing: { before: 200, after: 120 },
        });
    }
    /** 题库导入格式：生成学习通智能导入兼容的 Word 文档 */
    async function buildBankImportBlob(title, sections) {
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = docx;
        const children = [];
        children.push(new Paragraph({
            children: [
                new TextRun({
                    text: title || '题库导入',
                    font: FONT,
                    size: 32,
                    bold: true,
                    color: COLOR.title,
                }),
            ],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { after: 240 },
        }));
        let qNum = 0;
        for (const section of sections) {
            for (const qtype of section.typeOrder) {
                const questions = section.results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                const prefix = BANK_TYPE_LABELS[qtype];
                for (const q of questions) {
                    qNum++;
                    // 题干（题号 + 题型标签 + 题干内容），括号归一化
                    const stemText = (0, legacy_bridge_4.formatRichForText)((0, legacy_bridge_4.questionContent)(q));
                    const stem = prefix + stemText.replace(/\(\s{2,}\)/g, '（ ）').replace(/（\s{2,}）/g, '（ ）');
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `${qNum}.${stem}`, font: FONT, size: 22 })],
                        spacing: { after: 40 },
                    }));
                    // 题干中的图片
                    const stemContent = (0, legacy_bridge_4.questionContent)(q);
                    for (const part of stemContent) {
                        if (part.type === 'image') {
                            const imgRun = await buildImageRun(part.url);
                            if (imgRun) {
                                children.push(new Paragraph({
                                    children: [imgRun],
                                    spacing: { after: 80 },
                                }));
                            }
                        }
                    }
                    // 选项
                    const options = q.options || [];
                    for (const opt of options) {
                        const runs = await buildRichRuns((0, legacy_bridge_4.optionContent)(opt), `${opt.letter}. `);
                        children.push(new Paragraph({
                            children: runs,
                            indent: { left: 400, hanging: 200 },
                            spacing: { after: 40 },
                        }));
                    }
                    // 答案
                    const answer = (0, legacy_bridge_4.formatRichForText)((0, legacy_bridge_4.answerContent)(q)).trim();
                    if (answer) {
                        // 格式化答案内容
                        let formattedAnswerParts = null;
                        if (qtype === '多选') {
                            const parts = answer.replace(/\s+/g, '').split('');
                            formattedAnswerParts = [{ type: 'text', text: parts.join('，') }];
                        }
                        else if (qtype === '判断') {
                            if (/^[√✓Tt]|正确|True|TRUE/.test(answer)) {
                                formattedAnswerParts = [{ type: 'text', text: '对' }];
                            }
                            else {
                                formattedAnswerParts = [{ type: 'text', text: '错' }];
                            }
                        }
                        children.push(...(await buildRichParagraphs(formattedAnswerParts || [{ type: 'text', text: answer }], '答案：', 120)));
                    }
                    else {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: '', font: FONT, size: 22 })],
                            spacing: { after: 120 },
                        }));
                    }
                }
            }
        }
        const doc = new Document({
            styles: {
                paragraphStyles: [
                    {
                        id: 'Normal',
                        name: 'Normal',
                        run: { font: FONT, size: 22, color: COLOR.title },
                        paragraph: { spacing: { line: 330, lineRule: 'auto' } },
                    },
                ],
            },
            sections: bankImportSections(children),
        });
        return await Packer.toBlob(doc);
    }
    /**
     * 试卷正文：标题 + 统计摘要 + 各章节题目（按题型分组）。
     *
     * title 是文档大标题（多章节时是「课程 - N 个章节」这类聚合标题），
     * 与 section.title（章节名，多章节时作为正文里的分节标题）不是一回事。
     */
    async function buildPaperBody(title, sections) {
        const { Paragraph, TextRun, AlignmentType, HeadingLevel } = docx;
        const children = [];
        const multiSection = sections.length > 1;
        // 标题
        children.push(new Paragraph({
            children: [
                new TextRun({ text: title || '试卷', font: FONT, size: 32, bold: true, color: '2563EB' }),
            ],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 120 },
        }));
        // 统计摘要（多章节时跨章节汇总）
        let totalQ = 0;
        const typeCount = {};
        for (const section of sections) {
            for (const qtype of section.typeOrder) {
                const questions = section.results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                typeCount[qtype] = (typeCount[qtype] ?? 0) + questions.length;
                totalQ += questions.length;
            }
        }
        const summary = Object.entries(typeCount)
            .map(([type, count]) => `${type} ${count} 道`)
            .join(' / ');
        children.push(new Paragraph({
            children: [
                new TextRun({ text: `共 ${totalQ} 道题`, font: FONT, size: 20, color: COLOR.muted }),
                ...(summary
                    ? [new TextRun({ text: ` · ${summary}`, font: FONT, size: 20, color: COLOR.muted })]
                    : []),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 260 },
        }));
        for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
            const section = sections[sectionIndex];
            if (!section)
                continue;
            // 多章节单文件：每章前插章节大标题
            if (multiSection)
                children.push(...chapterHeadingParagraph(section.title, sectionIndex > 0));
            // 题号与题型序号都按章重新开始：每章从第 1 题、从「一、」数起
            const typeHeaders = buildTypeHeaders(section);
            let qNum = 0;
            for (const qtype of section.typeOrder) {
                const questions = section.results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                const header = typeHeaders.get(qtype) ?? TYPE_NAMES[qtype] ?? qtype;
                const count = questions.length;
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `${header}（本大题共${count}小题）`,
                            font: FONT,
                            size: SIZE_TYPE_HEADER,
                            bold: true,
                            color: COLOR.title,
                        }),
                    ],
                    spacing: { after: 120 },
                }));
                for (const q of questions) {
                    qNum++;
                    const stemContent = (0, legacy_bridge_4.questionContent)(q);
                    const typeMetaRun = q.typeMeta
                        ? new TextRun({ text: `${q.typeMeta} `, font: FONT, size: 22, color: COLOR.type })
                        : null;
                    if (qtype === '单选' || qtype === '多选') {
                        // 题目块
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `${qNum}. `,
                                    font: FONT,
                                    size: 22,
                                    bold: true,
                                    color: COLOR.title,
                                }),
                                ...(typeMetaRun ? [typeMetaRun] : []),
                                ...(await buildRichRuns(stemContent)),
                            ],
                            spacing: { before: 220, after: 100 },
                        }));
                        const options = q.options || [];
                        if (options.length > 0) {
                            // 单选/多选选项统一一列多行排列
                            for (const opt of options) {
                                children.push(new Paragraph({
                                    children: await buildRichRuns((0, legacy_bridge_4.optionContent)(opt), `${opt.letter}. `),
                                    indent: { left: 620, hanging: 220 },
                                    spacing: { before: 30, after: 30 },
                                }));
                            }
                            children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
                        }
                    }
                    else if (qtype === '填空') {
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `${qNum}. `,
                                    font: FONT,
                                    size: 22,
                                    bold: true,
                                    color: COLOR.title,
                                }),
                                ...(await buildRichRuns(stemContent)),
                            ],
                            spacing: { before: 220, after: 40 },
                        }));
                        children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
                    }
                    else if (qtype === '判断') {
                        const judgeRuns = await buildRichRuns(stemContent);
                        judgeRuns.push(new TextRun({ text: '（  ）', font: FONT, size: 22 }));
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `${qNum}. `,
                                    font: FONT,
                                    size: 22,
                                    bold: true,
                                    color: COLOR.title,
                                }),
                                ...(typeMetaRun ? [typeMetaRun] : []),
                                ...judgeRuns,
                            ],
                            spacing: { before: 220, after: 120 },
                        }));
                    }
                    else if (qtype === '简答') {
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `${qNum}. `,
                                    font: FONT,
                                    size: 22,
                                    bold: true,
                                    color: COLOR.title,
                                }),
                                ...(typeMetaRun ? [typeMetaRun] : []),
                                ...(await buildRichRuns(stemContent)),
                            ],
                            spacing: { before: 220, after: 40 },
                        }));
                        for (let i = 0; i < 8; i++) {
                            children.push(new Paragraph({
                                children: [new TextRun({ text: '', font: FONT, size: 22 })],
                                spacing: { after: 40 },
                            }));
                        }
                        children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
                    }
                }
                // 判断题与简答题之间补一行空行
                if (qtype === '判断') {
                    children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
                }
            }
        }
        return children;
    }
    /** 答案页：分页后按章节、按题型列出正确答案 */
    async function appendAnswerPage(children, sections) {
        const { Paragraph, TextRun, AlignmentType, HeadingLevel, PageBreak } = docx;
        const multiSection = sections.length > 1;
        children.push(new Paragraph({
            children: [new PageBreak()],
            spacing: { after: 0 },
        }));
        children.push(new Paragraph({
            children: [
                new TextRun({
                    text: '正确答案',
                    font: FONT,
                    size: 32,
                    bold: true,
                    color: COLOR.title,
                }),
            ],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 240 },
        }));
        for (const section of sections) {
            if (multiSection)
                children.push(chapterLabelParagraph(section.title));
            // 答案编号与题目页一致：按章重新从 1 数起，题型序号同样按本章重排
            const typeHeaders = buildTypeHeaders(section);
            let aNum = 0;
            for (const qtype of section.typeOrder) {
                const questions = section.results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                const header = typeHeaders.get(qtype) ?? TYPE_NAMES[qtype] ?? qtype;
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: header,
                            font: FONT,
                            size: SIZE_TYPE_HEADER,
                            bold: true,
                            color: COLOR.title,
                        }),
                    ],
                    spacing: { after: 120 },
                }));
                for (const q of questions) {
                    aNum++;
                    // 清洗答案文本，去除“正确答案:/我的答案:”等标签组合
                    const answerContentArr = (0, legacy_bridge_4.answerContent)(q);
                    const cleanedContent = answerContentArr.map((part) => {
                        if (part.type === 'text') {
                            return { ...part, text: cleanAnswerText(part.text) };
                        }
                        return part;
                    });
                    const answerRuns = await buildRichRuns(cleanedContent);
                    if (answerRuns.length === 0) {
                        answerRuns.push(new TextRun({ text: '（未找到答案）', font: FONT, size: 22, color: COLOR.muted }));
                    }
                    else {
                        // 将答案文本改为绿色加粗
                        answerRuns.forEach((run) => {
                            if (run.font)
                                run.font = FONT;
                            run.size = 22;
                            run.bold = true;
                            run.color = COLOR.answer;
                        });
                    }
                    children.push(new Paragraph({
                        children: [
                            new TextRun({ text: `${aNum}. `, font: FONT, size: 22, color: COLOR.title }),
                            ...answerRuns,
                        ],
                        spacing: { before: 90, after: 70 },
                        indent: { left: 420 },
                    }));
                    // 简答题答案常有多行，之间空一行便于区分
                    if (qtype === '简答') {
                        children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
                    }
                }
            }
        }
    }
    /** 错题汇总页：分页后列出答错的题目、我的答案与正确答案 */
    async function appendWrongPage(children, sections) {
        const { Paragraph, TextRun, AlignmentType, HeadingLevel, PageBreak } = docx;
        const multiSection = sections.length > 1;
        const hasAnyWrong = sections.some((section) => section.typeOrder.some((qtype) => (section.results[qtype] ?? []).some((question) => question.isWrong)));
        if (!hasAnyWrong)
            return;
        children.push(new Paragraph({
            children: [new PageBreak()],
            spacing: { after: 0 },
        }));
        children.push(new Paragraph({
            children: [
                new TextRun({
                    text: '错题汇总',
                    font: FONT,
                    size: 32,
                    bold: true,
                    color: COLOR.title,
                }),
            ],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 240 },
        }));
        for (const section of sections) {
            if (multiSection) {
                // 该章一道错题都没有时，连章节标题一起省掉（与「无错题的题型不出现」一致）
                const sectionHasWrong = section.typeOrder.some((qtype) => (section.results[qtype] ?? []).some((question) => question.isWrong));
                const before = children.length;
                if (sectionHasWrong)
                    children.push(chapterLabelParagraph(section.title));
                // 章节标题先占位，若无错题则回滚（保持既有「空section不出标题」的写法简单）
                if (!sectionHasWrong && children.length !== before)
                    children.length = before;
            }
            // 错题题号要能对上题目页，因此同样按章重新从 1 数起
            const typeHeaders = buildTypeHeaders(section);
            let globalNum = 0;
            for (const qtype of section.typeOrder) {
                const questions = section.results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                if (qtype === '简答') {
                    globalNum += questions.length;
                    continue;
                }
                const header = typeHeaders.get(qtype) ?? TYPE_NAMES[qtype] ?? qtype;
                let sectionHasWrong = false;
                for (const q of questions) {
                    if (q.isWrong) {
                        sectionHasWrong = true;
                        break;
                    }
                }
                if (!sectionHasWrong) {
                    globalNum += questions.length;
                    continue;
                }
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: header,
                            font: FONT,
                            size: SIZE_TYPE_HEADER,
                            bold: true,
                            color: COLOR.title,
                        }),
                    ],
                    spacing: { after: 120 },
                }));
                for (const q of questions) {
                    globalNum++;
                    if (!q.isWrong)
                        continue;
                    children.push(new Paragraph({
                        children: await buildRichRuns((0, legacy_bridge_4.questionContent)(q), `${globalNum}. `),
                        spacing: { before: 160, after: 40 },
                    }));
                    const options = q.options || [];
                    if (options.length > 0) {
                        for (const opt of options) {
                            children.push(new Paragraph({
                                children: await buildRichRuns((0, legacy_bridge_4.optionContent)(opt), `${opt.letter}. `),
                                indent: { left: 620, hanging: 220 },
                                spacing: { before: 30, after: 30 },
                            }));
                        }
                    }
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `我的答案: ${q.myAnswer || '无'}`,
                                font: FONT,
                                size: 22,
                                color: COLOR.wrong,
                            }),
                        ],
                        spacing: { before: 60, after: 40 },
                        indent: { left: 420 },
                    }));
                    const answerContentArr = (0, legacy_bridge_4.answerContent)(q);
                    const cleanedContent = answerContentArr.map((part) => {
                        if (part.type === 'text') {
                            return { ...part, text: cleanAnswerText(part.text) };
                        }
                        return part;
                    });
                    const correctRuns = await buildRichRuns(cleanedContent);
                    if (correctRuns.length === 0) {
                        correctRuns.push(new TextRun({ text: '（未找到答案）', font: FONT, size: 22, color: COLOR.muted }));
                    }
                    else {
                        correctRuns.forEach((run) => {
                            if (run.font)
                                run.font = FONT;
                            run.size = 22;
                            run.bold = true;
                            run.color = COLOR.answer;
                        });
                    }
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: '正确答案：',
                                font: FONT,
                                size: 22,
                                bold: true,
                                color: COLOR.answer,
                            }),
                            ...correctRuns,
                        ],
                        spacing: { before: 40, after: 120 },
                        indent: { left: 420 },
                    }));
                    // 每道题之间空一行
                    children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
                }
                // 每个类别之间空一行
                children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
            }
        }
    }
    async function generateWordBlob(title, sections, options) {
        const { Document, Packer } = docx;
        if (options.bankImport) {
            return await buildBankImportBlob(title, sections);
        }
        const children = await buildPaperBody(title, sections);
        if (options.withAnswers)
            await appendAnswerPage(children, sections);
        if (options.withWrong)
            await appendWrongPage(children, sections);
        const doc = new Document({
            styles: {
                paragraphStyles: [
                    {
                        id: 'Normal',
                        name: 'Normal',
                        run: { font: FONT, size: 22, color: COLOR.title },
                        paragraph: { spacing: { line: 330, lineRule: 'auto' } },
                    },
                ],
            },
            sections: [
                {
                    properties: {
                        page: {
                            size: { width: 11906, height: 16838 },
                            margin: {
                                top: 1134,
                                right: 1134,
                                bottom: 1134,
                                left: 1134,
                            },
                        },
                    },
                    children,
                },
            ],
        });
        return await Packer.toBlob(doc);
    }
});
define("exporters/export-service", ["require", "exports", "exporters/legacy-bridge", "exporters/text-formatter", "exporters/markdown-formatter", "exporters/word-exporter", "exporters/word-plan"], function (require, exports, legacy_bridge_5, text_formatter_1, markdown_formatter_1, word_exporter_1, word_plan_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ExportService = void 0;
    // 原版 getOutputText：根据选项生成输出文本
    // （打乱时题目和答案用打乱顺序重新生成，错题汇总始终用原始顺序）
    function getOutputText(legacy, options) {
        const fmt = options.format;
        const withAnswers = options.withAnswers;
        const withWrong = options.withWrong;
        const doShuffle = options.shuffle;
        let base = '';
        if (doShuffle) {
            const shuffled = (0, legacy_bridge_5.shuffleQuestions)(legacy.results, legacy.typeOrder);
            if (fmt === 'md') {
                base = withAnswers
                    ? (0, markdown_formatter_1.formatOutputWithAnswersMD)(shuffled, legacy.typeOrder)
                    : (0, markdown_formatter_1.formatOutputMD)(shuffled, legacy.typeOrder);
            }
            else {
                base = withAnswers
                    ? (0, text_formatter_1.formatOutputWithAnswers)(shuffled, legacy.typeOrder)
                    : (0, text_formatter_1.formatOutput)(shuffled, legacy.typeOrder);
            }
        }
        else {
            if (fmt === 'md') {
                base = withAnswers
                    ? (0, markdown_formatter_1.formatOutputWithAnswersMD)(legacy.results, legacy.typeOrder)
                    : (0, markdown_formatter_1.formatOutputMD)(legacy.results, legacy.typeOrder);
            }
            else {
                base = withAnswers
                    ? (0, text_formatter_1.formatOutputWithAnswers)(legacy.results, legacy.typeOrder)
                    : (0, text_formatter_1.formatOutput)(legacy.results, legacy.typeOrder);
            }
        }
        if (withWrong) {
            const wrong = fmt === 'md'
                ? (0, markdown_formatter_1.formatWrongQuestionsMD)(legacy.results, legacy.typeOrder)
                : (0, text_formatter_1.formatWrongQuestionsTXT)(legacy.results, legacy.typeOrder);
            return base + wrong;
        }
        return base;
    }
    // 原版 getChapterText：为单个章节生成输出文本（用于分章节下载）
    function getChapterText(chapter, options) {
        const results = chapter.results || {};
        const typeOrder = chapter.typeOrder || [];
        const withAnswers = options.withAnswers;
        const withWrong = options.withWrong;
        const doShuffle = options.shuffle;
        const activeResults = doShuffle ? (0, legacy_bridge_5.shuffleQuestions)(results, typeOrder) : results;
        if (options.format === 'md') {
            if (withWrong)
                return (0, markdown_formatter_1.formatWrongQuestionsMD)(activeResults, typeOrder);
            if (withAnswers)
                return (0, markdown_formatter_1.formatOutputWithAnswersMD)(activeResults, typeOrder);
            return (0, markdown_formatter_1.formatOutputMD)(activeResults, typeOrder);
        }
        if (withWrong)
            return (0, text_formatter_1.formatWrongQuestionsTXT)(activeResults, typeOrder);
        if (withAnswers)
            return (0, text_formatter_1.formatOutputWithAnswers)(activeResults, typeOrder);
        return (0, text_formatter_1.formatOutput)(activeResults, typeOrder);
    }
    class ExportService {
        async createArtifacts(sourceResult, options) {
            const legacy = (0, legacy_bridge_5.toLegacyResult)(sourceResult);
            // 原版规则：文件名输入去掉扩展名后作为基础名
            const baseFilename = (options.filename || legacy.title || '学习通题目').replace(/\.(txt|md|docx)$/, '');
            // Word 导出（原版规则：题库导入必带答案并禁用打乱/错题）
            // 文件划分交给 word-plan：勾选「按章节拆分文件」时一章一个 .docx；
            // 否则只产出一个文件，多章节时正文按「章节 → 题型」分节，不再把各章同题型混在一起。
            if (options.format === 'word') {
                const isBankImport = options.bankImport;
                const withWrong = !isBankImport && options.withWrong;
                const plans = (0, word_plan_1.planWordDocuments)(sourceResult, options);
                const artifacts = [];
                window.__xxt_failed_image_count = 0;
                for (const plan of plans) {
                    const failedBefore = window.__xxt_failed_image_count || 0;
                    const blob = await (0, word_exporter_1.generateWordBlob)(plan.title, plan.sections, {
                        withAnswers: isBankImport || options.withAnswers,
                        withWrong,
                        bankImport: isBankImport,
                    });
                    artifacts.push({
                        blob,
                        filename: plan.filename,
                        mimeType: word_exporter_1.WORD_MIME,
                        // 图片计数是全局累加的，这里取本文件的增量
                        failedImages: (window.__xxt_failed_image_count || 0) - failedBefore,
                    });
                }
                return artifacts;
            }
            // TXT / MD 导出
            const ext = options.format === 'md' ? '.md' : '.txt';
            const mime = options.format === 'md' ? 'text/markdown;charset=utf-8' : 'text/plain;charset=utf-8';
            // 分章节下载：每个章节单独一个文件
            if (options.splitByChapter && sourceResult.chapters && sourceResult.chapters.length > 1) {
                const artifacts = [];
                for (let i = 0; i < sourceResult.chapters.length; i++) {
                    const chapter = sourceResult.chapters[i];
                    if (!chapter)
                        continue;
                    const chapterData = (0, legacy_bridge_5.toLegacyChapter)(chapter);
                    // 生成单章文本
                    const chText = getChapterText(chapterData, options);
                    const safeTitle = (chapterData.title || `章节${i + 1}`).replace(/[\\/:*?"<>|]/g, '_');
                    artifacts.push({
                        blob: new Blob([chText], { type: mime }),
                        filename: `${baseFilename}_${i + 1}_${safeTitle}${ext}`,
                        mimeType: mime,
                    });
                }
                return artifacts;
            }
            // 合并下载
            const text = getOutputText(legacy, options);
            return [
                {
                    blob: new Blob([text], { type: mime }),
                    filename: `${baseFilename}${ext}`,
                    mimeType: mime,
                },
            ];
        }
        previewText(result, options) {
            // 原版行为：Word 格式不支持复制文本
            if (options.format === 'word')
                return '';
            return getOutputText((0, legacy_bridge_5.toLegacyResult)(result), options);
        }
    }
    exports.ExportService = ExportService;
});
define("infrastructure/clipboard-service", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ClipboardService = void 0;
    class ClipboardService {
        async writeText(value) {
            if (navigator.clipboard?.writeText) {
                await navigator.clipboard.writeText(value);
                return;
            }
            const textarea = document.createElement('textarea');
            textarea.value = value;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            const copied = document.execCommand('copy');
            textarea.remove();
            if (!copied)
                throw new Error('浏览器拒绝访问剪贴板');
        }
    }
    exports.ClipboardService = ClipboardService;
});
define("infrastructure/download-service", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.DownloadService = void 0;
    class DownloadService {
        // 原版下载方式：创建临时 <a> 点击后立即释放 URL
        async download(artifact) {
            const url = URL.createObjectURL(artifact.blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = artifact.filename;
            a.click();
            URL.revokeObjectURL(url);
        }
        // 多文件下载间隔 300ms，避免多个下载被浏览器拦截
        async downloadMany(artifacts) {
            for (const artifact of artifacts) {
                await this.download(artifact);
                // 浏览器下载间隔，避免多个下载被拦截
                await new Promise((resolve) => setTimeout(resolve, 300));
            }
        }
    }
    exports.DownloadService = DownloadService;
});
define("utils/async", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.delay = delay;
    exports.debounce = debounce;
    function delay(milliseconds) {
        return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
    }
    function debounce(callback, waitMs) {
        let timer = null;
        return (...arguments_) => {
            if (timer !== null)
                window.clearTimeout(timer);
            timer = window.setTimeout(() => {
                timer = null;
                callback(...arguments_);
            }, waitMs);
        };
    }
});
define("infrastructure/logger", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.Logger = void 0;
    class Logger {
        constructor(prefix = '[Chaoxing Work Export]') {
            this.prefix = prefix;
        }
        debug(message, details) {
            try {
                if (window.localStorage.getItem('chaoxing-work-export:debug') !== '1')
                    return;
            }
            catch {
                return;
            }
            console.debug(this.prefix, message, details ?? '');
        }
        warn(message, details) {
            console.warn(this.prefix, message, details ?? '');
        }
        error(message, details) {
            console.error(this.prefix, message, details ?? '');
        }
    }
    exports.Logger = Logger;
});
define("infrastructure/frame-bridge", ["require", "exports", "extractors/composite-extractor", "utils/async", "utils/hash", "infrastructure/logger"], function (require, exports, composite_extractor_2, async_1, hash_2, logger_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.FrameAgent = exports.FrameBridge = void 0;
    const MESSAGE_TYPE = 'chaoxing-work-export:frame-result';
    const REQUEST_TYPE = 'chaoxing-work-export:frame-request';
    const PROTOCOL_VERSION = 1;
    function isRecord(value) {
        return typeof value === 'object' && value !== null;
    }
    function isFrameRequest(value) {
        return (isRecord(value) &&
            value.type === REQUEST_TYPE &&
            value.protocolVersion === PROTOCOL_VERSION);
    }
    function isFrameMessage(value) {
        if (!isRecord(value))
            return false;
        const result = value.result;
        return (value.type === MESSAGE_TYPE &&
            value.protocolVersion === PROTOCOL_VERSION &&
            typeof value.frameUrl === 'string' &&
            typeof value.sentAt === 'number' &&
            isRecord(result) &&
            Array.isArray(result.questions) &&
            result.questions.length <= 10000 &&
            typeof result.title === 'string' &&
            typeof result.sourceUrl === 'string' &&
            typeof result.extractedAt === 'string');
    }
    function isTrustedOrigin(origin) {
        if (origin === window.location.origin)
            return true;
        try {
            const hostname = new URL(origin).hostname.toLowerCase();
            return hostname === 'chaoxing.com' || hostname.endsWith('.chaoxing.com');
        }
        catch {
            return false;
        }
    }
    class FrameBridge {
        constructor() {
            this.results = new Map();
            this.onMessage = (event) => {
                if (!isTrustedOrigin(event.origin) || !isFrameMessage(event.data))
                    return;
                this.results.set(event.data.frameUrl, {
                    receivedAt: Date.now(),
                    frameUrl: event.data.frameUrl,
                    result: event.data.result,
                });
            };
        }
        start() {
            window.addEventListener('message', this.onMessage);
        }
        stop() {
            window.removeEventListener('message', this.onMessage);
        }
        latest(after = 0) {
            return ([...this.results.values()]
                .filter((entry) => entry.receivedAt >= after)
                .sort((left, right) => right.receivedAt - left.receivedAt)[0] ?? null);
        }
        requestRefresh(root = document) {
            const request = { type: REQUEST_TYPE, protocolVersion: PROTOCOL_VERSION };
            const visited = new Set();
            const visit = (current, depth) => {
                if (visited.has(current) || depth > 5)
                    return;
                visited.add(current);
                current.querySelectorAll('iframe').forEach((frame) => {
                    try {
                        frame.contentWindow?.postMessage(request, '*');
                        if (frame.contentDocument)
                            visit(frame.contentDocument, depth + 1);
                    }
                    catch {
                        // Posting to a cross-origin WindowProxy is allowed; reading its document is not.
                    }
                });
            };
            visit(root, 0);
        }
        clear() {
            this.results.clear();
        }
    }
    exports.FrameBridge = FrameBridge;
    class FrameAgent {
        constructor(extractor = new composite_extractor_2.CompositeExtractor()) {
            this.extractor = extractor;
            this.observer = null;
            this.lastFingerprint = '';
            this.logger = new logger_1.Logger();
            this.onRequest = (event) => {
                if (event.source !== window.parent || !isTrustedOrigin(event.origin) || !isFrameRequest(event.data))
                    return;
                this.sendResult(true);
            };
            this.sendDebounced = (0, async_1.debounce)(() => this.sendResult(), 350);
        }
        start() {
            window.addEventListener('message', this.onRequest);
            this.sendResult();
            if (!document.body) {
                document.addEventListener('DOMContentLoaded', () => this.observe(), { once: true });
                return;
            }
            this.observe();
        }
        stop() {
            window.removeEventListener('message', this.onRequest);
            this.observer?.disconnect();
            this.observer = null;
        }
        observe() {
            if (!document.body)
                return;
            this.observer?.disconnect();
            this.observer = new MutationObserver(() => this.sendDebounced());
            this.observer.observe(document.body, { childList: true, subtree: true, characterData: true });
            window.setTimeout(() => this.sendResult(), 1000);
            window.setTimeout(() => this.sendResult(), 3000);
        }
        sendResult(force = false) {
            const result = this.extractor.extract(document);
            if (!result)
                return;
            const fingerprint = (0, hash_2.stableHash)(JSON.stringify(result.questions.map((question) => ({
                id: question.id,
                correctAnswer: question.correctAnswer,
                userAnswer: question.userAnswer,
                analysis: question.analysis,
                isWrong: question.isWrong,
            }))));
            if (!force && fingerprint === this.lastFingerprint)
                return;
            this.lastFingerprint = fingerprint;
            const message = {
                type: MESSAGE_TYPE,
                protocolVersion: PROTOCOL_VERSION,
                frameUrl: window.location.href,
                sentAt: Date.now(),
                result,
            };
            try {
                window.parent.postMessage(message, this.parentOrigin());
            }
            catch (error) {
                this.logger.debug('Unable to post iframe result', error);
            }
        }
        parentOrigin() {
            try {
                const origin = new URL(document.referrer).origin;
                return isTrustedOrigin(origin) ? origin : '*';
            }
            catch {
                return '*';
            }
        }
    }
    exports.FrameAgent = FrameAgent;
});
define("infrastructure/safe-storage", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.SafeStorage = void 0;
    class SafeStorage {
        constructor(storage = window.localStorage) {
            this.storage = storage;
        }
        read(key) {
            try {
                const value = this.storage.getItem(key);
                return value === null ? null : JSON.parse(value);
            }
            catch {
                return null;
            }
        }
        write(key, value) {
            try {
                this.storage.setItem(key, JSON.stringify(value));
                return true;
            }
            catch {
                return false;
            }
        }
        remove(key) {
            try {
                this.storage.removeItem(key);
            }
            catch {
                // Storage can be unavailable in privacy modes. Failing silently keeps extraction usable.
            }
        }
    }
    exports.SafeStorage = SafeStorage;
});
define("infrastructure/history-repository", ["require", "exports", "domain/question", "extractors/rich-content", "utils/hash", "infrastructure/safe-storage"], function (require, exports, question_2, rich_content_8, hash_3, safe_storage_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.HistoryRepository = void 0;
    const KEY = 'chaoxing-work-export:history:v3';
    const LEGACY_KEY = 'xxt_history';
    const MAX_ENTRIES = 10;
    function isRecord(value) {
        return typeof value === 'object' && value !== null && !Array.isArray(value);
    }
    function isHistoryEntry(value) {
        if (!isRecord(value) || !isRecord(value.result) || !isRecord(value.options))
            return false;
        return (typeof value.id === 'string' &&
            typeof value.createdAt === 'string' &&
            typeof value.title === 'string' &&
            typeof value.result.title === 'string' &&
            Array.isArray(value.result.questions) &&
            typeof value.options.format === 'string');
    }
    function legacyType(value) {
        const normalized = value.replace(/[【】()[\]（）\s]/gu, '');
        const map = {
            单选: 'single-choice',
            单选题: 'single-choice',
            单项选择题: 'single-choice',
            多选: 'multiple-choice',
            多选题: 'multiple-choice',
            多项选择题: 'multiple-choice',
            填空: 'fill-blank',
            填空题: 'fill-blank',
            判断: 'true-false',
            判断题: 'true-false',
            简答: 'short-answer',
            简答题: 'short-answer',
        };
        return map[normalized] ?? null;
    }
    function legacyRich(value, fallback) {
        if (Array.isArray(value)) {
            const valid = value.filter((part) => {
                return isRecord(part) && ['text', 'image', 'break'].includes(String(part.type));
            });
            if (valid.length > 0)
                return valid;
        }
        if (typeof value === 'string')
            return (0, rich_content_8.textContent)(value);
        return typeof fallback === 'string' ? (0, rich_content_8.textContent)(fallback) : [];
    }
    function legacyOptions(value) {
        if (!Array.isArray(value))
            return [];
        return value.flatMap((option, index) => {
            if (typeof option === 'string') {
                const match = option.match(/^\s*([A-Z])\s*[.、．:：]?\s*(.*)$/iu);
                return [
                    {
                        key: match?.[1]?.toUpperCase() ?? String.fromCharCode(65 + index),
                        content: (0, rich_content_8.textContent)(match?.[2] ?? option),
                    },
                ];
            }
            if (!isRecord(option))
                return [];
            const keySource = option.key ?? option.letter ?? option.label;
            const key = typeof keySource === 'string' ? keySource : String.fromCharCode(65 + index);
            return [
                {
                    key: key.replace(/[.、．:：\s]/gu, '').toUpperCase(),
                    content: legacyRich(option.content ?? option.richContent, option.text ?? option.value),
                },
            ];
        });
    }
    function legacyPayload(entry) {
        return isRecord(entry.data) ? { ...entry, ...entry.data } : entry;
    }
    function extractLegacyQuestions(source) {
        const results = isRecord(source.results)
            ? source.results
            : isRecord(source.questionGroups)
                ? source.questionGroups
                : null;
        if (!results)
            return [];
        const questions = [];
        for (const [legacyLabel, rawQuestions] of Object.entries(results)) {
            const type = legacyType(legacyLabel);
            if (!type || !Array.isArray(rawQuestions))
                continue;
            rawQuestions.forEach((rawQuestion, index) => {
                if (!isRecord(rawQuestion))
                    return;
                const stem = legacyRich(rawQuestion.stemContent ?? rawQuestion.questionContent, rawQuestion.stem ?? rawQuestion.question);
                const options = legacyOptions(rawQuestion.options);
                const fingerprint = `${legacyLabel}|${index}|${JSON.stringify(stem)}`;
                questions.push({
                    id: `legacy-${(0, hash_3.stableHash)(fingerprint)}`,
                    number: typeof rawQuestion.qnum === 'number'
                        ? rawQuestion.qnum
                        : typeof rawQuestion.number === 'number'
                            ? rawQuestion.number
                            : undefined,
                    type,
                    typeMeta: typeof rawQuestion.typeMeta === 'string' ? rawQuestion.typeMeta : undefined,
                    stem,
                    options,
                    correctAnswer: legacyRich(rawQuestion.correctAnswerContent ?? rawQuestion.answerContent, rawQuestion.correctAnswer ?? rawQuestion.answer),
                    userAnswer: legacyRich(rawQuestion.myAnswerContent ?? rawQuestion.userAnswerContent, rawQuestion.myAnswer ?? rawQuestion.userAnswer),
                    analysis: legacyRich(rawQuestion.analysisContent, rawQuestion.analysis),
                    isWrong: rawQuestion.isWrong === true,
                    source: {
                        extractor: 'legacy-history',
                        pageUrl: typeof source.sourceUrl === 'string' ? source.sourceUrl : '',
                    },
                });
            });
        }
        return questions;
    }
    function migrateLegacyResult(entry) {
        const source = legacyPayload(entry);
        const rawChapters = source.chapterDataList;
        const chapters = [];
        if (Array.isArray(rawChapters)) {
            rawChapters.forEach((rawChapter, index) => {
                if (!isRecord(rawChapter))
                    return;
                const chapterSource = legacyPayload(rawChapter);
                const titleSource = chapterSource.title ?? chapterSource.chapterTitle ?? chapterSource.name;
                const title = typeof titleSource === 'string' ? titleSource : `第 ${index + 1} 章`;
                const idSource = chapterSource.id ?? chapterSource.chapterId;
                const id = typeof idSource === 'string' ? idSource : `legacy-chapter-${index + 1}`;
                const questions = extractLegacyQuestions(chapterSource).map((question, questionIndex) => ({
                    ...question,
                    id: `legacy-${(0, hash_3.stableHash)(`${id}|${question.id}|${questionIndex}`)}`,
                    chapterId: id,
                    chapterTitle: title,
                }));
                if (questions.length === 0)
                    return;
                chapters.push({
                    id,
                    title,
                    questions,
                    typeOrder: (0, question_2.deriveTypeOrder)(questions),
                    statistics: (0, question_2.buildStatistics)(questions),
                    sourceUrl: typeof chapterSource.sourceUrl === 'string' ? chapterSource.sourceUrl : '',
                    extractor: 'legacy-history',
                });
            });
        }
        const questions = chapters.length > 0
            ? chapters.flatMap((chapter) => chapter.questions)
            : extractLegacyQuestions(source);
        if (questions.length === 0)
            return null;
        const titleSource = source.title ?? entry.title;
        const title = typeof titleSource === 'string' ? titleSource : '历史题目';
        return {
            title,
            questions,
            typeOrder: (0, question_2.deriveTypeOrder)(questions),
            statistics: (0, question_2.buildStatistics)(questions),
            sourceUrl: typeof source.sourceUrl === 'string' ? source.sourceUrl : '',
            extractor: 'legacy-history',
            extractedAt: typeof source.extractedAt === 'string' ? source.extractedAt : new Date().toISOString(),
            chapters: chapters.length > 0 ? chapters : undefined,
        };
    }
    function migrateLegacyEntry(value) {
        if (!isRecord(value))
            return null;
        const payload = legacyPayload(value);
        const result = migrateLegacyResult(value);
        if (!result)
            return null;
        const rawFormat = payload.format ?? (isRecord(payload.options) ? payload.options.format : undefined);
        const format = rawFormat === 'txt' || rawFormat === 'md' ? rawFormat : 'word';
        const title = typeof payload.title === 'string' ? payload.title : result.title;
        const legacyOptions = isRecord(payload.options) ? payload.options : payload;
        const options = {
            format,
            filename: typeof legacyOptions.filename === 'string' ? legacyOptions.filename : title,
            withAnswers: legacyOptions.withAnswers === true,
            withWrong: legacyOptions.withWrong === true,
            shuffle: legacyOptions.shuffle === true,
            bankImport: legacyOptions.bankImport === true,
            splitByChapter: legacyOptions.splitByChapter === true,
        };
        const idSource = payload.id;
        const createdAtSource = payload.createdAt ?? payload.date;
        return {
            id: typeof idSource === 'string' ? idSource : `legacy-${Date.now()}-${(0, hash_3.stableHash)(title)}`,
            createdAt: typeof createdAtSource === 'string' ? createdAtSource : new Date().toISOString(),
            title,
            result,
            options,
            textSnapshot: typeof payload.outputText === 'string' ? payload.outputText : undefined,
        };
    }
    class HistoryRepository {
        constructor(storage = new safe_storage_1.SafeStorage()) {
            this.storage = storage;
        }
        load() {
            const current = this.storage.read(KEY);
            if (Array.isArray(current))
                return current.filter(isHistoryEntry).slice(0, MAX_ENTRIES);
            const legacy = this.storage.read(LEGACY_KEY);
            if (Array.isArray(legacy)) {
                const migrated = legacy
                    .map(migrateLegacyEntry)
                    .filter((entry) => entry !== null)
                    .slice(0, MAX_ENTRIES);
                return this.save(migrated);
            }
            return [];
        }
        add(entry) {
            const entries = [entry, ...this.load().filter((existing) => existing.id !== entry.id)].slice(0, MAX_ENTRIES);
            return this.save(entries);
        }
        remove(id) {
            return this.save(this.load().filter((entry) => entry.id !== id));
        }
        clear() {
            this.storage.remove(KEY);
        }
        save(entries) {
            let candidates = [...entries].slice(0, MAX_ENTRIES);
            while (candidates.length > 0) {
                if (this.storage.write(KEY, candidates))
                    return candidates;
                candidates = candidates.slice(0, -1);
            }
            this.storage.remove(KEY);
            return [];
        }
    }
    exports.HistoryRepository = HistoryRepository;
});
define("utils/shortcut", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.matchesShortcut = matchesShortcut;
    exports.formatShortcut = formatShortcut;
    exports.parseShortcut = parseShortcut;
    function matchesShortcut(event, shortcut) {
        return (event.ctrlKey === shortcut.ctrl &&
            event.shiftKey === shortcut.shift &&
            event.altKey === shortcut.alt &&
            event.key.toLowerCase() === shortcut.key.toLowerCase());
    }
    function formatShortcut(shortcut) {
        return [
            shortcut.ctrl ? 'Ctrl' : '',
            shortcut.shift ? 'Shift' : '',
            shortcut.alt ? 'Alt' : '',
            shortcut.key.length === 1 ? shortcut.key.toUpperCase() : shortcut.key,
        ]
            .filter(Boolean)
            .join('+');
    }
    function parseShortcut(value, fallback) {
        const parts = value
            .split('+')
            .map((part) => part.trim())
            .filter(Boolean);
        const lower = parts.map((part) => part.toLowerCase());
        const key = parts.find((part) => !['ctrl', 'control', 'shift', 'alt', 'option'].includes(part.toLowerCase()));
        if (!key)
            return { ...fallback };
        return {
            ctrl: lower.includes('ctrl') || lower.includes('control'),
            shift: lower.includes('shift'),
            alt: lower.includes('alt') || lower.includes('option'),
            key: key.toLowerCase(),
        };
    }
});
define("infrastructure/settings-repository", ["require", "exports", "domain/settings", "utils/shortcut", "infrastructure/safe-storage"], function (require, exports, settings_1, shortcut_1, safe_storage_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.SettingsRepository = void 0;
    const KEY = 'chaoxing-work-export:settings:v3';
    const LEGACY_KEY = 'xxt_settings';
    const LEGACY_EXPORT_KEYS = ['xxt_export_config', 'xxt_export_options'];
    // 「打开窗口自动提取」默认值升级的标记（只升级一次，之后尊重用户的显式选择）
    const AUTO_EXTRACT_DEFAULT_KEY = 'chaoxing-work-export:auto-extract-default:v1';
    function isRecord(value) {
        return typeof value === 'object' && value !== null && !Array.isArray(value);
    }
    function booleanValue(value, fallback) {
        return typeof value === 'boolean' ? value : fallback;
    }
    function shortcutValue(value, fallback) {
        if (typeof value === 'string')
            return (0, shortcut_1.parseShortcut)(value, fallback);
        if (!isRecord(value))
            return { ...fallback };
        return {
            ctrl: booleanValue(value.ctrl, fallback.ctrl),
            shift: booleanValue(value.shift, fallback.shift),
            alt: booleanValue(value.alt, fallback.alt),
            key: typeof value.key === 'string' && value.key ? value.key.toLowerCase() : fallback.key,
        };
    }
    function panelPositionValue(value) {
        if (!isRecord(value))
            return null;
        return typeof value.left === 'number' && typeof value.top === 'number'
            ? { left: value.left, top: value.top }
            : null;
    }
    function themeValue(value) {
        if (value === 'system')
            return 'auto';
        return value === 'light' || value === 'dark' || value === 'auto' ? value : 'auto';
    }
    function firstBoolean(source, keys, fallback) {
        for (const key of keys) {
            if (typeof source[key] === 'boolean')
                return source[key];
        }
        return fallback;
    }
    function exportPreferencesValue(value) {
        const source = isRecord(value) ? value : {};
        const rawFormat = source.format ?? source.exportFormat;
        const format = rawFormat === 'txt' || rawFormat === 'md' ? rawFormat : 'word';
        return {
            format,
            withAnswers: firstBoolean(source, ['withAnswers', 'includeAnswers', 'appendAnswers'], settings_1.DEFAULT_SETTINGS.exportPreferences.withAnswers),
            withWrong: firstBoolean(source, ['withWrong', 'includeWrong', 'appendWrong'], settings_1.DEFAULT_SETTINGS.exportPreferences.withWrong),
            shuffle: firstBoolean(source, ['shuffle', 'randomOrder'], settings_1.DEFAULT_SETTINGS.exportPreferences.shuffle),
            bankImport: firstBoolean(source, ['bankImport', 'questionBankFormat', 'bankMode'], settings_1.DEFAULT_SETTINGS.exportPreferences.bankImport),
            splitByChapter: firstBoolean(source, ['splitByChapter', 'splitChapters'], settings_1.DEFAULT_SETTINGS.exportPreferences.splitByChapter),
        };
    }
    function settingsValue(value) {
        const source = isRecord(value) ? value : {};
        const legacyExport = source.exportConfig ?? source.exportOptions;
        return {
            theme: themeValue(source.theme),
            shortcut: shortcutValue(source.shortcut ?? source.extractShortcut, settings_1.DEFAULT_SETTINGS.shortcut),
            hideShortcut: shortcutValue(source.hideShortcut ?? source.toggleShortcut, settings_1.DEFAULT_SETTINGS.hideShortcut),
            exportPreferences: exportPreferencesValue(source.exportPreferences ?? legacyExport ?? source),
            enableDrag: booleanValue(source.enableDrag ?? source.draggable, settings_1.DEFAULT_SETTINGS.enableDrag),
            rememberPanelPosition: booleanValue(source.rememberPanelPosition ?? source.rememberPosition, settings_1.DEFAULT_SETTINGS.rememberPanelPosition),
            panelPosition: panelPositionValue(source.panelPosition ?? source.panelPos),
            autoExtractOnLoad: booleanValue(source.autoExtractOnLoad, settings_1.DEFAULT_SETTINGS.autoExtractOnLoad),
        };
    }
    class SettingsRepository {
        constructor(storage = new safe_storage_2.SafeStorage()) {
            this.storage = storage;
        }
        load() {
            const current = this.storage.read(KEY);
            if (current !== null)
                return this.upgradeAutoExtractDefault(settingsValue(current));
            const legacySettings = this.storage.read(LEGACY_KEY);
            const legacyExport = LEGACY_EXPORT_KEYS.map((key) => this.storage.read(key)).find((value) => value !== null);
            if (legacySettings !== null || legacyExport !== undefined) {
                const merged = isRecord(legacySettings) ? { ...legacySettings } : {};
                if (legacyExport !== undefined)
                    merged.exportPreferences = legacyExport;
                const migrated = settingsValue(merged);
                this.save(migrated);
                return migrated;
            }
            return settingsValue(settings_1.DEFAULT_SETTINGS);
        }
        save(settings) {
            this.storage.write(KEY, settings);
        }
        /**
         * 一次性默认值升级：「打开窗口自动提取」的默认值由关闭改为打开。
         *
         * 只改 DEFAULT_SETTINGS 对存量配置无效 —— 该字段早就被显式写入过（值为旧的默认值 false），
         * 会一直盖住新的默认值。这里在首次加载时补成新默认值并记下标记；标记置位后
         * 用户在设置里主动关掉会被正常保存，不会再次被改回。
         */
        upgradeAutoExtractDefault(settings) {
            if (this.storage.read(AUTO_EXTRACT_DEFAULT_KEY) !== null)
                return settings;
            this.storage.write(AUTO_EXTRACT_DEFAULT_KEY, true);
            if (settings.autoExtractOnLoad)
                return settings;
            const upgraded = { ...settings, autoExtractOnLoad: true };
            this.save(upgraded);
            return upgraded;
        }
    }
    exports.SettingsRepository = SettingsRepository;
});
define("application/chapter-locator", ["require", "exports", "utils/text"], function (require, exports, text_6) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ChapterLocator = void 0;
    const CHAPTER_CONTAINER_SELECTORS = [
        '#coursetree',
        '.catalog_points',
        '.chapter-list',
        '.catalog-list',
        '.posCatalog_list',
    ];
    // 章节条目：优先取真正承载章节名的叶子节点（学习通真实结构是 span.posCatalog_name，
    // 它是唯一挂 onclick 的节点），容器 div 只作为兜底候选
    const CHAPTER_ITEM_SELECTORS = [
        '.posCatalog_name',
        '.catalog_name',
        '.chapter-item',
        '[data-chapter-id]',
        '[data-id][class*="catalog"]',
        '.posCatalog_select',
    ];
    // 分组标题：第一级目录项（真实结构为 .posCatalog_select.firstLayer > span.posCatalog_title），
    // 它们不是可提取章节，必须排除
    const CHAPTER_GROUP_SELECTORS = ['.firstLayer', '.posCatalog_title'];
    const ACTIVE_CLASSES = ['active', 'cur', 'current', 'on', 'selected', 'posCatalog_active'];
    /** Finds and activates course chapter entries without embedding extraction logic. */
    class ChapterLocator {
        constructor(root = document) {
            this.root = root;
        }
        list() {
            return this.elements().map((element, index) => ({
                index,
                id: this.idOf(element, index),
                title: this.titleOf(element, index),
                active: this.isActive(element),
            }));
        }
        activeIndex() {
            const index = this.elements().findIndex((element) => this.isActive(element));
            return index >= 0 ? index : null;
        }
        activate(index) {
            const element = this.elements()[index];
            if (!element)
                return false;
            const target = this.clickableTarget(element);
            target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
            target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
            target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
            target.click();
            return true;
        }
        elements() {
            const searchRoots = [this.root];
            for (const selector of CHAPTER_CONTAINER_SELECTORS) {
                const container = this.root.querySelector(selector);
                if (container)
                    searchRoots.unshift(container);
            }
            for (const searchRoot of searchRoots) {
                for (const selector of CHAPTER_ITEM_SELECTORS) {
                    const raw = [...searchRoot.querySelectorAll(selector)];
                    const filtered = raw.filter((element) => this.isChapterItem(element));
                    if (filtered.length > 1)
                        return this.removeNestedDuplicates(filtered);
                }
            }
            return [];
        }
        /** 是否为可提取章节：排除分组标题与不可见占位项 */
        isChapterItem(element) {
            // 分组标题（.firstLayer / .posCatalog_title）不是章节
            const isGroup = CHAPTER_GROUP_SELECTORS.some((selector) => element.matches(selector) || Boolean(element.closest(selector)));
            if (isGroup)
                return false;
            if (!this.isVisible(element))
                return false;
            const title = this.titleOf(element, 0);
            return title.length > 0 && !/^(?:目录|章节|返回)$/u.test(title);
        }
        /** 可见性判定：测试环境（linkedom）没有布局信息，此时不做隐藏判定 */
        isVisible(element) {
            if (typeof element.getClientRects !== 'function')
                return true;
            if (element.getClientRects().length > 0)
                return true;
            return Boolean(element.offsetParent);
        }
        /**
         * 选中真正可点击的节点。
         * 学习通把 onclick 挂在章节名节点（span.posCatalog_name）上，容器 div 自身没有事件，
         * 对容器派发点击不会向下传递到子节点，因此必须精确定位到带 onclick 的元素。
         */
        clickableTarget(element) {
            if (element.hasAttribute('onclick'))
                return element;
            const handler = element.querySelector('[onclick]');
            if (handler)
                return handler;
            if (element.matches('a, button, [role="button"]'))
                return element;
            return element.querySelector('a, button, [role="button"]') ?? element;
        }
        removeNestedDuplicates(elements) {
            return elements.filter((element) => {
                return !elements.some((other) => other !== element && other.contains(element));
            });
        }
        idOf(element, index) {
            return (element.dataset.chapterId ??
                element.dataset.id ??
                element.getAttribute('id') ??
                `chapter-${index + 1}`);
        }
        titleOf(element, index) {
            const explicit = element.getAttribute('title') ??
                element.querySelector('[title]')?.getAttribute('title') ??
                element.querySelector('.catalog_name, .chapter-title, .posCatalog_name')
                    ?.innerText ??
                element.innerText ??
                // 兜底：章节名节点不带 title 时也要取到真实名称，避免退化成「第 N 章」
                element.textContent;
            const title = (0, text_6.normalizeInlineWhitespace)(explicit ?? '');
            return title || `第 ${index + 1} 章`;
        }
        isActive(element) {
            if (element.getAttribute('aria-current') === 'true')
                return true;
            if (element.getAttribute('aria-selected') === 'true')
                return true;
            // 激活态类名挂在容器 div 上（class="posCatalog_select posCatalog_active"），
            // 而章节条目现在是它内部的名字节点，因此必须连同所属容器一起判断
            const owner = element.closest('.posCatalog_select');
            const nodes = owner && owner !== element ? [element, owner] : [element];
            if (nodes.some((node) => ACTIVE_CLASSES.some((className) => node.classList.contains(className)))) {
                return true;
            }
            return Boolean(element.closest(ACTIVE_CLASSES.map((className) => `.${className}`).join(',')));
        }
    }
    exports.ChapterLocator = ChapterLocator;
});
define("application/extraction-service", ["require", "exports", "domain/question", "extractors/composite-extractor", "extractors/page-context", "infrastructure/frame-bridge", "utils/async", "utils/hash", "utils/dom"], function (require, exports, question_3, composite_extractor_3, page_context_2, frame_bridge_1, async_2, hash_4, dom_8) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ExtractionService = void 0;
    exports.incompleteChoiceCount = incompleteChoiceCount;
    // 采样间隔：只决定「题目出现」被检测到的延迟，越小越快（每轮含一次跨层提取）
    const QUESTION_SETTLE_INTERVAL_MS = 120;
    const DEFAULT_QUESTION_TIMEOUT_MS = 12000;
    // 结果需连续多帧完全一致才判定「渲染完毕」：答题页会分几批渲染题目，
    // 只比对两帧会在两批之间误判为结束（表现为只拿到前面几道题）
    const SETTLE_STABLE_MS = 600;
    // 选择题确认渲染不出选项时的兜底：稳定这么久后按现状返回，避免空等超时
    const PARTIAL_SETTLE_MS = 3000;
    // 判断「文档里是否已经有题目结构」用：只要命中就说明该文档可能还有选项在路上
    const QUESTION_CONTAINER_SELECTOR = '.TiMu, .questionLi, .answerBg, .mark_item, .mark_name';
    /** 需要选项才算「渲染完整」的题型 */
    const CHOICE_TYPES = new Set([
        'single-choice',
        'multiple-choice',
    ]);
    /**
     * 统计「渲染明显不完整」的选择题数量：选择题的题干与选项是同一批 HTML，
     * 只出现题干而没有选项，几乎只可能是页面还没渲染完。
     * 答题页中途被采样到时若直接采用，导出结果就会只剩题干。
     */
    function incompleteChoiceCount(result) {
        if (!result)
            return 0;
        return result.questions.filter((question) => CHOICE_TYPES.has(question.type) && question.options.length === 0).length;
    }
    /** 结果丰富度：题数优先，其次选项总数，用于在加载过程中保留最完整的一帧 */
    function richnessOf(result) {
        const options = result.questions.reduce((total, question) => total + question.options.length, 0);
        return result.questions.length * 1000 + options;
    }
    /**
     * Coordinates all extraction sources visible to the current top-level page.
     *
     * A Chaoxing assignment may render directly in the document, in a same-origin
     * iframe, or in a cross-origin iframe. The service deliberately keeps these
     * transport concerns out of individual DOM extractors.
     */
    class ExtractionService {
        constructor(extractor = new composite_extractor_3.CompositeExtractor(), frameBridge = new frame_bridge_1.FrameBridge()) {
            this.extractor = extractor;
            this.frameBridge = frameBridge;
        }
        extract() {
            return this.applyScopeTitle(this.selectBest([
                ...this.collectDocumentCandidates(document),
                ...this.collectFrameCandidate(this.frameBridge.latest()),
            ]));
        }
        /** Extracts only documents that can be accessed synchronously from this window. */
        extractAccessibleDocuments() {
            return this.applyScopeTitle(this.selectBest(this.collectDocumentCandidates(document)));
        }
        /** Extracts only the current document, without recursively reading iframes. */
        extractCurrentDocument(root = document) {
            return this.extractor.extract(root);
        }
        fingerprint(result) {
            if (!result)
                return 'empty';
            const identity = result.questions
                .map((question) => `${question.id}:${question.type}`)
                .join('|');
            // 标题不参与指纹：外层壳的章节名（.prev_title）会比题目 DOM 更早更新，
            // 若纳入指纹，chapter 提取会把上一章残留的题目误判成本章的新内容
            return (0, hash_4.stableHash)(`${result.sourceUrl}|${identity}`);
        }
        /**
         * 用外层壳（学生学习页面）的章节名覆盖标题。
         *
         * 章节页的题目在嵌套 iframe 里，嵌套文档只能解析出自身标题（「章节测验 待完成」），
         * 而调用方需要的是章节名，因此在此统一覆盖；无 #prev_title / .prev_title 时行为不变。
         */
        applyScopeTitle(result) {
            if (!result)
                return null;
            const scopeTitle = (0, page_context_2.resolvePrevTitle)(document);
            if (!scopeTitle || scopeTitle === result.title)
                return result;
            return { ...result, title: scopeTitle };
        }
        /**
         * 采样等待题目「渲染完整」，而不是等到第一眼看见题目就收工。
         *
         * 页面就绪的时刻并不确定：任务卡切换、答题页 AJAX 渲染都可能晚于我们的点击。
         * 「切一次卡 + 固定长等待」要么空等、要么在旧页面上点了空，因此改为按采样间隔反复确认，
         * 期间通过 repair 修正被新页面重置的状态（例如任务卡回到默认的「视频」卡）。
         *
         * 但只判断「有没有题目」并不够：答题页会分几批渲染，中途被采样到时可能只渲染出题干、
         * 选项还没补上。因此这里做了两件事：
         * 1. 结果要连续多帧完全一致才认为渲染结束（`SETTLE_STABLE_MS`）；
         * 2. 选择题缺选项视为「仍在渲染」，除非页面稳定很久仍是如此（`PARTIAL_SETTLE_MS` 兜底，
         *    避免个别确实没有选项的页面白等满整段超时）；若承载题目的文档还在 loading，
         *    说明 HTML 还没解析完，此时不允许按兜底提前返回，继续等（`deadline` 仍然兜底）；
         * 3. 全程记录「最丰富的一帧」，超时或兜底返回它，绝不返回比中间帧更差的结果。
         */
        async settleQuestions(options = {}) {
            const previousFingerprint = options.previousFingerprint;
            const intervalMs = options.intervalMs ?? QUESTION_SETTLE_INTERVAL_MS;
            const deadline = Date.now() + (options.timeoutMs ?? DEFAULT_QUESTION_TIMEOUT_MS);
            let best = null;
            let bestRichness = Number.NEGATIVE_INFINITY;
            let seenFingerprint = '';
            let stableMs = 0;
            for (;;) {
                options.repair?.();
                const result = this.extract();
                if (result && result.questions.length > 0) {
                    const fingerprint = this.fingerprint(result);
                    // 指纹未变化说明页面还停在上一个内容上（章节尚未切过去），这一类帧不参与判定
                    const changed = previousFingerprint === undefined || fingerprint !== previousFingerprint;
                    if (changed) {
                        const richness = richnessOf(result);
                        if (richness > bestRichness) {
                            best = result;
                            bestRichness = richness;
                        }
                        stableMs = fingerprint === seenFingerprint ? stableMs + intervalMs : 0;
                        seenFingerprint = fingerprint;
                        const incomplete = incompleteChoiceCount(result) > 0;
                        // 答题页是一整个大 HTML，浏览器按到达顺序解析：题干（.Zy_TItle，结构里在前）
                        // 可能已经入 DOM，而同一题的选项（ul.Zy_ulTop，结构里在后）还没解析到。
                        // 这种「题干有了、选项还在路上」的中间态不满足「稳定 3 秒就收工」的前提，
                        // 只要承载题目的文档还在 loading 就继续等（最终仍由 deadline 兜底）。
                        const waitingForOptions = incomplete && this.hasLoadingQuestionDocument();
                        if (stableMs >= SETTLE_STABLE_MS &&
                            (!incomplete || (stableMs >= PARTIAL_SETTLE_MS && !waitingForOptions))) {
                            return incomplete ? (best ?? result) : result;
                        }
                    }
                }
                if (Date.now() >= deadline)
                    return best;
                await (0, async_2.delay)(intervalMs);
            }
        }
        /**
         * 是否还有「承载题目的文档」处于加载中。
         *
         * 只有「选择题缺选项」的结果会用到它：分批到达的 HTML 会让题干先出现、选项后到，
         * 这时页面既没有变化也没有选项，但选项并非不会来，所以不能按「稳定 3 秒」收工。
         * 结果本身是完整的（选择题都有选项）时完全不参与判定，正常页面行为不变。
         */
        hasLoadingQuestionDocument() {
            for (const { document: current } of (0, dom_8.collectAccessibleDocuments)(document)) {
                if (current.readyState !== 'loading')
                    continue;
                try {
                    if (current.querySelector(QUESTION_CONTAINER_SELECTOR))
                        return true;
                }
                catch {
                    // 不可访问的文档直接跳过
                }
            }
            return false;
        }
        collectDocumentCandidates(root) {
            const candidates = [];
            // 跨层收集交给公共遍历工具，和任务卡/章节定位共用同一套 iframe 规则
            for (const { document: current, depth } of (0, dom_8.collectAccessibleDocuments)(root)) {
                try {
                    const result = this.extractor.extract(current);
                    if (result) {
                        candidates.push({
                            result,
                            sourcePriority: Math.max(1, 100 - depth * 10),
                            timestamp: Date.parse(result.extractedAt) || Date.now(),
                        });
                    }
                }
                catch {
                    // One malformed document must not prevent other frames from being inspected.
                }
            }
            return candidates;
        }
        collectFrameCandidate(envelope) {
            if (!envelope)
                return [];
            return [
                {
                    result: envelope.result,
                    sourcePriority: 95,
                    timestamp: envelope.receivedAt,
                },
            ];
        }
        selectBest(candidates) {
            if (candidates.length === 0)
                return null;
            const sorted = [...candidates].sort((left, right) => {
                const questionDifference = right.result.questions.length - left.result.questions.length;
                if (questionDifference !== 0)
                    return questionDifference;
                // 题数相同时优先选选项更完整的候选：同一份题目可能有「只渲染了题干」的中间态，
                // 若按来源优先级挑选，半渲染的候选会把带选项的完整候选顶掉
                const richnessDifference = richnessOf(right.result) - richnessOf(left.result);
                if (richnessDifference !== 0)
                    return richnessDifference;
                const priorityDifference = right.sourcePriority - left.sourcePriority;
                if (priorityDifference !== 0)
                    return priorityDifference;
                return right.timestamp - left.timestamp;
            });
            const best = sorted[0];
            if (!best)
                return null;
            // A page can expose the same question in nested accessible documents. Merge
            // candidates with the winning title only when their question identities differ.
            const questions = new Map();
            for (const candidate of sorted) {
                if (candidate.result.title !== best.result.title && best.result.questions.length > 0)
                    continue;
                for (const question of candidate.result.questions)
                    questions.set(question.id, question);
            }
            const merged = [...questions.values()];
            return {
                ...best.result,
                questions: merged,
                typeOrder: (0, question_3.deriveTypeOrder)(merged),
                statistics: (0, question_3.buildStatistics)(merged),
            };
        }
    }
    exports.ExtractionService = ExtractionService;
});
define("application/task-tab-locator", ["require", "exports", "utils/dom", "utils/text"], function (require, exports, dom_9, text_7) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.QuestionTabGuard = exports.TaskTabLocator = void 0;
    // 任务卡切换栏容器：知识卡片页顶部（视频 / 章节测验 / 作业 …）所在位置
    const TAB_BAR_SELECTORS = ['#prev_tab', '.prev_list', '.prev_ul'];
    // 任务卡条目：真实结构是 li[role="option"]，文案在 span.spanText 中
    const TAB_ITEM_SELECTORS = ['li[role="option"]', '.prev_ul li', 'li'];
    // 承载题目的任务卡关键词（章节测验 / 作业 / 考试 / 练习…）
    const QUESTION_TAB_PATTERN = /测验|作业|考试|练习|测试|自测/u;
    // 激活态类名：学习通实际使用 .active，其余作为兼容候选
    const ACTIVE_CLASSES = ['active', 'on', 'cur', 'current', 'selected'];
    // 补切任务卡的默认节流参数：单章最多补切几次、两次之间至少间隔多久。
    // 冷却时间只需覆盖「新页面重建任务卡栏」的短暂窗口，太长会让章节切换变慢。
    const DEFAULT_MAX_CLICKS = 4;
    const DEFAULT_CLICK_COOLDOWN_MS = 400;
    /**
     * 定位知识卡片页的「任务点切换栏」。
     *
     * 任务卡内容是懒加载的：卡片的真实地址写在内部 iframe 的 `_src` 上，只有点击该卡触发
     * `changeDisplayContent()` 之后才会写回 `src` 并加载内容。因此任务卡不切换时，
     * 页面里根本不存在题目 DOM —— 这正是「必须手动切到章节测验才能提取」的根因。
     */
    class TaskTabLocator {
        constructor(root = document) {
            this.root = root;
        }
        /** 承载任务卡栏的文档：知识卡片页通常位于同源 iframe 内 */
        holder() {
            for (const { document: candidate } of (0, dom_9.collectAccessibleDocuments)(this.root)) {
                if (TAB_BAR_SELECTORS.some((selector) => candidate.querySelector(selector)))
                    return candidate;
            }
            return null;
        }
        /** 任务卡栏是否已经渲染出来 */
        present() {
            return this.holder() !== null;
        }
        list() {
            return this.elements().map((element, index) => ({
                index,
                id: element.getAttribute('id') ?? `tab-${index + 1}`,
                title: this.titleOf(element),
                active: this.isActive(element),
            }));
        }
        activeIndex() {
            const index = this.elements().findIndex((element) => this.isActive(element));
            return index >= 0 ? index : null;
        }
        /** 题目类任务卡在列表中的位置（章节测验 / 作业 / 考试…） */
        questionTabIndex() {
            const index = this.elements().findIndex((element) => QUESTION_TAB_PATTERN.test(this.titleOf(element)));
            return index >= 0 ? index : null;
        }
        /**
         * 确保题目类任务卡处于激活状态。
         * 已激活返回 active；本次完成切换返回 switched；没有任务卡栏或没有题目卡返回 missing。
         */
        ensureQuestionTab() {
            const element = this.elements().find((candidate) => {
                return QUESTION_TAB_PATTERN.test(this.titleOf(candidate));
            });
            if (!element)
                return 'missing';
            if (this.isActive(element))
                return 'active';
            return this.click(element) ? 'switched' : 'missing';
        }
        activate(index) {
            const element = this.elements()[index];
            return element ? this.click(element) : false;
        }
        /** 触发任务卡自身的 onclick（学习通挂在 li 上，对应 changeDisplayContent()） */
        click(element) {
            element.scrollIntoView({ block: 'nearest', inline: 'nearest' });
            element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
            element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
            element.click();
            return true;
        }
        elements() {
            const holder = this.holder();
            if (!holder)
                return [];
            for (const barSelector of TAB_BAR_SELECTORS) {
                const bar = holder.querySelector(barSelector);
                if (!bar)
                    continue;
                for (const itemSelector of TAB_ITEM_SELECTORS) {
                    const items = [...bar.querySelectorAll(itemSelector)].filter((element) => {
                        return this.titleOf(element).length > 0;
                    });
                    if (items.length > 0)
                        return items;
                }
            }
            return [];
        }
        titleOf(element) {
            const candidates = [
                element.getAttribute('title'),
                element.querySelector('.spanText')?.textContent,
                // 兜底：直接读文本，并去掉左侧序号（结构为 <span class="num">2</span><span class="spanText">章节测验</span>）
                (element.textContent ?? '').replace(/^\d+\s*/u, ''),
            ];
            for (const candidate of candidates) {
                const title = (0, text_7.normalizeInlineWhitespace)(candidate ?? '');
                if (title)
                    return title;
            }
            return '';
        }
        isActive(element) {
            if (element.getAttribute('aria-current') === 'true')
                return true;
            if (element.getAttribute('aria-selected') === 'true')
                return true;
            if (ACTIVE_CLASSES.some((className) => element.classList.contains(className)))
                return true;
            return Boolean(element.closest(ACTIVE_CLASSES.map((className) => `.${className}`).join(',')));
        }
    }
    exports.TaskTabLocator = TaskTabLocator;
    /**
     * 任务卡补切节流器，配合「采样等待」使用。
     *
     * 章节切换会重建卡片页，新页面默认停在「视频」卡，只切一次的任务卡会随旧页面一起失效；
     * 而在页面加载过程中频繁点击又会触发重复加载。这里限定「最多补切几次 + 冷却多久」，
     * 让采样循环可以放心地每轮都尝试补切一次。
     */
    class QuestionTabGuard {
        constructor(tabs, maxClicks = DEFAULT_MAX_CLICKS, cooldownMs = DEFAULT_CLICK_COOLDOWN_MS) {
            this.tabs = tabs;
            this.maxClicks = maxClicks;
            this.cooldownMs = cooldownMs;
            this.clicks = 0;
            this.lastClickedAt = Number.NEGATIVE_INFINITY;
        }
        /** 需要时补切一次题目卡，返回本次是否真的点击了 */
        ensure(now = Date.now()) {
            if (this.clicks >= this.maxClicks)
                return false;
            if (now - this.lastClickedAt < this.cooldownMs)
                return false;
            if (this.tabs.ensureQuestionTab() !== 'switched')
                return false;
            this.clicks += 1;
            this.lastClickedAt = now;
            return true;
        }
    }
    exports.QuestionTabGuard = QuestionTabGuard;
});
define("application/chapter-extraction-service", ["require", "exports", "domain/question", "extractors/page-context", "utils/async", "utils/hash", "application/chapter-locator", "application/extraction-service", "application/task-tab-locator"], function (require, exports, question_4, page_context_3, async_3, hash_5, chapter_locator_1, extraction_service_1, task_tab_locator_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ChapterExtractionService = void 0;
    // 单章等待上限：章节切换后要依次等卡片页重建 → 任务卡切换 → 答题页加载
    const CHAPTER_SETTLE_TIMEOUT_MS = 20000;
    // 页面已停在目标章节、没有任务卡可切时，等待「选项中补上来」的额外上限
    const PARTIAL_FILL_TIMEOUT_MS = 5000;
    // 收尾还原任务卡前的短暂等待：章节页面正在重建，任务卡栏属于新页面
    const RESTORE_TAB_DELAY_MS = 400;
    /** Sequentially visits selected chapters and produces a single aggregate snapshot. */
    class ChapterExtractionService {
        constructor(extractionService, locator = new chapter_locator_1.ChapterLocator(), taskTabs = new task_tab_locator_1.TaskTabLocator()) {
            this.extractionService = extractionService;
            this.locator = locator;
            this.taskTabs = taskTabs;
        }
        listChapters() {
            return this.locator.list();
        }
        async extractSelected(selectedIndexes, onProgress = () => undefined) {
            const chapters = this.locator.list();
            const selected = [...new Set(selectedIndexes)]
                .map((index) => chapters[index])
                .filter((chapter) => Boolean(chapter));
            if (selected.length === 0)
                throw new Error('未选择可提取的章节');
            const originalIndex = this.locator.activeIndex();
            // 任务卡栏只在知识卡片页存在；批量开始前探测一次，避免逐章空等
            const hasTaskTabs = this.taskTabs.present();
            const originalTabIndex = hasTaskTabs ? this.taskTabs.activeIndex() : null;
            const chapterResults = [];
            let previousFingerprint = this.extractionService.fingerprint(this.extractionService.extract());
            try {
                for (let index = 0; index < selected.length; index += 1) {
                    const chapter = selected[index];
                    onProgress({
                        completed: index,
                        total: selected.length,
                        chapter,
                        state: 'loading',
                    });
                    const alreadyActive = this.locator.activeIndex() === chapter.index;
                    if (!alreadyActive && !this.locator.activate(chapter.index)) {
                        onProgress({
                            completed: index,
                            total: selected.length,
                            chapter,
                            state: 'failed',
                            message: '章节入口已变化或不可点击',
                        });
                        continue;
                    }
                    const result = await this.settleChapterResult({
                        alreadyActive,
                        hasTaskTabs,
                        previousFingerprint,
                    });
                    if (!result) {
                        onProgress({
                            completed: index + 1,
                            total: selected.length,
                            chapter,
                            state: 'failed',
                            message: '未在限定时间内识别到题目',
                        });
                        continue;
                    }
                    const questions = result.questions.map((question, questionIndex) => {
                        return this.attachChapter(question, chapter, questionIndex);
                    });
                    const chapterExtraction = {
                        id: chapter.id,
                        title: chapter.title,
                        questions,
                        typeOrder: (0, question_4.deriveTypeOrder)(questions),
                        statistics: (0, question_4.buildStatistics)(questions),
                        sourceUrl: result.sourceUrl,
                        extractor: result.extractor,
                    };
                    chapterResults.push(chapterExtraction);
                    previousFingerprint = this.extractionService.fingerprint(result);
                    onProgress({
                        completed: index + 1,
                        total: selected.length,
                        chapter,
                        state: 'success',
                    });
                }
            }
            finally {
                // 收尾：先还原章节，再还原任务卡（任务卡栏依赖章节页面重建后的结构）
                if (originalIndex !== null) {
                    this.locator.activate(originalIndex);
                    if (hasTaskTabs && originalTabIndex !== null) {
                        await (0, async_3.delay)(RESTORE_TAB_DELAY_MS);
                        this.taskTabs.activate(originalTabIndex);
                    }
                }
            }
            if (chapterResults.length === 0)
                throw new Error('所选章节均未提取到题目');
            const questions = chapterResults.flatMap((chapter) => chapter.questions);
            return {
                title: this.aggregateTitle(chapterResults),
                questions,
                typeOrder: (0, question_4.deriveTypeOrder)(questions),
                statistics: (0, question_4.buildStatistics)(questions),
                sourceUrl: window.location.href,
                extractor: [...new Set(chapterResults.map((chapter) => chapter.extractor))].join('+'),
                extractedAt: new Date().toISOString(),
                chapters: chapterResults,
            };
        }
        /**
         * 采样等待本章题目就绪，题目一出现立即返回。
         *
         * 章节切换会重建卡片页，新页面默认回到「视频」卡，只切一次的任务卡会随旧页面失效，
         * 因此不用「切一次卡 + 固定长等待」，而是在采样循环里每轮补切一次（限次 + 冷却）。
         * 章节刚切换时旧章节的题目 DOM 可能还在，必须等指纹变化才算本章内容，避免串页。
         */
        async settleChapterResult(options) {
            // 章节未切换且页面没有任务卡栏：结构不会自己变出题目，只取一次即可。
            // 例外是「选择题只渲染了题干、选项还没补上」，这种中间态要短暂等待补齐
            if (options.alreadyActive && !options.hasTaskTabs) {
                const immediate = this.extractionService.extract();
                if (!immediate || (0, extraction_service_1.incompleteChoiceCount)(immediate) === 0)
                    return immediate;
                return ((await this.extractionService.settleQuestions({ timeoutMs: PARTIAL_FILL_TIMEOUT_MS })) ??
                    immediate);
            }
            const guard = new task_tab_locator_1.QuestionTabGuard(this.taskTabs);
            return this.extractionService.settleQuestions({
                timeoutMs: CHAPTER_SETTLE_TIMEOUT_MS,
                previousFingerprint: options.alreadyActive ? undefined : options.previousFingerprint,
                repair: options.hasTaskTabs ? () => void guard.ensure() : undefined,
            });
        }
        attachChapter(question, chapter, questionIndex) {
            return {
                ...question,
                id: (0, hash_5.stableHash)(`${chapter.id}|${question.id}|${questionIndex}`),
                chapterId: chapter.id,
                chapterTitle: chapter.title,
            };
        }
        /**
         * 结果标题：单章节直接用章节名（文件名即章节名）；
         * 多章节用「页面标题 - N 个章节」，页面标题已内含 #prev_title 优先级。
         */
        aggregateTitle(chapters) {
            const scopeTitle = (0, page_context_3.resolvePageTitle)(document).trim();
            if (chapters.length === 1) {
                return chapters[0]?.title.trim() || scopeTitle || '学习通课程';
            }
            return [scopeTitle || '学习通课程', `${chapters.length} 个章节`].join(' - ');
        }
    }
    exports.ChapterExtractionService = ChapterExtractionService;
});
define("utils/assert", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.assertDefined = assertDefined;
    exports.queryRequired = queryRequired;
    function assertDefined(value, message) {
        if (value === null || value === undefined)
            throw new Error(message);
        return value;
    }
    function queryRequired(root, selector) {
        return assertDefined(root.querySelector(selector), `Missing required element: ${selector}`);
    }
});
define("ui/styles", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.PANEL_STYLES = void 0;
    exports.PANEL_STYLES = `
:host {
  all: initial;
  --cwe-primary: #2563eb;
  --cwe-primary-hover: #1d4ed8;
  --cwe-primary-soft: #eff6ff;
  --cwe-bg: #ffffff;
  --cwe-bg-soft: #f8fafc;
  --cwe-bg-muted: #f1f5f9;
  --cwe-text: #0f172a;
  --cwe-text-secondary: #475569;
  --cwe-text-muted: #94a3b8;
  --cwe-border: #e2e8f0;
  --cwe-success: #15803d;
  --cwe-success-bg: #f0fdf4;
  --cwe-warning: #b45309;
  --cwe-warning-bg: #fffbeb;
  --cwe-danger: #dc2626;
  --cwe-danger-bg: #fef2f2;
  --cwe-shadow: 0 20px 45px rgba(15, 23, 42, 0.18), 0 4px 12px rgba(15, 23, 42, 0.1);
  font-family: "PingFang SC", "Microsoft YaHei", system-ui, -apple-system, sans-serif;
  color: var(--cwe-text);
}

:host([data-theme="dark"]) {
  --cwe-primary: #60a5fa;
  --cwe-primary-hover: #93c5fd;
  --cwe-primary-soft: #172554;
  --cwe-bg: #0f172a;
  --cwe-bg-soft: #111827;
  --cwe-bg-muted: #1e293b;
  --cwe-text: #f8fafc;
  --cwe-text-secondary: #cbd5e1;
  --cwe-text-muted: #94a3b8;
  --cwe-border: #334155;
  --cwe-success: #4ade80;
  --cwe-success-bg: #052e16;
  --cwe-warning: #fbbf24;
  --cwe-warning-bg: #422006;
  --cwe-danger: #f87171;
  --cwe-danger-bg: #450a0a;
  --cwe-shadow: 0 24px 55px rgba(0, 0, 0, 0.5);
}

*, *::before, *::after { box-sizing: border-box; }
button, input, select { font: inherit; }
button { -webkit-tap-highlight-color: transparent; }

.cwe-hidden { display: none !important; }
.cwe-shell[hidden] { display: none !important; }

.cwe-launcher {
  position: fixed;
  z-index: 2147483000;
  top: 120px;
  right: 0;
  width: 42px;
  min-height: 116px;
  padding: 13px 9px;
  border: 0;
  border-radius: 12px 0 0 12px;
  background: linear-gradient(155deg, #3b82f6, #1d4ed8);
  color: #fff;
  box-shadow: 0 10px 28px rgba(37, 99, 235, 0.34);
  cursor: pointer;
  writing-mode: vertical-rl;
  letter-spacing: 4px;
  font-size: 13px;
  font-weight: 700;
  transition: width 160ms ease, filter 160ms ease, transform 160ms ease;
}
.cwe-launcher:hover { width: 47px; filter: brightness(1.05); }
.cwe-launcher:active { transform: translateX(1px); }

.cwe-panel {
  position: fixed;
  z-index: 2147483001;
  top: 72px;
  right: 18px;
  width: min(410px, calc(100vw - 24px));
  max-height: calc(100vh - 96px);
  display: flex;
  flex-direction: column;
  overflow: hidden;
  border: 1px solid var(--cwe-border);
  border-radius: 18px;
  background: var(--cwe-bg);
  color: var(--cwe-text);
  box-shadow: var(--cwe-shadow);
  opacity: 0;
  transform: translateX(24px) scale(0.98);
  pointer-events: none;
  transition: opacity 180ms ease, transform 180ms ease;
}
.cwe-panel[data-open="true"] {
  opacity: 1;
  transform: translateX(0) scale(1);
  pointer-events: auto;
}

.cwe-header {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 15px 16px;
  border-bottom: 1px solid var(--cwe-border);
  background: var(--cwe-bg);
  user-select: none;
}
.cwe-header-main { min-width: 0; flex: 1; }
.cwe-title { margin: 0; font-size: 15px; font-weight: 800; line-height: 1.3; }
/* 版本号常驻标题右侧：用户报问题时不必再问「装的是哪一版」 */
.cwe-version { margin-left: 6px; color: var(--cwe-text-muted); font-size: 10.5px; font-weight: 600; }
.cwe-header-actions { display: flex; gap: 5px; }
.cwe-icon-button {
  width: 31px;
  height: 31px;
  display: inline-grid;
  place-items: center;
  border: 1px solid transparent;
  border-radius: 9px;
  background: transparent;
  color: var(--cwe-text-secondary);
  cursor: pointer;
}
.cwe-icon-button:hover { border-color: var(--cwe-border); background: var(--cwe-bg-muted); color: var(--cwe-text); }
.cwe-icon-button svg { width: 16px; height: 16px; fill: currentColor; }

.cwe-scroll { overflow: auto; padding: 15px 16px 18px; scrollbar-width: thin; }
.cwe-scroll::-webkit-scrollbar { width: 6px; }
.cwe-scroll::-webkit-scrollbar-thumb { background: var(--cwe-border); border-radius: 99px; }

.cwe-status-row { display: flex; align-items: center; gap: 8px; }
.cwe-status {
  flex: 1;
  min-width: 0;
  min-height: 38px;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 9px 11px;
  border: 1px solid var(--cwe-border);
  border-radius: 10px;
  background: var(--cwe-bg-soft);
  color: var(--cwe-text-secondary);
  font-size: 12px;
  line-height: 1.5;
  text-align: center;
}
.cwe-status[data-kind="success"] { border-color: color-mix(in srgb, var(--cwe-success) 35%, transparent); background: var(--cwe-success-bg); color: var(--cwe-success); }
.cwe-status[data-kind="warning"] { border-color: color-mix(in srgb, var(--cwe-warning) 35%, transparent); background: var(--cwe-warning-bg); color: var(--cwe-warning); }
.cwe-status[data-kind="error"] { border-color: color-mix(in srgb, var(--cwe-danger) 35%, transparent); background: var(--cwe-danger-bg); color: var(--cwe-danger); }

.cwe-extract-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 9px; margin-top: 11px; }
.cwe-button {
  min-height: 40px;
  padding: 9px 13px;
  border: 1px solid var(--cwe-border);
  border-radius: 10px;
  background: var(--cwe-bg);
  color: var(--cwe-text-secondary);
  cursor: pointer;
  font-size: 12.5px;
  font-weight: 700;
  transition: transform 120ms ease, border-color 120ms ease, background 120ms ease;
}
.cwe-button:hover:not(:disabled) { border-color: var(--cwe-primary); color: var(--cwe-primary); }
.cwe-button:active:not(:disabled) { transform: scale(0.985); }
.cwe-button:disabled { opacity: 0.5; cursor: not-allowed; }
.cwe-button-primary { border-color: var(--cwe-primary); background: var(--cwe-primary); color: #fff; }
.cwe-button-primary:hover:not(:disabled) { background: var(--cwe-primary-hover); color: #fff; }
.cwe-button-soft { border-color: color-mix(in srgb, var(--cwe-primary) 35%, transparent); background: var(--cwe-primary-soft); color: var(--cwe-primary); }

.cwe-stats { display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; margin: 12px 0; }
.cwe-stat { padding: 8px 3px 7px; border: 1px solid var(--cwe-border); border-radius: 10px; background: var(--cwe-bg-soft); text-align: center; }
.cwe-stat-value { display: block; color: var(--cwe-primary); font-size: 18px; font-weight: 800; line-height: 1.15; }
.cwe-stat-label { display: block; margin-top: 3px; color: var(--cwe-text-muted); font-size: 9.5px; }

.cwe-section { margin-top: 11px; padding: 12px; border: 1px solid var(--cwe-border); border-radius: 12px; background: var(--cwe-bg-soft); }
.cwe-section-title { margin: 0 0 9px; color: var(--cwe-text-secondary); font-size: 11px; font-weight: 800; letter-spacing: 0.04em; }
.cwe-field { display: grid; gap: 6px; }
.cwe-label { color: var(--cwe-text-secondary); font-size: 11.5px; font-weight: 700; }
.cwe-input, .cwe-select {
  width: 100%;
  min-height: 36px;
  padding: 7px 10px;
  border: 1px solid var(--cwe-border);
  border-radius: 9px;
  outline: none;
  background: var(--cwe-bg);
  color: var(--cwe-text);
  font-size: 12px;
}
.cwe-input:focus, .cwe-select:focus { border-color: var(--cwe-primary); box-shadow: 0 0 0 3px color-mix(in srgb, var(--cwe-primary) 14%, transparent); }

.cwe-formats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 7px; }
.cwe-format { position: relative; }
.cwe-format input { position: absolute; opacity: 0; pointer-events: none; }
.cwe-format span { display: block; padding: 8px 5px; border: 1px solid var(--cwe-border); border-radius: 9px; background: var(--cwe-bg); color: var(--cwe-text-secondary); font-size: 11.5px; font-weight: 700; text-align: center; cursor: pointer; }
.cwe-format input:checked + span { border-color: var(--cwe-primary); background: var(--cwe-primary-soft); color: var(--cwe-primary); }
.cwe-format input:focus-visible + span { outline: 2px solid var(--cwe-primary); outline-offset: 2px; }

.cwe-options { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 9px; margin-top: 10px; }
.cwe-check { min-height: 34px; display: flex; align-items: center; gap: 8px; padding: 7px 8px; border-radius: 8px; color: var(--cwe-text-secondary); cursor: pointer; font-size: 12.5px; }
.cwe-check:hover { background: var(--cwe-bg-muted); }
.cwe-check input { width: 15px; height: 15px; margin: 0; accent-color: var(--cwe-primary); }
.cwe-check:has(input:disabled) { opacity: 0.45; cursor: not-allowed; }

.cwe-actions { display: grid; grid-template-columns: 1.25fr 1fr; gap: 9px; margin-top: 12px; }
.cwe-progress { width: 14px; height: 14px; display: inline-block; margin-right: 6px; border: 2px solid currentColor; border-right-color: transparent; border-radius: 50%; vertical-align: -2px; animation: cwe-spin 700ms linear infinite; }
@keyframes cwe-spin { to { transform: rotate(360deg); } }

.cwe-footer { margin-top: 10px; color: var(--cwe-text-muted); font-size: 10px; text-align: center; }

.cwe-modal-backdrop {
  position: fixed;
  z-index: 2147483005;
  inset: 0;
  display: grid;
  place-items: center;
  padding: 18px;
  background: rgba(15, 23, 42, 0.5);
  opacity: 0;
  pointer-events: none;
  transition: opacity 150ms ease;
}
.cwe-modal-backdrop[data-open="true"] { opacity: 1; pointer-events: auto; }
.cwe-modal {
  width: min(390px, calc(100vw - 28px));
  max-height: min(680px, calc(100vh - 36px));
  display: flex;
  flex-direction: column;
  overflow: hidden;
  border: 1px solid var(--cwe-border);
  border-radius: 16px;
  background: var(--cwe-bg);
  box-shadow: var(--cwe-shadow);
  transform: translateY(8px) scale(0.98);
  transition: transform 150ms ease;
}
.cwe-modal-backdrop[data-open="true"] .cwe-modal { transform: translateY(0) scale(1); }
/* 章节选择弹窗的条目与标题都更长，单独放宽 */
.cwe-modal-wide { width: min(560px, calc(100vw - 28px)); }
.cwe-modal-header { display: flex; align-items: center; gap: 8px; padding: 14px 15px; border-bottom: 1px solid var(--cwe-border); }
.cwe-modal-title { flex: 1; margin: 0; font-size: 14px; font-weight: 800; }
.cwe-modal-body { overflow: auto; padding: 14px 15px; }
.cwe-modal-footer { display: flex; justify-content: flex-end; gap: 8px; padding: 12px 15px; border-top: 1px solid var(--cwe-border); }

.cwe-setting-row { display: grid; grid-template-columns: 118px 1fr; align-items: center; gap: 10px; margin-bottom: 12px; }
.cwe-setting-row:last-child { margin-bottom: 0; }
/* 快捷键设置已移到开关组下方，用分隔线与上方区分 */
.cwe-setting-row-split { margin-top: 2px; padding-top: 13px; border-top: 1px solid var(--cwe-border); }
.cwe-setting-label { color: var(--cwe-text-secondary); font-size: 11.5px; font-weight: 700; }
/* 开关组行距放宽：原 9px 太挤，与后面的设置行连成一片；标签字重与上方设置行一致 */
.cwe-switch-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 13px 0; border-top: 1px solid var(--cwe-border); color: var(--cwe-text-secondary); font-size: 11.5px; font-weight: 700; }

.cwe-history-modal {
  display: flex;
  flex-direction: column;
  width: 460px;
  max-width: calc(100vw - 36px);
  max-height: 520px;
  padding: 24px;
  border-radius: 14px;
  background: #fff;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.15);
  transform: translateY(12px);
  transition: transform 0.25s ease;
}
.cwe-modal-backdrop[data-open="true"] .cwe-history-modal { transform: translateY(0); }
:host([data-theme="dark"]) .cwe-history-modal { background: #1e1e2e; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4); }
.cwe-history-modal-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 14px;
  padding-bottom: 14px;
  border-bottom: 1px solid #eee;
  flex-shrink: 0;
}
:host([data-theme="dark"]) .cwe-history-modal-header { border-bottom-color: #313244; }
.cwe-history-modal-title { font-size: 15px; font-weight: 700; color: #222; margin: 0; }
:host([data-theme="dark"]) .cwe-history-modal-title { color: #cdd6f4; }
.cwe-history-modal-close {
  width: 24px;
  height: 24px;
  border: none;
  background: #f0f0f0;
  border-radius: 50%;
  font-size: 15px;
  color: #999;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s ease;
}
.cwe-history-modal-close:hover { background: #e0e0e0; color: #555; }
:host([data-theme="dark"]) .cwe-history-modal-close { background: #313244; color: #a6adc8; }
:host([data-theme="dark"]) .cwe-history-modal-close:hover { background: #45475a; color: #cdd6f4; }
.cwe-history-list { flex: 1; overflow-y: auto; }
.cwe-history-empty { text-align: center; color: #bbb; padding: 40px 0; font-size: 13px; }
:host([data-theme="dark"]) .cwe-history-empty { color: #585b70; }
.cwe-history-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 14px;
  border-radius: 10px;
  background: #fafbfc;
  border: 1px solid #eef0f2;
  margin-bottom: 8px;
  transition: background 0.2s;
  cursor: pointer;
}
.cwe-history-item:hover { background: #f0f2f5; }
:host([data-theme="dark"]) .cwe-history-item { background: #181825; border-color: #313244; }
:host([data-theme="dark"]) .cwe-history-item:hover { background: #1e1e2e; }
.cwe-history-info { flex: 1; min-width: 0; }
.cwe-history-title {
  font-size: 13px;
  font-weight: 600;
  color: #333;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
:host([data-theme="dark"]) .cwe-history-title { color: #cdd6f4; }
.cwe-history-meta { font-size: 11px; color: #999; margin-top: 3px; }
:host([data-theme="dark"]) .cwe-history-meta { color: #a6adc8; }
.cwe-history-delete {
  width: 26px;
  height: 26px;
  border: none;
  background: transparent;
  border-radius: 50%;
  cursor: pointer;
  color: #ccc;
  font-size: 16px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s ease;
  flex-shrink: 0;
  margin-left: 8px;
}
.cwe-history-delete:hover { background: #fee2e2; color: #ef4444; }
:host([data-theme="dark"]) .cwe-history-delete { color: #585b70; }
:host([data-theme="dark"]) .cwe-history-delete:hover { background: #450a0a; color: #f87171; }
.cwe-history-download {
  width: 26px;
  height: 26px;
  border: none;
  background: transparent;
  border-radius: 50%;
  cursor: pointer;
  color: #ccc;
  font-size: 14px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s ease;
  flex-shrink: 0;
  margin-left: 4px;
}
.cwe-history-download:hover { background: #e3f2fd; color: #1e88e5; }
:host([data-theme="dark"]) .cwe-history-download { color: #585b70; }
:host([data-theme="dark"]) .cwe-history-download:hover { background: #0d2137; color: #60a5fa; }
.cwe-mini-button { padding: 5px 8px; border: 1px solid var(--cwe-border); border-radius: 7px; background: var(--cwe-bg); color: var(--cwe-text-secondary); cursor: pointer; font-size: 10.5px; }
.cwe-mini-button:hover { border-color: var(--cwe-primary); color: var(--cwe-primary); }
.cwe-mini-button-danger:hover { border-color: var(--cwe-danger); color: var(--cwe-danger); }

.cwe-chapter-toolbar { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 10px; }
/* 全选框取代原来的两个按钮，右侧跟选择状态文本 */
.cwe-chapter-all { display: inline-flex; align-items: center; gap: 7px; color: var(--cwe-text-secondary); cursor: pointer; font-size: 11.5px; font-weight: 700; }
.cwe-chapter-all input { width: 15px; height: 15px; margin: 0; accent-color: var(--cwe-primary); }
.cwe-chapter-all:has(input:disabled) { opacity: 0.5; cursor: not-allowed; }
.cwe-chapter-status { color: var(--cwe-text-muted); font-size: 11px; }
.cwe-chapter-list { display: grid; gap: 6px; max-height: 390px; overflow: auto; }
.cwe-chapter { display: flex; align-items: flex-start; gap: 8px; padding: 9px; border: 1px solid var(--cwe-border); border-radius: 9px; background: var(--cwe-bg-soft); color: var(--cwe-text-secondary); cursor: pointer; font-size: 11.5px; line-height: 1.4; }
.cwe-chapter:hover { border-color: var(--cwe-primary); }
.cwe-chapter input { margin-top: 1px; accent-color: var(--cwe-primary); }
.cwe-chapter-progress { min-height: 18px; margin-top: 9px; color: var(--cwe-text-muted); font-size: 10.5px; }

@media (max-width: 560px) {
  .cwe-panel { top: 12px; right: 12px; max-height: calc(100vh - 24px); }
  .cwe-stats { grid-template-columns: repeat(3, 1fr); }
  .cwe-options { grid-template-columns: 1fr; }
}
`;
});
define("ui/panel-view", ["require", "exports", "app-config", "utils/assert", "utils/shortcut", "utils/text", "ui/styles"], function (require, exports, app_config_1, assert_1, shortcut_2, text_8, styles_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.PanelView = void 0;
    const NOOP_CALLBACKS = {
        onExtract: () => undefined,
        onOpenChapters: () => undefined,
        onExtractChapters: () => undefined,
        onDownload: () => undefined,
        onCopy: () => undefined,
        onDiagnose: () => undefined,
        onPreferencesChange: () => undefined,
        onSettingsSave: () => undefined,
        onHistoryRestore: () => undefined,
        onHistoryDownload: () => undefined,
        onHistoryDelete: () => undefined,
        onPanelPositionChange: () => undefined,
    };
    const ICONS = {
        settings: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.62l-1.92-3.32a.49.49 0 0 0-.59-.22l-2.39.96a7.2 7.2 0 0 0-1.62-.94L14.4 2.8a.48.48 0 0 0-.48-.4h-3.84a.48.48 0 0 0-.48.4l-.36 2.54c-.59.24-1.13.56-1.62.94l-2.39-.96a.48.48 0 0 0-.59.22L2.72 8.86a.48.48 0 0 0 .12.62l2.03 1.58c-.05.3-.08.63-.08.94s.03.64.08.94l-2.03 1.58a.49.49 0 0 0-.12.62l1.92 3.32c.12.22.37.3.59.22l2.39-.96c.49.38 1.03.7 1.62.94l.36 2.54c.04.23.24.4.48.4h3.84c.24 0 .44-.17.48-.4l.36-2.54a7.2 7.2 0 0 0 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32a.49.49 0 0 0-.12-.62l-2.02-1.58ZM12 15.5A3.5 3.5 0 1 1 12 8a3.5 3.5 0 0 1 0 7.5Z"/></svg>',
        history: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M13 3a9 9 0 1 0 8.48 6H19.3A7 7 0 1 1 13 5c1.93 0 3.68.78 4.95 2.05L15 10h7V3l-2.63 2.63A8.96 8.96 0 0 0 13 3Zm-1 4v6l5 3 .75-1.23-4.25-2.52V7H12Z"/></svg>',
        close: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m18.3 5.7-1-1L12 10l-5.3-5.3-1 1L11 11l-5.3 5.3 1 1L12 12l5.3 5.3 1-1L13 11l5.3-5.3Z"/></svg>',
    };
    // 概览统计区展示的题型，与卡片顺序一致
    const TYPE_STAT_KEYS = [
        'single-choice',
        'multiple-choice',
        'fill-blank',
        'true-false',
        'short-answer',
    ];
    /** 选择题却没有选项的题目数（只做纯数据判断，不依赖应用层） */
    function countChoiceQuestionsWithoutOptions(result) {
        return result.questions.filter((question) => {
            const isChoice = question.type === 'single-choice' || question.type === 'multiple-choice';
            return isChoice && question.options.length === 0;
        }).length;
    }
    function template() {
        return `
    <style>${styles_1.PANEL_STYLES}</style>
    <div class="cwe-shell">
      <button class="cwe-launcher" type="button" aria-label="打开题目导出面板">提取题目</button>
      <section class="cwe-panel" data-open="false" role="dialog" aria-label="${app_config_1.APP_NAME}">
        <header class="cwe-header">
          <div class="cwe-header-main">
            <h2 class="cwe-title">学习通题目导出<span class="cwe-version">v${app_config_1.APP_VERSION}</span></h2>
          </div>
          <div class="cwe-header-actions">
            <button class="cwe-icon-button" data-action="history" type="button" title="下载历史">${ICONS.history}</button>
            <button class="cwe-icon-button" data-action="settings" type="button" title="设置">${ICONS.settings}</button>
            <button class="cwe-icon-button" data-action="close" type="button" title="关闭">${ICONS.close}</button>
          </div>
        </header>
        <div class="cwe-scroll">
          <div class="cwe-status-row">
            <div class="cwe-status" data-kind="neutral">进入作业、考试或章节练习页面后开始提取</div>
            <button class="cwe-mini-button cwe-diagnose" data-action="diagnose" type="button" hidden>复制诊断信息</button>
          </div>
          <div class="cwe-extract-grid">
            <button class="cwe-button cwe-button-primary" data-action="extract" type="button">提取当前页面</button>
            <button class="cwe-button cwe-button-soft" data-action="chapters" type="button" disabled>提取多个章节</button>
          </div>

          <div class="cwe-stats" aria-label="提取统计">
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="single-choice">0</span><span class="cwe-stat-label">单选</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="multiple-choice">0</span><span class="cwe-stat-label">多选</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="fill-blank">0</span><span class="cwe-stat-label">填空</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="true-false">0</span><span class="cwe-stat-label">判断</span></div>
            <div class="cwe-stat"><span class="cwe-stat-value" data-stat="short-answer">0</span><span class="cwe-stat-label">简答</span></div>
          </div>

          <section class="cwe-section">
            <h3 class="cwe-section-title">导出文件</h3>
            <label class="cwe-field">
              <span class="cwe-label">文件名</span>
              <input class="cwe-input" data-field="filename" type="text" maxlength="120" value="学习通题目" />
            </label>
            <div class="cwe-formats" style="margin-top: 9px">
              <label class="cwe-format"><input type="radio" name="cwe-format" value="word" checked><span>Word 试卷</span></label>
              <label class="cwe-format"><input type="radio" name="cwe-format" value="txt"><span>TXT 文本</span></label>
              <label class="cwe-format"><input type="radio" name="cwe-format" value="md"><span>Markdown</span></label>
            </div>
            <div class="cwe-options">
              <label class="cwe-check"><input data-option="withAnswers" type="checkbox">附加答案</label>
              <label class="cwe-check"><input data-option="withWrong" type="checkbox">附加错题</label>
              <label class="cwe-check"><input data-option="shuffle" type="checkbox">题型内乱序</label>
              <label class="cwe-check"><input data-option="bankImport" type="checkbox">题库导入</label>
              <label class="cwe-check"><input data-option="splitByChapter" type="checkbox" disabled>按章节拆分文件</label>
            </div>
          </section>

          <div class="cwe-actions">
            <button class="cwe-button cwe-button-primary" data-action="download" type="button" disabled>下载文件</button>
            <button class="cwe-button" data-action="copy" type="button" disabled>复制文本</button>
          </div>
          <div class="cwe-footer">仅处理当前浏览器中已加载的题目；请遵守课程与平台规则。</div>
        </div>
      </section>

      <div class="cwe-modal-backdrop" data-modal="settings" data-open="false">
        <section class="cwe-modal" role="dialog" aria-modal="true" aria-label="脚本设置">
          <header class="cwe-modal-header">
            <h3 class="cwe-modal-title">脚本设置</h3>
            <button class="cwe-icon-button" data-modal-close="settings" type="button">${ICONS.close}</button>
          </header>
          <div class="cwe-modal-body">
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">界面主题</span>
              <select class="cwe-select" data-setting="theme">
                <option value="auto">跟随系统</option><option value="light">浅色</option><option value="dark">深色</option>
              </select>
            </label>
            <label class="cwe-switch-row"><span>允许拖动面板</span><input data-setting="enableDrag" type="checkbox"></label>
            <label class="cwe-switch-row"><span>记忆面板位置</span><input data-setting="rememberPanelPosition" type="checkbox"></label>
            <label class="cwe-switch-row"><span>打开窗口自动提取</span><input data-setting="autoExtractOnLoad" type="checkbox"></label>
            <label class="cwe-setting-row cwe-setting-row-split">
              <span class="cwe-setting-label">提取快捷键</span>
              <input class="cwe-input" data-setting="shortcut" placeholder="Ctrl+Shift+E" />
            </label>
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">隐藏快捷键</span>
              <input class="cwe-input" data-setting="hideShortcut" placeholder="Ctrl+Shift+H" />
            </label>
          </div>
          <footer class="cwe-modal-footer">
            <button class="cwe-button" data-modal-close="settings" type="button">取消</button>
            <button class="cwe-button cwe-button-primary" data-action="save-settings" type="button">保存设置</button>
          </footer>
        </section>
      </div>

      <div class="cwe-modal-backdrop" data-modal="history" data-open="false">
        <section class="cwe-history-modal" role="dialog" aria-modal="true" aria-label="下载历史">
          <header class="cwe-history-modal-header">
            <h3 class="cwe-history-modal-title">历史记录</h3>
            <button class="cwe-history-modal-close" data-modal-close="history" type="button" aria-label="关闭">&times;</button>
          </header>
          <div class="cwe-history-list" data-history-list></div>
        </section>
      </div>

      <div class="cwe-modal-backdrop" data-modal="chapters" data-open="false">
        <section class="cwe-modal cwe-modal-wide" role="dialog" aria-modal="true" aria-label="选择章节">
          <header class="cwe-modal-header">
            <h3 class="cwe-modal-title">选择要提取的章节</h3>
            <button class="cwe-icon-button" data-modal-close="chapters" type="button">${ICONS.close}</button>
          </header>
          <div class="cwe-modal-body">
            <div class="cwe-chapter-toolbar">
              <label class="cwe-chapter-all">
                <input type="checkbox" data-action="chapter-all-toggle" checked><span>全选</span>
              </label>
              <span class="cwe-chapter-status" data-chapter-status></span>
            </div>
            <div class="cwe-chapter-list" data-chapter-list></div>
            <div class="cwe-chapter-progress" data-chapter-progress></div>
          </div>
          <footer class="cwe-modal-footer">
            <button class="cwe-button" data-modal-close="chapters" type="button">取消</button>
            <button class="cwe-button cwe-button-primary" data-action="start-chapters" type="button">开始提取</button>
          </footer>
        </section>
      </div>
    </div>
  `;
    }
    class PanelView {
        constructor() {
            this.callbacks = NOOP_CALLBACKS;
            this.settings = null;
            this.chapterCount = 0;
            this.busy = null;
            this.globallyHidden = false;
            // 提取结果是否含参考答案（无答案时隐藏“附加参考答案/附加错题汇总”选项）
            this.answerOptionsAvailable = true;
            this.mediaQuery = null;
            this.onColorSchemeChange = () => this.applyResolvedTheme();
            const existing = document.getElementById(app_config_1.UI_HOST_ID);
            existing?.remove();
            this.host = document.createElement('div');
            this.host.id = app_config_1.UI_HOST_ID;
            this.shadow = this.host.attachShadow({ mode: 'open' });
            this.shadow.innerHTML = template();
            document.body.appendChild(this.host);
            this.bindEvents();
            this.enableDragging();
        }
        setCallbacks(callbacks) {
            this.callbacks = callbacks;
        }
        open() {
            this.panel().dataset.open = 'true';
            this.launcher().classList.add('cwe-hidden');
        }
        close() {
            this.closeAllModals();
            this.panel().dataset.open = 'false';
            if (!this.globallyHidden)
                this.launcher().classList.remove('cwe-hidden');
        }
        toggle() {
            if (this.panel().dataset.open === 'true')
                this.close();
            else
                this.open();
        }
        toggleGlobalVisibility() {
            this.globallyHidden = !this.globallyHidden;
            this.shell().hidden = this.globallyHidden;
            return this.globallyHidden;
        }
        setStatus(message, kind = 'neutral') {
            const status = this.element('.cwe-status');
            status.textContent = message;
            status.dataset.kind = kind;
        }
        /**
         * 显示/隐藏「复制诊断信息」按钮。
         * 只在提取结果里存在「选择题却没有选项」时出现 —— 此时真机现场的结构是定位所必需的，
         * 而离线快照复现不出来；正常结果不显示，避免干扰。
         */
        setDiagnoseVisible(visible) {
            this.element('[data-action="diagnose"]').hidden = !visible;
        }
        setBusy(action, message) {
            this.busy = action;
            const controls = this.shadow.querySelectorAll('[data-action="extract"], [data-action="chapters"], [data-action="download"], [data-action="copy"], [data-action="diagnose"]');
            controls.forEach((button) => {
                button.disabled = action !== null || this.shouldDisableButton(button.dataset.action ?? '');
            });
            const chapterBusy = action === 'chapters';
            this.element('[data-action="start-chapters"]').disabled = chapterBusy;
            this.element('[data-action="chapter-all-toggle"]').disabled = chapterBusy;
            this.shadow.querySelectorAll('[data-chapter-list] input').forEach((input) => {
                input.disabled = chapterBusy;
            });
            const extract = this.element('[data-action="extract"]');
            const chapters = this.element('[data-action="chapters"]');
            const download = this.element('[data-action="download"]');
            const copy = this.element('[data-action="copy"]');
            extract.innerHTML =
                action === 'extract' ? '<span class="cwe-progress"></span>提取中' : '提取当前页面';
            chapters.innerHTML =
                action === 'chapters' ? '<span class="cwe-progress"></span>遍历中' : '提取多个章节';
            download.innerHTML =
                action === 'download' ? '<span class="cwe-progress"></span>生成中' : '下载文件';
            copy.innerHTML = action === 'copy' ? '<span class="cwe-progress"></span>复制中' : '复制文本';
            if (message)
                this.setStatus(message, 'neutral');
        }
        renderResult(result) {
            for (const name of TYPE_STAT_KEYS) {
                const count = result.statistics.byType[name];
                this.stat(name).textContent = String(count);
                this.setStatVisible(name, count > 0);
            }
            this.element('[data-field="filename"]').value = (0, text_8.sanitizeFilename)(result.title);
            // 提取结果无参考答案时，隐藏“附加参考答案/附加错题汇总”选项
            this.setAnswerOptionsAvailable(result.statistics.withCorrectAnswer > 0);
            this.refreshActionAvailability(true);
            this.updateFilenamePreview();
            const chapterMessage = result.chapters?.length ? `，来自 ${result.chapters.length} 个章节` : '';
            // 选择题却没有选项，是「页面还没渲染完就采到了」的典型特征：直接说出来，
            // 否则用户只会看到「已提取 N 道题」而不知道选项这一层根本没取到
            const withoutOptions = countChoiceQuestionsWithoutOptions(result);
            this.setDiagnoseVisible(withoutOptions > 0);
            if (withoutOptions > 0) {
                this.setStatus(`已提取 ${result.statistics.total} 道题${chapterMessage}，其中 ${withoutOptions} 道选择题没提取到选项` +
                    `（点「复制诊断信息」把现场结构复制出来发给维护者，比重试更快定位）`, 'warning');
                return;
            }
            this.setStatus(`已提取 ${result.statistics.total} 道题${chapterMessage}`, 'success');
        }
        clearResult() {
            for (const name of TYPE_STAT_KEYS) {
                this.stat(name).textContent = '0';
                this.setStatVisible(name, false);
            }
            // 重置为默认显示，等待下次提取结果决定
            this.setAnswerOptionsAvailable(true);
            this.setDiagnoseVisible(false);
            this.refreshActionAvailability(false);
        }
        setChapterCount(count) {
            this.chapterCount = count;
            const button = this.element('[data-action="chapters"]');
            button.disabled = this.busy !== null || count < 2;
            button.title = count < 2 ? '当前页面未识别到多个章节入口' : `已识别 ${count} 个章节`;
            const split = this.option('splitByChapter');
            split.disabled = count < 2;
            if (count < 2)
                split.checked = false;
        }
        readExportOptions() {
            const formatElement = this.shadow.querySelector('input[name="cwe-format"]:checked');
            const format = formatElement?.value === 'txt' || formatElement?.value === 'md'
                ? formatElement.value
                : 'word';
            return {
                format,
                filename: this.element('[data-field="filename"]').value,
                withAnswers: this.option('withAnswers').checked,
                withWrong: this.option('withWrong').checked,
                shuffle: this.option('shuffle').checked,
                bankImport: format === 'word' && this.option('bankImport').checked,
                splitByChapter: this.option('splitByChapter').checked,
            };
        }
        readExportPreferences() {
            const options = this.readExportOptions();
            return {
                format: options.format,
                withAnswers: options.withAnswers,
                withWrong: options.withWrong,
                shuffle: options.shuffle,
                bankImport: options.bankImport,
                splitByChapter: options.splitByChapter,
            };
        }
        applyExportOptions(options) {
            this.element('[data-field="filename"]').value = options.filename;
            const format = this.shadow.querySelector(`input[name="cwe-format"][value="${options.format}"]`);
            if (format)
                format.checked = true;
            this.option('withAnswers').checked = options.withAnswers;
            this.option('withWrong').checked = options.withWrong;
            this.option('shuffle').checked = options.shuffle;
            this.option('bankImport').checked = options.bankImport;
            this.option('splitByChapter').checked = options.splitByChapter && this.chapterCount > 1;
            this.updateOptionAvailability();
        }
        applySettings(settings) {
            this.settings = settings;
            this.element('[data-setting="theme"]').value = settings.theme;
            this.element('[data-setting="shortcut"]').value = (0, shortcut_2.formatShortcut)(settings.shortcut);
            this.element('[data-setting="hideShortcut"]').value = (0, shortcut_2.formatShortcut)(settings.hideShortcut);
            this.element('[data-setting="enableDrag"]').checked = settings.enableDrag;
            this.element('[data-setting="rememberPanelPosition"]').checked =
                settings.rememberPanelPosition;
            this.element('[data-setting="autoExtractOnLoad"]').checked =
                settings.autoExtractOnLoad;
            const preferences = settings.exportPreferences;
            this.applyExportOptions({
                ...preferences,
                filename: this.element('[data-field="filename"]').value,
            });
            this.applyResolvedTheme();
            this.applyPosition(settings.panelPosition);
        }
        renderHistory(entries) {
            const list = this.element('[data-history-list]');
            list.replaceChildren();
            if (entries.length === 0) {
                const empty = document.createElement('div');
                empty.className = 'cwe-history-empty';
                empty.textContent = '暂无历史记录';
                list.appendChild(empty);
                return;
            }
            const fmtNames = { txt: 'TXT', md: 'MD', word: 'Word' };
            for (const entry of entries) {
                const item = document.createElement('article');
                item.className = 'cwe-history-item';
                item.dataset.historyId = entry.id;
                const info = document.createElement('div');
                info.className = 'cwe-history-info';
                const title = document.createElement('div');
                title.className = 'cwe-history-title';
                title.textContent = entry.title;
                title.title = entry.title;
                // 与主脚本一致：按导出配置显示“含答案、含错题、打乱、题库导入”标志
                const flags = [];
                if (entry.options.withAnswers)
                    flags.push('含答案');
                if (entry.options.withWrong)
                    flags.push('含错题');
                if (entry.options.shuffle)
                    flags.push('打乱');
                if (entry.options.bankImport)
                    flags.push('题库导入');
                const flagStr = flags.length > 0 ? ` · ${flags.join('、')}` : '';
                const fmtName = fmtNames[entry.options.format] || entry.options.format.toUpperCase();
                const meta = document.createElement('div');
                meta.className = 'cwe-history-meta';
                meta.textContent = `${new Date(entry.createdAt).toLocaleString()} · ${entry.result.statistics.total}题 · ${fmtName}${flagStr}`;
                const downloadBtn = document.createElement('button');
                downloadBtn.type = 'button';
                downloadBtn.className = 'cwe-history-download';
                downloadBtn.title = '重新下载';
                downloadBtn.textContent = '⤓';
                downloadBtn.dataset.historyAction = 'download';
                const deleteBtn = document.createElement('button');
                deleteBtn.type = 'button';
                deleteBtn.className = 'cwe-history-delete';
                deleteBtn.title = '删除';
                deleteBtn.textContent = '✕';
                deleteBtn.dataset.historyAction = 'delete';
                info.append(title, meta);
                item.append(info, downloadBtn, deleteBtn);
                list.appendChild(item);
            }
        }
        openHistory() {
            this.openModal('history');
        }
        closeHistory() {
            this.closeModal('history');
        }
        showChapterDialog(chapters) {
            const list = this.element('[data-chapter-list]');
            list.replaceChildren();
            chapters.forEach((chapter) => {
                const label = document.createElement('label');
                label.className = 'cwe-chapter';
                const input = document.createElement('input');
                input.type = 'checkbox';
                input.checked = true;
                input.value = String(chapter.index);
                const text = document.createElement('span');
                text.textContent = chapter.title;
                label.append(input, text);
                list.appendChild(label);
            });
            // 章节数改由工具栏右侧的状态文本显示，底部这行留给提取进度
            this.element('[data-chapter-progress]').textContent = '';
            this.refreshChapterSelection();
            this.openModal('chapters');
        }
        /**
         * 刷新章节弹窗的「全选」框与右侧状态文本。
         *
         * 全选框兼作选择状态的指示：全选时勾上、一个都没选时取消、只选了一部分用
         * indeterminate（半选）表示，避免用户以为「已全选」。
         */
        refreshChapterSelection() {
            const inputs = [...this.shadow.querySelectorAll('[data-chapter-list] input')];
            const total = inputs.length;
            const selected = inputs.filter((input) => input.checked).length;
            const toggle = this.element('[data-action="chapter-all-toggle"]');
            toggle.checked = total > 0 && selected === total;
            toggle.indeterminate = selected > 0 && selected < total;
            this.element('[data-chapter-status]').textContent =
                total === 0 ? '未识别到章节' : `已选 ${selected} / 共 ${total} 个章节`;
        }
        /**
         * 章节遍历进度。
         *
         * `completed` 在 loading 态是「已完成的章节数」、在 success/failed 态是「含本章的完成数」，
         * 因此直接当「已提取 N/共 M 章节」的计数用。弹窗里的那行给完整信息（状态 + 章节名 + 失败原因），
         * 面板状态栏同步同一个计数，避免关掉弹窗后看不到进度。
         */
        updateChapterProgress(progress) {
            const stateLabel = progress.state === 'loading' ? '正在加载' : progress.state === 'success' ? '已完成' : '失败';
            const counter = `已提取 ${progress.completed}/${progress.total} 章节`;
            const detail = `${progress.chapter.title}${progress.message ? `（${progress.message}）` : ''}`;
            this.element('[data-chapter-progress]').textContent =
                `${counter} · ${stateLabel}：${detail}`;
            this.setStatus(`${counter} · 当前：${progress.chapter.title}`, progress.state === 'failed' ? 'warning' : 'neutral');
        }
        closeChapterDialog() {
            this.closeModal('chapters');
        }
        destroy() {
            this.mediaQuery?.removeEventListener('change', this.onColorSchemeChange);
            this.host.remove();
        }
        bindEvents() {
            this.launcher().addEventListener('click', () => this.open());
            this.element('[data-action="close"]').addEventListener('click', () => this.close());
            this.element('[data-action="extract"]').addEventListener('click', () => this.callbacks.onExtract());
            this.element('[data-action="chapters"]').addEventListener('click', () => this.callbacks.onOpenChapters());
            this.element('[data-action="download"]').addEventListener('click', () => this.callbacks.onDownload());
            this.element('[data-action="copy"]').addEventListener('click', () => this.callbacks.onCopy());
            this.element('[data-action="diagnose"]').addEventListener('click', () => this.callbacks.onDiagnose());
            this.element('[data-action="settings"]').addEventListener('click', () => this.openModal('settings'));
            this.element('[data-action="history"]').addEventListener('click', () => this.openHistory());
            this.shadow
                .querySelectorAll('input[name="cwe-format"], [data-option]')
                .forEach((input) => {
                input.addEventListener('change', () => {
                    this.updateOptionAvailability(input);
                    this.callbacks.onPreferencesChange(this.readExportPreferences());
                });
            });
            this.element('[data-action="save-settings"]').addEventListener('click', () => this.saveSettings());
            // 全选框取代原来的「全选 / 取消全选」两个按钮
            this.element('[data-action="chapter-all-toggle"]').addEventListener('change', (event) => {
                this.setAllChapters(event.target.checked);
            });
            // 单章勾选变化时同步工具栏的全选状态与「已选 N / 共 N」文本
            this.element('[data-chapter-list]').addEventListener('change', () => this.refreshChapterSelection());
            this.element('[data-action="start-chapters"]').addEventListener('click', () => {
                const indexes = [
                    ...this.shadow.querySelectorAll('[data-chapter-list] input:checked'),
                ]
                    .map((input) => Number.parseInt(input.value, 10))
                    .filter(Number.isFinite);
                this.callbacks.onExtractChapters(indexes);
            });
            this.shadow.querySelectorAll('[data-modal-close]').forEach((button) => {
                button.addEventListener('click', () => this.closeModal(button.dataset.modalClose ?? ''));
            });
            this.shadow.querySelectorAll('.cwe-modal-backdrop').forEach((backdrop) => {
                backdrop.addEventListener('click', (event) => {
                    if (event.target === backdrop)
                        this.closeModal(backdrop.dataset.modal ?? '');
                });
            });
            this.element('[data-history-list]').addEventListener('click', (event) => {
                const item = event.target instanceof Element
                    ? event.target.closest('[data-history-id]')
                    : null;
                const id = item?.dataset.historyId;
                if (!id)
                    return;
                const actionBtn = event.target instanceof Element
                    ? event.target.closest('[data-history-action]')
                    : null;
                // 与主脚本一致：点击条目本身恢复，图标按钮分别触发重新下载/删除
                if (actionBtn) {
                    const action = actionBtn.dataset.historyAction;
                    if (action === 'download')
                        this.callbacks.onHistoryDownload(id);
                    else if (action === 'delete')
                        this.callbacks.onHistoryDelete(id);
                }
                else {
                    this.callbacks.onHistoryRestore(id);
                }
            });
            window.addEventListener('keydown', (event) => {
                if (event.key !== 'Escape')
                    return;
                const openModal = this.shadow.querySelector('.cwe-modal-backdrop[data-open="true"]');
                if (openModal)
                    this.closeModal(openModal.dataset.modal ?? '');
                else if (this.panel().dataset.open === 'true')
                    this.close();
            });
        }
        enableDragging() {
            const panel = this.panel();
            const header = this.element('.cwe-header');
            let startX = 0;
            let startY = 0;
            let startLeft = 0;
            let startTop = 0;
            let dragging = false;
            header.addEventListener('pointerdown', (event) => {
                if (!this.settings?.enableDrag || event.button !== 0)
                    return;
                if (event.target instanceof Element && event.target.closest('button'))
                    return;
                const rectangle = panel.getBoundingClientRect();
                dragging = true;
                startX = event.clientX;
                startY = event.clientY;
                startLeft = rectangle.left;
                startTop = rectangle.top;
                header.setPointerCapture(event.pointerId);
                event.preventDefault();
            });
            header.addEventListener('pointermove', (event) => {
                if (!dragging)
                    return;
                const maxLeft = Math.max(0, window.innerWidth - panel.offsetWidth);
                const maxTop = Math.max(0, window.innerHeight - 48);
                const left = Math.min(maxLeft, Math.max(0, startLeft + event.clientX - startX));
                const top = Math.min(maxTop, Math.max(0, startTop + event.clientY - startY));
                panel.style.left = `${left}px`;
                panel.style.top = `${top}px`;
                panel.style.right = 'auto';
            });
            const finish = (event) => {
                if (!dragging)
                    return;
                dragging = false;
                if (header.hasPointerCapture(event.pointerId))
                    header.releasePointerCapture(event.pointerId);
                const rectangle = panel.getBoundingClientRect();
                this.callbacks.onPanelPositionChange({ left: rectangle.left, top: rectangle.top });
            };
            header.addEventListener('pointerup', finish);
            header.addEventListener('pointercancel', finish);
        }
        updateOptionAvailability(changed) {
            const selectedFormat = this.shadow.querySelector('input[name="cwe-format"]:checked')?.value;
            const bank = this.option('bankImport');
            if (selectedFormat !== 'word')
                bank.checked = false;
            if (changed === bank && bank.checked) {
                const word = this.shadow.querySelector('input[name="cwe-format"][value="word"]');
                if (word)
                    word.checked = true;
                this.option('withAnswers').checked = false;
                this.option('withWrong').checked = false;
                this.option('shuffle').checked = false;
            }
            const bankEnabled = bank.checked;
            if (bankEnabled) {
                this.option('withAnswers').checked = false;
                this.option('withWrong').checked = false;
                this.option('shuffle').checked = false;
            }
            // 无参考答案时强制取消勾选（覆盖历史记录恢复的偏好）
            if (!this.answerOptionsAvailable) {
                this.option('withAnswers').checked = false;
                this.option('withWrong').checked = false;
            }
            this.option('withAnswers').disabled = bankEnabled;
            this.option('withWrong').disabled = bankEnabled;
            this.option('shuffle').disabled = bankEnabled;
            // 题库导入仅在 Word 格式下可用，TXT/MD 格式下始终禁用
            this.option('bankImport').disabled = selectedFormat !== 'word';
            this.option('splitByChapter').disabled = this.chapterCount < 2;
            this.updateFilenamePreview();
        }
        // 文件名预览：附加答案时追加“（含答案）”，题库导入时追加“（题库导入）”，取消勾选时移除
        updateFilenamePreview() {
            const input = this.element('[data-field="filename"]');
            const name = input.value.replace(/(?:（含答案）|（题库导入）)+$/, '');
            const suffix = this.option('bankImport').checked
                ? '（题库导入）'
                : this.option('withAnswers').checked
                    ? '（含答案）'
                    : '';
            input.value = name + suffix;
        }
        // 控制“附加参考答案/附加错题汇总”选项的显示与勾选状态
        setAnswerOptionsAvailable(available) {
            this.answerOptionsAvailable = available;
            this.option('withAnswers')
                .closest('.cwe-check')
                ?.classList.toggle('cwe-hidden', !available);
            this.option('withWrong')
                .closest('.cwe-check')
                ?.classList.toggle('cwe-hidden', !available);
            if (!available) {
                this.option('withAnswers').checked = false;
                this.option('withWrong').checked = false;
            }
        }
        refreshActionAvailability(hasResult) {
            const download = this.element('[data-action="download"]');
            const copy = this.element('[data-action="copy"]');
            download.dataset.hasResult = String(hasResult);
            copy.dataset.hasResult = String(hasResult);
            download.disabled = this.busy !== null || !hasResult;
            copy.disabled = this.busy !== null || !hasResult;
        }
        shouldDisableButton(action) {
            if (action === 'chapters')
                return this.chapterCount < 2;
            if (action === 'download' || action === 'copy') {
                return (this.element(`[data-action="${action}"]`).dataset.hasResult !== 'true');
            }
            return false;
        }
        saveSettings() {
            if (!this.settings)
                return;
            const rememberPanelPosition = this.element('[data-setting="rememberPanelPosition"]').checked;
            const settings = {
                ...this.settings,
                theme: this.element('[data-setting="theme"]')
                    .value,
                shortcut: (0, shortcut_2.parseShortcut)(this.element('[data-setting="shortcut"]').value, this.settings.shortcut),
                hideShortcut: (0, shortcut_2.parseShortcut)(this.element('[data-setting="hideShortcut"]').value, this.settings.hideShortcut),
                enableDrag: this.element('[data-setting="enableDrag"]').checked,
                rememberPanelPosition,
                autoExtractOnLoad: this.element('[data-setting="autoExtractOnLoad"]')
                    .checked,
                panelPosition: rememberPanelPosition ? this.settings.panelPosition : null,
                exportPreferences: this.readExportPreferences(),
            };
            this.callbacks.onSettingsSave(settings);
            this.closeModal('settings');
            this.setStatus('设置已保存', 'success');
        }
        applyResolvedTheme() {
            const mode = this.settings?.theme ?? 'auto';
            this.mediaQuery?.removeEventListener('change', this.onColorSchemeChange);
            this.mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
            if (mode === 'auto')
                this.mediaQuery.addEventListener('change', this.onColorSchemeChange);
            const dark = mode === 'dark' || (mode === 'auto' && this.mediaQuery.matches);
            this.host.dataset.theme = dark ? 'dark' : 'light';
        }
        applyPosition(position) {
            const panel = this.panel();
            if (!position || !this.settings?.rememberPanelPosition) {
                panel.style.removeProperty('left');
                panel.style.removeProperty('top');
                panel.style.removeProperty('right');
                return;
            }
            const left = Math.min(Math.max(0, position.left), Math.max(0, window.innerWidth - panel.offsetWidth));
            const top = Math.min(Math.max(0, position.top), Math.max(0, window.innerHeight - 48));
            panel.style.left = `${left}px`;
            panel.style.top = `${top}px`;
            panel.style.right = 'auto';
        }
        setAllChapters(checked) {
            this.shadow.querySelectorAll('[data-chapter-list] input').forEach((input) => {
                input.checked = checked;
            });
            this.refreshChapterSelection();
        }
        openModal(name) {
            this.closeAllModals();
            const modal = this.shadow.querySelector(`[data-modal="${name}"]`);
            if (modal)
                modal.dataset.open = 'true';
        }
        closeModal(name) {
            const modal = this.shadow.querySelector(`[data-modal="${name}"]`);
            if (modal)
                modal.dataset.open = 'false';
        }
        closeAllModals() {
            this.shadow.querySelectorAll('[data-modal]').forEach((modal) => {
                modal.dataset.open = 'false';
            });
        }
        shell() {
            return this.element('.cwe-shell');
        }
        launcher() {
            return this.element('.cwe-launcher');
        }
        panel() {
            return this.element('.cwe-panel');
        }
        option(name) {
            return this.element(`[data-option="${name}"]`);
        }
        stat(name) {
            return this.element(`[data-stat="${name}"]`);
        }
        setStatVisible(name, visible) {
            this.stat(name).closest('.cwe-stat')?.classList.toggle('cwe-hidden', !visible);
        }
        element(selector) {
            return (0, assert_1.queryRequired)(this.shadow, selector);
        }
    }
    exports.PanelView = PanelView;
});
define("application/app-controller", ["require", "exports", "extractors/dom-diagnostics", "infrastructure/clipboard-service", "infrastructure/download-service", "infrastructure/history-repository", "infrastructure/logger", "infrastructure/settings-repository", "utils/async", "app-config", "utils/hash", "utils/dom", "utils/shortcut", "utils/text", "application/extraction-service", "application/task-tab-locator"], function (require, exports, dom_diagnostics_1, clipboard_service_1, download_service_1, history_repository_1, logger_2, settings_repository_1, async_4, app_config_2, hash_6, dom_10, shortcut_3, text_9, extraction_service_2, task_tab_locator_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.AppController = void 0;
    // 页面存在题目卡时等待题目加载的上限（切换任务卡后答题页还要 AJAX 渲染）
    const QUESTION_SETTLE_TIMEOUT_MS = 12000;
    // 页面没有题目卡时不再空等：只留出页面自身渲染完成的时间（与原 4.5s 等待的容错度一致）
    const PAGE_IDLE_TIMEOUT_MS = 4500;
    class AppController {
        constructor(view, extractionService, chapterService, exportService, frameBridge, taskTabs = new task_tab_locator_2.TaskTabLocator(), settingsRepository = new settings_repository_1.SettingsRepository(), historyRepository = new history_repository_1.HistoryRepository(), downloadService = new download_service_1.DownloadService(), clipboardService = new clipboard_service_1.ClipboardService()) {
            this.view = view;
            this.extractionService = extractionService;
            this.chapterService = chapterService;
            this.exportService = exportService;
            this.frameBridge = frameBridge;
            this.taskTabs = taskTabs;
            this.settingsRepository = settingsRepository;
            this.historyRepository = historyRepository;
            this.downloadService = downloadService;
            this.clipboardService = clipboardService;
            this.currentResult = null;
            this.observer = null;
            this.logger = new logger_2.Logger();
            this.refreshChaptersDebounced = (0, async_4.debounce)(() => this.refreshChapterAvailability(), 700);
            this.onShortcut = (event) => {
                if ((0, dom_10.isEditableTarget)(event.target))
                    return;
                if ((0, shortcut_3.matchesShortcut)(event, this.settings.hideShortcut)) {
                    event.preventDefault();
                    const hidden = this.view.toggleGlobalVisibility();
                    if (!hidden)
                        this.view.setStatus('浮窗已恢复', 'success');
                    return;
                }
                if ((0, shortcut_3.matchesShortcut)(event, this.settings.shortcut)) {
                    event.preventDefault();
                    this.view.open();
                    void this.extractCurrent();
                }
            };
            this.settings = this.settingsRepository.load();
            this.history = this.historyRepository.load();
        }
        start() {
            this.frameBridge.start();
            this.view.setCallbacks({
                onExtract: () => void this.extractCurrent(),
                onOpenChapters: () => this.openChapterDialog(),
                onExtractChapters: (indexes) => void this.extractChapters(indexes),
                onDownload: () => void this.downloadCurrent(),
                onCopy: () => void this.copyCurrent(),
                onDiagnose: () => void this.copyDiagnostics(),
                onPreferencesChange: (preferences) => this.savePreferences(preferences),
                onSettingsSave: (settings) => this.saveSettings(settings),
                onHistoryRestore: (id) => this.restoreHistory(id),
                onHistoryDownload: (id) => void this.downloadHistory(id),
                onHistoryDelete: (id) => this.deleteHistory(id),
                onPanelPositionChange: (position) => this.savePanelPosition(position),
            });
            this.view.applySettings(this.settings);
            this.view.renderHistory(this.history);
            this.refreshChapterAvailability();
            window.addEventListener('keydown', this.onShortcut, true);
            if (document.body) {
                this.observer = new MutationObserver(() => this.refreshChaptersDebounced());
                this.observer.observe(document.body, { childList: true, subtree: true });
            }
            // 自动提取：进入页面时仅提取一次，后续页面变化不再自动提取
            if (this.settings.autoExtractOnLoad) {
                window.setTimeout(() => void this.extractCurrent(), 800);
            }
        }
        stop() {
            window.removeEventListener('keydown', this.onShortcut, true);
            this.observer?.disconnect();
            this.observer = null;
            this.frameBridge.stop();
            this.view.destroy();
        }
        async extractCurrent() {
            this.view.setBusy('extract', '正在识别当前页面和已加载的 iframe…');
            try {
                this.frameBridge.requestRefresh();
                this.acceptResult(await this.resolveCurrentResult());
            }
            catch (error) {
                this.logger.warn('Extraction failed', error);
                this.view.setStatus((0, text_9.humanizeError)(error), 'error');
            }
            finally {
                this.view.setBusy(null);
                this.refreshChapterAvailability();
            }
        }
        /**
         * 取当前页面的题目。
         *
         * 默认停留在「视频/文档」卡时，题目 iframe 的真实地址还留在 `_src` 上，页面里没有题目 DOM：
         * 此时采样等待并在需要时补切题目卡，避免固定等满整段超时。
         *
         * 已经取到题目时也不是立刻收工：答题页会分几批渲染，中途取到的结果可能只有题干、
         * 选项还没补上，这类「选择题缺选项」的结果要交给采样等待补齐。
         */
        async resolveCurrentResult() {
            const immediate = this.extractionService.extractAccessibleDocuments();
            if (immediate && immediate.questions.length > 0 && (0, extraction_service_2.incompleteChoiceCount)(immediate) === 0) {
                return immediate;
            }
            const guard = new task_tab_locator_2.QuestionTabGuard(this.taskTabs);
            // 页面根本没有题目卡时，结构不会自己变出题目，只留出页面自身渲染的时间
            const timeoutMs = this.taskTabs.questionTabIndex() === null ? PAGE_IDLE_TIMEOUT_MS : QUESTION_SETTLE_TIMEOUT_MS;
            const settled = await this.extractionService.settleQuestions({
                timeoutMs,
                repair: () => {
                    if (guard.ensure()) {
                        this.view.setStatus('已自动切换到题目所在任务卡，正在等待题目加载…', 'neutral');
                    }
                },
            });
            if (!settled)
                throw new Error('当前页面未识别到题目，请确认题目已经加载完成');
            return settled;
        }
        openChapterDialog() {
            const chapters = this.chapterService.listChapters();
            this.view.setChapterCount(chapters.length);
            if (chapters.length < 2) {
                this.view.setStatus('当前页面未识别到可批量遍历的章节列表', 'warning');
                return;
            }
            this.view.showChapterDialog(chapters);
        }
        async extractChapters(indexes) {
            if (indexes.length === 0) {
                this.view.setStatus('请至少选择一个章节', 'warning');
                return;
            }
            this.view.setBusy('chapters', `准备提取 ${indexes.length} 个章节…`);
            try {
                const result = await this.chapterService.extractSelected(indexes, (progress) => {
                    this.view.updateChapterProgress(progress);
                });
                this.acceptResult(result);
                this.view.closeChapterDialog();
            }
            catch (error) {
                this.logger.warn('Chapter extraction failed', error);
                this.view.setStatus((0, text_9.humanizeError)(error), 'error');
            }
            finally {
                this.view.setBusy(null);
                this.refreshChapterAvailability();
            }
        }
        acceptResult(result) {
            this.currentResult = result;
            this.view.renderResult(result);
            const options = this.view.readExportOptions();
            this.view.applyExportOptions({ ...options, filename: (0, text_9.sanitizeFilename)(result.title) });
        }
        async downloadCurrent() {
            if (!this.currentResult) {
                this.view.setStatus('请先提取题目', 'warning');
                return;
            }
            const options = this.view.readExportOptions();
            this.view.setBusy('download', options.format === 'word' ? '正在生成 Word 文档并处理图片…' : '正在生成文件…');
            try {
                const artifacts = await this.exportService.createArtifacts(this.currentResult, options);
                await this.downloadService.downloadMany(artifacts);
                const failedImages = artifacts.reduce((total, artifact) => total + (artifact.failedImages ?? 0), 0);
                this.addHistory(this.currentResult, options);
                // 下载状态提示与主脚本保持一致
                const warnMsg = failedImages > 0 ? `（${failedImages} 张图片加载失败）` : '';
                let message;
                if (artifacts.length > 1) {
                    // 勾选「按章节拆分文件」时，各种格式都是一章一个文件
                    message = `已下载 ${artifacts.length} 个章节文件`;
                }
                else if (options.format === 'word') {
                    const withAnswers = options.bankImport || options.withAnswers;
                    const withWrong = !options.bankImport && options.withWrong;
                    message = options.bankImport
                        ? '题库导入格式已下载'
                        : 'Word 试卷' +
                            (withAnswers ? '（含答案）' : '') +
                            (withWrong ? '（含错题）' : '') +
                            '已下载';
                }
                else {
                    message = '文件已下载';
                }
                this.view.setStatus(message + warnMsg, failedImages > 0 ? 'warning' : 'success');
            }
            catch (error) {
                this.logger.error('Export failed', error);
                const prefix = options.format === 'word' ? 'Word 导出失败: ' : '导出失败：';
                this.view.setStatus(`${prefix}${(0, text_9.humanizeError)(error)}`, 'error');
            }
            finally {
                this.view.setBusy(null);
            }
        }
        async copyCurrent() {
            if (!this.currentResult) {
                this.view.setStatus('请先提取题目', 'warning');
                return;
            }
            const options = this.view.readExportOptions();
            // 原版行为：Word 格式不支持复制
            if (options.format === 'word') {
                this.view.setStatus('Word 格式不支持复制，请使用下载', 'warning');
                return;
            }
            this.view.setBusy('copy', '正在生成可复制文本…');
            try {
                const text = this.exportService.previewText(this.currentResult, options);
                await this.clipboardService.writeText(text);
                this.view.setStatus('已复制到剪贴板', 'success');
            }
            catch (error) {
                this.view.setStatus(`复制失败：${(0, text_9.humanizeError)(error)}`, 'error');
            }
            finally {
                this.view.setBusy(null);
            }
        }
        /**
         * 复制提取诊断信息。
         *
         * 「提取到题干、提取不到选项」在离线快照上复现不出来（快照是渲染完成后的 DOM），
         * 所以把当前页面各层文档的题目/选项结构一次性复制出来，供定位成因：
         * 选择器没命中、选项在题目容器之外，还是选项节点本身就解析为空。
         */
        async copyDiagnostics() {
            try {
                await this.clipboardService.writeText((0, dom_diagnostics_1.buildExtractionDiagnostics)({ version: app_config_2.APP_VERSION }));
                this.view.setStatus('诊断信息已复制到剪贴板，直接粘贴发给维护者即可', 'success');
            }
            catch (error) {
                this.view.setStatus(`复制诊断失败：${(0, text_9.humanizeError)(error)}`, 'error');
            }
        }
        addHistory(result, options) {
            const createdAt = new Date().toISOString();
            const entry = {
                id: `history-${(0, hash_6.stableHash)(`${createdAt}|${result.title}|${result.questions.length}`)}`,
                createdAt,
                title: options.filename || result.title,
                result,
                options,
            };
            this.history = this.historyRepository.add(entry);
            this.view.renderHistory(this.history);
        }
        restoreHistory(id) {
            const entry = this.history.find((candidate) => candidate.id === id);
            if (!entry) {
                this.view.setStatus('历史记录不存在或已被清理', 'warning');
                return;
            }
            this.currentResult = entry.result;
            this.view.renderResult(entry.result);
            this.view.applyExportOptions(entry.options);
            this.view.closeHistory();
            this.view.open();
            this.view.setStatus(`已恢复历史记录：${entry.title}`, 'success');
        }
        async downloadHistory(id) {
            const entry = this.history.find((candidate) => candidate.id === id);
            if (!entry)
                return;
            this.view.closeHistory();
            this.view.setBusy('download', '正在重新生成历史文件…');
            try {
                const artifacts = await this.exportService.createArtifacts(entry.result, entry.options);
                await this.downloadService.downloadMany(artifacts);
                this.view.setStatus(`历史文件已重新下载`, 'success');
            }
            catch (error) {
                this.view.setStatus(`重新下载失败：${(0, text_9.humanizeError)(error)}`, 'error');
            }
            finally {
                this.view.setBusy(null);
            }
        }
        deleteHistory(id) {
            this.history = this.historyRepository.remove(id);
            this.view.renderHistory(this.history);
            this.view.setStatus('历史记录已删除', 'success');
        }
        savePreferences(preferences) {
            this.settings = { ...this.settings, exportPreferences: preferences };
            this.settingsRepository.save(this.settings);
        }
        saveSettings(settings) {
            this.settings = settings;
            this.settingsRepository.save(settings);
            this.view.applySettings(settings);
        }
        savePanelPosition(position) {
            if (!this.settings.rememberPanelPosition)
                return;
            this.settings = { ...this.settings, panelPosition: position };
            this.settingsRepository.save(this.settings);
            this.view.applySettings(this.settings);
        }
        refreshChapterAvailability() {
            try {
                this.view.setChapterCount(this.chapterService.listChapters().length);
            }
            catch (error) {
                this.logger.debug('Chapter discovery skipped', error);
                this.view.setChapterCount(0);
            }
        }
    }
    exports.AppController = AppController;
});
define("main", ["require", "exports", "application/app-controller", "application/chapter-extraction-service", "application/chapter-locator", "application/extraction-service", "application/task-tab-locator", "exporters/export-service", "extractors/composite-extractor", "infrastructure/frame-bridge", "ui/panel-view"], function (require, exports, app_controller_1, chapter_extraction_service_1, chapter_locator_2, extraction_service_3, task_tab_locator_3, export_service_1, composite_extractor_4, frame_bridge_2, panel_view_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    function inIframe() {
        try {
            return window.self !== window.top;
        }
        catch {
            return true;
        }
    }
    function bootTopWindow() {
        const frameBridge = new frame_bridge_2.FrameBridge();
        const extractionService = new extraction_service_3.ExtractionService(new composite_extractor_4.CompositeExtractor(), frameBridge);
        // 章节切换与任务卡切换共用同一个定位器实例
        const taskTabs = new task_tab_locator_3.TaskTabLocator();
        const chapterService = new chapter_extraction_service_1.ChapterExtractionService(extractionService, new chapter_locator_2.ChapterLocator(), taskTabs);
        const controller = new app_controller_1.AppController(new panel_view_1.PanelView(), extractionService, chapterService, new export_service_1.ExportService(), frameBridge, taskTabs);
        controller.start();
    }
    function boot() {
        if (inIframe()) {
            new frame_bridge_2.FrameAgent().start();
            return;
        }
        if (document.body)
            bootTopWindow();
        else
            document.addEventListener('DOMContentLoaded', bootTopWindow, { once: true });
    }
    boot();
});


  load('main');
})();
