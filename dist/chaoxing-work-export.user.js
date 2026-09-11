// ==UserScript==
// @name         超星学习通作业/考试一键提取导出 Word 文档
// @namespace    https://github.com/chenhuilin2/chaoxing-work-export
// @version      3.0.0
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
    exports.APP_VERSION = '3.0.0';
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
            includeAnalysis: true,
            shuffle: false,
            bankImport: false,
            splitByChapter: false,
        },
        enableDrag: false,
        rememberPanelPosition: true,
        panelPosition: null,
        autoExtractOnLoad: false,
    };
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
                output += `${globalNum}. ${(0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.questionContent)(q))}\n`;
                if (q.options && q.options.length > 0) {
                    for (const opt of q.options) {
                        output += `${opt.letter}. ${(0, legacy_bridge_1.formatRichForText)((0, legacy_bridge_1.optionContent)(opt))}\n`;
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
                    const parts = answer
                        .split('；')
                        .map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
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
                output += `**${globalNum}.** ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.questionContent)(q))}\n\n`;
                if (q.options && q.options.length > 0) {
                    for (const opt of q.options) {
                        output += `- ${opt.letter}. ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.optionContent)(opt))}\n`;
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
                    const parts = answer
                        .split('；')
                        .map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
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
                output += `**${globalNum}.** ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.questionContent)(q))}\n\n`;
                output += `- 我的答案: ${q.myAnswer || '无'}\n`;
                output += `- 正确答案: ${(0, legacy_bridge_2.formatRichForMD)((0, legacy_bridge_2.answerContent)(q)) || '（未找到答案）'}\n\n`;
            }
        }
        return output.replace(/\n+$/, '');
    }
});
define("exporters/word-exporter", ["require", "exports", "exporters/legacy-bridge"], function (require, exports, legacy_bridge_3) {
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
    async function generateWordBlob(results, typeOrder, title, withAnswers, withWrong, bankImport) {
        const { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, HeadingLevel, PageBreak } = docx;
        const FONT = 'Microsoft YaHei';
        const COLOR = {
            title: '1F2937',
            muted: '6B7280',
            border: 'E5E7EB',
            answer: '00A870',
            wrong: 'DC2626',
            analysisBg: 'F7F9FC',
            type: '64748B',
        };
        const typeHeaders = {
            单选: '一、单项选择题',
            多选: '二、多项选择题',
            填空: '三、填空题',
            判断: '四、判断题',
            简答: '五、简答题',
        };
        const bankTypeLabels = {
            单选: '【单选题】',
            多选: '【多选题】',
            填空: '【填空题】',
            判断: '【判断题】',
            简答: '【简答题】',
        };
        // 题库导入格式：生成学习通智能导入兼容的 Word 文档
        if (bankImport) {
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
            for (const qtype of typeOrder) {
                const questions = results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                const prefix = bankTypeLabels[qtype];
                for (const q of questions) {
                    qNum++;
                    // 题干（题号 + 题型标签 + 题干内容），括号归一化
                    const stemText = (0, legacy_bridge_3.formatRichForText)((0, legacy_bridge_3.questionContent)(q));
                    const stem = prefix + stemText.replace(/\(\s{2,}\)/g, '（ ）').replace(/（\s{2,}）/g, '（ ）');
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `${qNum}.${stem}`, font: FONT, size: 22 })],
                        spacing: { after: 40 },
                    }));
                    // 题干中的图片
                    const stemContent = (0, legacy_bridge_3.questionContent)(q);
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
                        const runs = await buildRichRuns((0, legacy_bridge_3.optionContent)(opt), `${opt.letter}. `);
                        children.push(new Paragraph({
                            children: runs,
                            indent: { left: 400, hanging: 200 },
                            spacing: { after: 40 },
                        }));
                    }
                    // 答案
                    const answer = (0, legacy_bridge_3.formatRichForText)((0, legacy_bridge_3.answerContent)(q)).trim();
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
                                    top: convertMillimetersToTwip(20),
                                    bottom: convertMillimetersToTwip(20),
                                    left: convertMillimetersToTwip(25),
                                    right: convertMillimetersToTwip(25),
                                },
                            },
                        },
                        children,
                    },
                ],
            });
            return await Packer.toBlob(doc);
        }
        const children = [];
        // 标题
        children.push(new Paragraph({
            children: [
                new TextRun({ text: title || '试卷', font: FONT, size: 32, bold: true, color: '2563EB' }),
            ],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 120 },
        }));
        // 统计摘要
        let totalQ = 0;
        const typeCount = {};
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            typeCount[qtype] = questions.length;
            totalQ += questions.length;
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
        let qNum = 0;
        for (const qtype of typeOrder) {
            const questions = results[qtype];
            if (!questions || questions.length === 0)
                continue;
            const header = typeHeaders[qtype] || qtype;
            const count = questions.length;
            children.push(new Paragraph({
                children: [
                    new TextRun({
                        text: `${header}（本大题共${count}小题）`,
                        font: FONT,
                        size: 28,
                        bold: true,
                        color: COLOR.title,
                    }),
                ],
                spacing: { after: 120 },
            }));
            for (const q of questions) {
                qNum++;
                const stemContent = (0, legacy_bridge_3.questionContent)(q);
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
                                children: await buildRichRuns((0, legacy_bridge_3.optionContent)(opt), `${opt.letter}. `),
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
        // 答案页
        if (withAnswers) {
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
            let aNum = 0;
            for (const qtype of typeOrder) {
                const questions = results[qtype];
                if (!questions || questions.length === 0)
                    continue;
                const header = typeHeaders[qtype] || qtype;
                children.push(new Paragraph({
                    children: [new TextRun({ text: header, font: FONT, size: 28, bold: true, color: COLOR.title })],
                    spacing: { after: 120 },
                }));
                for (const q of questions) {
                    aNum++;
                    // 清洗答案文本，去除“正确答案:/我的答案:”等标签组合
                    const answerContentArr = (0, legacy_bridge_3.answerContent)(q);
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
                        children: [new TextRun({ text: `${aNum}. `, font: FONT, size: 22, color: COLOR.title }), ...answerRuns],
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
        // 错题汇总（Word 试卷）
        if (withWrong) {
            let hasWrong = false;
            for (const qtype of typeOrder) {
                const questions = results[qtype];
                if (!questions)
                    continue;
                for (const q of questions) {
                    if (q.isWrong) {
                        hasWrong = true;
                        break;
                    }
                }
                if (hasWrong)
                    break;
            }
            if (hasWrong) {
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
                let globalNum = 0;
                for (const qtype of typeOrder) {
                    const questions = results[qtype];
                    if (!questions || questions.length === 0)
                        continue;
                    if (qtype === '简答') {
                        globalNum += questions.length;
                        continue;
                    }
                    const header = typeHeaders[qtype] || qtype;
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
                            new TextRun({ text: header, font: FONT, size: 28, bold: true, color: COLOR.title }),
                        ],
                        spacing: { after: 120 },
                    }));
                    for (const q of questions) {
                        globalNum++;
                        if (!q.isWrong)
                            continue;
                        children.push(new Paragraph({
                            children: await buildRichRuns((0, legacy_bridge_3.questionContent)(q), `${globalNum}. `),
                            spacing: { before: 160, after: 40 },
                        }));
                        const options = q.options || [];
                        if (options.length > 0) {
                            for (const opt of options) {
                                children.push(new Paragraph({
                                    children: await buildRichRuns((0, legacy_bridge_3.optionContent)(opt), `${opt.letter}. `),
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
                        const answerContentArr = (0, legacy_bridge_3.answerContent)(q);
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
define("exporters/export-service", ["require", "exports", "exporters/legacy-bridge", "exporters/text-formatter", "exporters/markdown-formatter", "exporters/word-exporter"], function (require, exports, legacy_bridge_4, text_formatter_1, markdown_formatter_1, word_exporter_1) {
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
            const shuffled = (0, legacy_bridge_4.shuffleQuestions)(legacy.results, legacy.typeOrder);
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
        const activeResults = doShuffle ? (0, legacy_bridge_4.shuffleQuestions)(results, typeOrder) : results;
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
            const legacy = (0, legacy_bridge_4.toLegacyResult)(sourceResult);
            // 原版规则：文件名输入去掉扩展名后作为基础名
            const baseFilename = (options.filename || legacy.title || '学习通题目').replace(/\.(txt|md|docx)$/, '');
            // Word 试卷导出（原版逻辑：题库导入必带答案并禁用打乱/错题；Word 暂不支持分章节，合并导出）
            if (options.format === 'word') {
                const isBankImport = options.bankImport;
                const doShuffle = !isBankImport && options.shuffle;
                const withWrong = !isBankImport && options.withWrong;
                const activeResults = doShuffle
                    ? (0, legacy_bridge_4.shuffleQuestions)(legacy.results, legacy.typeOrder)
                    : legacy.results;
                window.__xxt_failed_image_count = 0;
                const blob = await (0, word_exporter_1.generateWordBlob)(activeResults, legacy.typeOrder, legacy.title, isBankImport || options.withAnswers, withWrong, isBankImport);
                const failedImages = window.__xxt_failed_image_count || 0;
                return [
                    {
                        blob,
                        filename: `${baseFilename}.docx`,
                        mimeType: word_exporter_1.WORD_MIME,
                        failedImages,
                    },
                ];
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
                    const chapterData = (0, legacy_bridge_4.toLegacyChapter)(chapter);
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
            return getOutputText((0, legacy_bridge_4.toLegacyResult)(result), options);
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
define("utils/dom", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.textOf = textOf;
    exports.firstMatch = firstMatch;
    exports.allMatches = allMatches;
    exports.parseLeadingNumber = parseLeadingNumber;
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
        const match = value.match(/^\s*(\d+)\s*[.、．]/u);
        if (!match?.[1])
            return undefined;
        const parsed = Number.parseInt(match[1], 10);
        return Number.isFinite(parsed) ? parsed : undefined;
    }
    function isEditableTarget(target) {
        if (!(target instanceof Element))
            return false;
        return Boolean(target.closest('input, textarea, select, [contenteditable="true"]'));
    }
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
define("extractors/rich-content", ["require", "exports", "utils/text"], function (require, exports, text_1) {
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
        if (!(element instanceof HTMLElement))
            return [];
        let background = element.style.backgroundImage;
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
    function extractRichContent(element) {
        if (!element)
            return [];
        const parts = [];
        const walk = (node, style) => {
            if (node.nodeType === Node.TEXT_NODE) {
                parts.push({ type: 'text', text: node.nodeValue ?? '', ...style });
                return;
            }
            if (!(node instanceof Element))
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
            if (fallback)
                normalized = [{ type: 'text', text: fallback }];
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
        ].join(',')));
    }
});
define("extractors/option-parser", ["require", "exports", "utils/text", "extractors/rich-content"], function (require, exports, text_3, rich_content_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
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
    function parseOptions(container, selectors) {
        for (const selector of selectors) {
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
    const CORRECT_SELECTORS = [
        '.mark_answer .mark_key .colorGreen .stuAnswerContent',
        '.mark_answer .mark_key .colorGreen',
        '.newAnswerBx .correctAnswerBx .answerCon',
        '.correctAnswerBx .answerCon',
        '.correctAnswerContent',
        '.correct-answer',
        '.rightAnswer',
        '[data-role="correct-answer"]',
    ];
    const USER_SELECTORS = [
        '.mark_answer .mark_key .colorDeep .stuAnswerContent',
        '.mark_answer .mark_key .colorDeep',
        '.newAnswerBx .myAnswerBx .answerCon',
        '.myAnswerBx .answerCon',
        '.myAnswer .answerCon',
        '.my-answer',
        '[data-role="user-answer"]',
    ];
    const ANALYSIS_SELECTORS = [
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
define("extractors/generic-exercise-extractor", ["require", "exports", "utils/dom", "utils/array", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-type", "extractors/rich-content"], function (require, exports, dom_2, array_1, answer_comparison_2, common_answer_1, option_parser_2, question_factory_1, question_type_1, rich_content_5) {
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
    const STEM_SELECTORS = [
        '[data-role="stem"]',
        '.question-stem',
        '.questionStem',
        '.subject-title',
        '.subject',
        '.qtContent',
        '.mark_name',
        '.Zy_TItle .qtContent',
    ];
    class GenericExerciseExtractor {
        constructor() {
            this.id = 'generic-exercise';
            this.confidence = 40;
        }
        supports(root) {
            return (0, dom_2.allMatches)(root, CONTAINER_SELECTORS).length > 0;
        }
        extract(context) {
            const containers = (0, array_1.uniqueBy)((0, dom_2.allMatches)(context.root, CONTAINER_SELECTORS), (element) => {
                return element.getAttribute('data-question-id') ?? `${element.tagName}:${(0, dom_2.textOf)(element).slice(0, 80)}`;
            });
            const questions = [];
            for (const container of containers) {
                const type = (0, question_type_1.detectQuestionType)(container.getAttribute('typeName'), container.getAttribute('data-question-type'), (0, dom_2.textOf)(container.querySelector('.question-type, .type_tit, .newTestType'))) ?? (0, question_type_1.inferQuestionType)(container);
                if (!type)
                    continue;
                let stemElement = null;
                for (const selector of STEM_SELECTORS) {
                    const candidate = container.querySelector(selector);
                    if (!candidate)
                        continue;
                    const content = (0, rich_content_5.stripQuestionPrefix)((0, rich_content_5.extractRichContent)(candidate));
                    if (!(0, rich_content_5.isRichContentEmpty)(content)) {
                        stemElement = candidate;
                        break;
                    }
                }
                if (!stemElement)
                    continue;
                const question = (0, question_factory_1.createQuestion)({
                    number: (0, dom_2.parseLeadingNumber)((0, dom_2.textOf)(stemElement)),
                    type,
                    typeMeta: (0, dom_2.textOf)(container.querySelector('.question-type, .colorShallow')) || undefined,
                    stem: (0, rich_content_5.stripQuestionPrefix)((0, rich_content_5.extractRichContent)(stemElement)),
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
define("extractors/mark-item-extractor", ["require", "exports", "utils/dom", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-type", "extractors/rich-content"], function (require, exports, dom_3, answer_comparison_3, common_answer_2, option_parser_3, question_factory_2, question_type_2, rich_content_6) {
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
                const sectionType = (0, question_type_2.detectQuestionType)((0, dom_3.textOf)(section.querySelector('.type_tit')));
                if (!sectionType)
                    return;
                section.querySelectorAll('.questionLi').forEach((container) => {
                    const stemElement = container.querySelector('.qtContent, .mark_name, .question-stem');
                    const type = (0, question_type_2.detectQuestionType)(container.getAttribute('typeName'), container.getAttribute('data-question-type'), (0, dom_3.textOf)(container.querySelector('.colorShallow'))) ?? sectionType;
                    const rawStem = (0, rich_content_6.extractRichContent)(stemElement);
                    const question = (0, question_factory_2.createQuestion)({
                        number: (0, dom_3.parseLeadingNumber)((0, dom_3.textOf)(stemElement)),
                        type,
                        typeMeta: (0, dom_3.textOf)(container.querySelector('.colorShallow')) || undefined,
                        stem: (0, rich_content_6.stripQuestionPrefix)(rawStem),
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
    function resolvePageTitle(root) {
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
define("extractors/question-li-extractor", ["require", "exports", "utils/dom", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-type", "extractors/rich-content"], function (require, exports, dom_4, answer_comparison_4, common_answer_3, option_parser_4, question_factory_3, question_type_3, rich_content_7) {
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
                const type = (0, question_type_3.detectQuestionType)(container.getAttribute('typeName'), container.getAttribute('typename'), container.getAttribute('data-question-type'), (0, dom_4.textOf)(container.querySelector('.type_tit, .question-type, .colorShallow'))) ?? (0, question_type_3.inferQuestionType)(container);
                if (!type)
                    return;
                const stemElement = container.querySelector('.mark_name, .qtContent, .questionStem, .question-stem, .subject, [data-role="stem"]');
                const question = (0, question_factory_3.createQuestion)({
                    number: (0, dom_4.parseLeadingNumber)((0, dom_4.textOf)(stemElement)),
                    type,
                    typeMeta: (0, dom_4.textOf)(container.querySelector('.colorShallow, .question-type')) || undefined,
                    stem: (0, rich_content_7.stripQuestionPrefix)((0, rich_content_7.extractRichContent)(stemElement)),
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
define("extractors/timu-extractor", ["require", "exports", "utils/dom", "extractors/answer-comparison", "extractors/common-answer", "extractors/option-parser", "extractors/question-factory", "extractors/question-type", "extractors/rich-content"], function (require, exports, dom_5, answer_comparison_5, common_answer_4, option_parser_5, question_factory_4, question_type_4, rich_content_8) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.TiMuExtractor = void 0;
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
                currentType =
                    (0, question_type_4.detectQuestionType)((0, dom_5.textOf)(area.querySelector('.newTestType')), area.getAttribute('data-question-type')) ?? currentType;
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
            const stemElement = title.querySelector('.qtContent, .question-content, .mark_name');
            const type = (0, question_type_4.detectQuestionType)((0, dom_5.textOf)(title.querySelector('.newZy_TItle')), container.getAttribute('typeName'), container.getAttribute('data-question-type')) ??
                inheritedType ??
                (0, question_type_4.inferQuestionType)(container);
            if (!type)
                return null;
            return (0, question_factory_4.createQuestion)({
                number: (0, dom_5.parseLeadingNumber)((0, dom_5.textOf)(title.querySelector('i.fl')) || (0, dom_5.textOf)(stemElement)),
                type,
                typeMeta: (0, dom_5.textOf)(title.querySelector('.newZy_TItle')) || undefined,
                stem: (0, rich_content_8.stripQuestionPrefix)((0, rich_content_8.extractRichContent)(stemElement)),
                options: (0, option_parser_5.parseOptions)(container, [
                    '.Zy_ulTop.qtDetail > li',
                    '.Zy_ulTop > li',
                    '.answerBg',
                    '.option-list > li',
                ]),
                correctAnswer: (0, common_answer_4.extractCorrectAnswer)(container),
                userAnswer: (0, common_answer_4.extractUserAnswer)(container),
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
define("utils/async", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.delay = delay;
    exports.poll = poll;
    exports.debounce = debounce;
    function delay(milliseconds) {
        return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
    }
    async function poll(producer, accept, options) {
        const deadline = Date.now() + options.timeoutMs;
        let lastValue = null;
        while (Date.now() < deadline) {
            lastValue = await producer();
            if (accept(lastValue))
                return lastValue;
            await delay(options.intervalMs);
        }
        return lastValue !== null && accept(lastValue) ? lastValue : null;
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
define("infrastructure/frame-bridge", ["require", "exports", "extractors/composite-extractor", "utils/async", "utils/hash", "infrastructure/logger"], function (require, exports, composite_extractor_1, async_1, hash_2, logger_1) {
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
        constructor(extractor = new composite_extractor_1.CompositeExtractor()) {
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
define("infrastructure/history-repository", ["require", "exports", "domain/question", "extractors/rich-content", "utils/hash", "infrastructure/safe-storage"], function (require, exports, question_2, rich_content_9, hash_3, safe_storage_1) {
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
            return (0, rich_content_9.textContent)(value);
        return typeof fallback === 'string' ? (0, rich_content_9.textContent)(fallback) : [];
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
                        content: (0, rich_content_9.textContent)(match?.[2] ?? option),
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
            includeAnalysis: legacyOptions.includeAnalysis !== false,
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
            includeAnalysis: firstBoolean(source, ['includeAnalysis', 'withAnalysis', 'appendAnalysis'], settings_1.DEFAULT_SETTINGS.exportPreferences.includeAnalysis),
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
                return settingsValue(current);
            const legacySettings = this.storage.read(LEGACY_KEY);
            const legacyExport = LEGACY_EXPORT_KEYS
                .map((key) => this.storage.read(key))
                .find((value) => value !== null);
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
    const CHAPTER_ITEM_SELECTORS = [
        '.posCatalog_select',
        '.posCatalog_name',
        '.catalog_name',
        '.chapter-item',
        '[data-chapter-id]',
        '[data-id][class*="catalog"]',
    ];
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
            const target = element.matches('a, button, [role="button"]')
                ? element
                : element.querySelector('a, button, [role="button"]') ?? element;
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
                    const filtered = raw.filter((element) => {
                        if (!element.offsetParent && element.getClientRects().length === 0)
                            return false;
                        const title = this.titleOf(element, 0);
                        return title.length > 0 && !/^(?:目录|章节|返回)$/u.test(title);
                    });
                    if (filtered.length > 1)
                        return this.removeNestedDuplicates(filtered);
                }
            }
            return [];
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
                element.innerText;
            const title = (0, text_6.normalizeInlineWhitespace)(explicit ?? '');
            return title || `第 ${index + 1} 章`;
        }
        isActive(element) {
            if (element.getAttribute('aria-current') === 'true')
                return true;
            if (element.getAttribute('aria-selected') === 'true')
                return true;
            return ACTIVE_CLASSES.some((className) => element.classList.contains(className)) ||
                Boolean(element.closest(ACTIVE_CLASSES.map((className) => `.${className}`).join(',')));
        }
    }
    exports.ChapterLocator = ChapterLocator;
});
define("application/extraction-service", ["require", "exports", "domain/question", "extractors/composite-extractor", "infrastructure/frame-bridge", "utils/hash", "utils/async"], function (require, exports, question_3, composite_extractor_2, frame_bridge_1, hash_4, async_2) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ExtractionService = void 0;
    /**
     * Coordinates all extraction sources visible to the current top-level page.
     *
     * A Chaoxing assignment may render directly in the document, in a same-origin
     * iframe, or in a cross-origin iframe. The service deliberately keeps these
     * transport concerns out of individual DOM extractors.
     */
    class ExtractionService {
        constructor(extractor = new composite_extractor_2.CompositeExtractor(), frameBridge = new frame_bridge_1.FrameBridge()) {
            this.extractor = extractor;
            this.frameBridge = frameBridge;
        }
        extract() {
            return this.selectBest([
                ...this.collectDocumentCandidates(document),
                ...this.collectFrameCandidate(this.frameBridge.latest()),
            ]);
        }
        /** Extracts only documents that can be accessed synchronously from this window. */
        extractAccessibleDocuments() {
            return this.selectBest(this.collectDocumentCandidates(document));
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
            return (0, hash_4.stableHash)(`${result.sourceUrl}|${result.title}|${identity}`);
        }
        async waitForChangedResult(options = {}) {
            const afterTimestamp = options.afterTimestamp ?? 0;
            const previousFingerprint = options.previousFingerprint ?? '';
            const timeoutMs = options.timeoutMs ?? 12000;
            const intervalMs = options.intervalMs ?? 300;
            return (0, async_2.poll)(() => {
                const freshFrame = this.frameBridge.latest(afterTimestamp);
                return freshFrame?.result ?? this.extractAccessibleDocuments();
            }, (result) => {
                if (!result || result.questions.length === 0)
                    return false;
                const hasFreshFrame = this.frameBridge.latest(afterTimestamp) !== null;
                return hasFreshFrame || this.fingerprint(result) !== previousFingerprint;
            }, { timeoutMs, intervalMs });
        }
        collectDocumentCandidates(root) {
            const candidates = [];
            const visited = new Set();
            const visit = (current, depth) => {
                if (visited.has(current) || depth > 5)
                    return;
                visited.add(current);
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
                current.querySelectorAll('iframe').forEach((frame) => {
                    try {
                        if (frame.contentDocument)
                            visit(frame.contentDocument, depth + 1);
                    }
                    catch {
                        // Cross-origin frames are handled through FrameBridge/postMessage.
                    }
                });
            };
            visit(root, 0);
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
define("application/chapter-extraction-service", ["require", "exports", "domain/question", "utils/async", "utils/hash", "application/chapter-locator"], function (require, exports, question_4, async_3, hash_5, chapter_locator_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ChapterExtractionService = void 0;
    /** Sequentially visits selected chapters and produces a single aggregate snapshot. */
    class ChapterExtractionService {
        constructor(extractionService, locator = new chapter_locator_1.ChapterLocator()) {
            this.extractionService = extractionService;
            this.locator = locator;
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
                    const startedAt = Date.now();
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
                    await (0, async_3.delay)(alreadyActive ? 150 : 450);
                    const result = alreadyActive
                        ? this.extractionService.extract()
                        : await this.extractionService.waitForChangedResult({
                            afterTimestamp: startedAt,
                            previousFingerprint,
                            timeoutMs: 14000,
                            intervalMs: 350,
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
                if (originalIndex !== null)
                    this.locator.activate(originalIndex);
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
        attachChapter(question, chapter, questionIndex) {
            return {
                ...question,
                id: (0, hash_5.stableHash)(`${chapter.id}|${question.id}|${questionIndex}`),
                chapterId: chapter.id,
                chapterTitle: chapter.title,
            };
        }
        aggregateTitle(chapters) {
            const documentTitle = document.title.trim();
            const suffix = chapters.length === 1 ? chapters[0]?.title : `${chapters.length} 个章节`;
            return [documentTitle || '学习通课程', suffix].filter(Boolean).join(' - ');
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

.cwe-status {
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
.cwe-modal-header { display: flex; align-items: center; gap: 8px; padding: 14px 15px; border-bottom: 1px solid var(--cwe-border); }
.cwe-modal-title { flex: 1; margin: 0; font-size: 14px; font-weight: 800; }
.cwe-modal-body { overflow: auto; padding: 14px 15px; }
.cwe-modal-footer { display: flex; justify-content: flex-end; gap: 8px; padding: 12px 15px; border-top: 1px solid var(--cwe-border); }

.cwe-setting-row { display: grid; grid-template-columns: 118px 1fr; align-items: center; gap: 10px; margin-bottom: 12px; }
.cwe-setting-row:last-child { margin-bottom: 0; }
.cwe-setting-label { color: var(--cwe-text-secondary); font-size: 11.5px; font-weight: 700; }
.cwe-switch-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 9px 0; border-top: 1px solid var(--cwe-border); color: var(--cwe-text-secondary); font-size: 11.5px; }

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

.cwe-chapter-toolbar { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 9px; }
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
define("ui/panel-view", ["require", "exports", "app-config", "utils/assert", "utils/shortcut", "utils/text", "ui/styles"], function (require, exports, app_config_1, assert_1, shortcut_2, text_7, styles_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.PanelView = void 0;
    const NOOP_CALLBACKS = {
        onExtract: () => undefined,
        onOpenChapters: () => undefined,
        onExtractChapters: () => undefined,
        onDownload: () => undefined,
        onCopy: () => undefined,
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
    function template() {
        return `
    <style>${styles_1.PANEL_STYLES}</style>
    <div class="cwe-shell">
      <button class="cwe-launcher" type="button" aria-label="打开题目导出面板">提取题目</button>
      <section class="cwe-panel" data-open="false" role="dialog" aria-label="${app_config_1.APP_NAME}">
        <header class="cwe-header">
          <div class="cwe-header-main">
            <h2 class="cwe-title">学习通题目导出</h2>
          </div>
          <div class="cwe-header-actions">
            <button class="cwe-icon-button" data-action="history" type="button" title="下载历史">${ICONS.history}</button>
            <button class="cwe-icon-button" data-action="settings" type="button" title="设置">${ICONS.settings}</button>
            <button class="cwe-icon-button" data-action="close" type="button" title="关闭">${ICONS.close}</button>
          </div>
        </header>
        <div class="cwe-scroll">
          <div class="cwe-status" data-kind="neutral">进入作业、考试或章节练习页面后开始提取</div>
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
              <label class="cwe-check"><input data-option="includeAnalysis" type="checkbox">附加解析</label>
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
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">提取快捷键</span>
              <input class="cwe-input" data-setting="shortcut" placeholder="Ctrl+Shift+E" />
            </label>
            <label class="cwe-setting-row">
              <span class="cwe-setting-label">隐藏快捷键</span>
              <input class="cwe-input" data-setting="hideShortcut" placeholder="Ctrl+Shift+H" />
            </label>
            <label class="cwe-switch-row"><span>允许拖动面板</span><input data-setting="enableDrag" type="checkbox"></label>
            <label class="cwe-switch-row"><span>记忆面板位置</span><input data-setting="rememberPanelPosition" type="checkbox"></label>
            <label class="cwe-switch-row"><span>打开窗口自动提取</span><input data-setting="autoExtractOnLoad" type="checkbox"></label>
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
        <section class="cwe-modal" role="dialog" aria-modal="true" aria-label="选择章节">
          <header class="cwe-modal-header">
            <h3 class="cwe-modal-title">选择要提取的章节</h3>
            <button class="cwe-icon-button" data-modal-close="chapters" type="button">${ICONS.close}</button>
          </header>
          <div class="cwe-modal-body">
            <div class="cwe-chapter-toolbar">
              <button class="cwe-mini-button" data-action="chapter-all" type="button">全选</button>
              <button class="cwe-mini-button" data-action="chapter-none" type="button">取消全选</button>
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
        setBusy(action, message) {
            this.busy = action;
            const controls = this.shadow.querySelectorAll('[data-action="extract"], [data-action="chapters"], [data-action="download"], [data-action="copy"]');
            controls.forEach((button) => {
                button.disabled = action !== null || this.shouldDisableButton(button.dataset.action ?? '');
            });
            const chapterBusy = action === 'chapters';
            this.element('[data-action="start-chapters"]').disabled = chapterBusy;
            this.shadow.querySelectorAll('[data-chapter-list] input').forEach((input) => {
                input.disabled = chapterBusy;
            });
            const extract = this.element('[data-action="extract"]');
            const chapters = this.element('[data-action="chapters"]');
            const download = this.element('[data-action="download"]');
            const copy = this.element('[data-action="copy"]');
            extract.innerHTML = action === 'extract' ? '<span class="cwe-progress"></span>提取中' : '提取当前页面';
            chapters.innerHTML = action === 'chapters' ? '<span class="cwe-progress"></span>遍历中' : '提取多个章节';
            download.innerHTML = action === 'download' ? '<span class="cwe-progress"></span>生成中' : '下载文件';
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
            this.element('[data-field="filename"]').value = (0, text_7.sanitizeFilename)(result.title);
            // 提取结果无参考答案时，隐藏“附加参考答案/附加错题汇总”选项
            this.setAnswerOptionsAvailable(result.statistics.withCorrectAnswer > 0);
            this.refreshActionAvailability(true);
            this.updateFilenamePreview();
            const chapterMessage = result.chapters?.length ? `，来自 ${result.chapters.length} 个章节` : '';
            this.setStatus(`已提取 ${result.statistics.total} 道题${chapterMessage}`, 'success');
        }
        clearResult() {
            for (const name of TYPE_STAT_KEYS) {
                this.stat(name).textContent = '0';
                this.setStatVisible(name, false);
            }
            // 重置为默认显示，等待下次提取结果决定
            this.setAnswerOptionsAvailable(true);
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
            const format = formatElement?.value === 'txt' || formatElement?.value === 'md' ? formatElement.value : 'word';
            return {
                format,
                filename: this.element('[data-field="filename"]').value,
                withAnswers: this.option('withAnswers').checked,
                withWrong: this.option('withWrong').checked,
                includeAnalysis: this.option('includeAnalysis').checked,
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
                includeAnalysis: options.includeAnalysis,
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
            this.option('includeAnalysis').checked = options.includeAnalysis;
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
            this.applyExportOptions({ ...preferences, filename: this.element('[data-field="filename"]').value });
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
            this.element('[data-chapter-progress]').textContent = `共识别 ${chapters.length} 个章节`;
            this.openModal('chapters');
        }
        updateChapterProgress(progress) {
            const stateLabel = progress.state === 'loading' ? '正在加载' : progress.state === 'success' ? '已完成' : '失败';
            this.element('[data-chapter-progress]').textContent = `${progress.completed}/${progress.total} · ${stateLabel}：${progress.chapter.title}${progress.message ? `（${progress.message}）` : ''}`;
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
            this.element('[data-action="settings"]').addEventListener('click', () => this.openModal('settings'));
            this.element('[data-action="history"]').addEventListener('click', () => this.openHistory());
            this.shadow.querySelectorAll('input[name="cwe-format"], [data-option]').forEach((input) => {
                input.addEventListener('change', () => {
                    this.updateOptionAvailability(input);
                    this.callbacks.onPreferencesChange(this.readExportPreferences());
                });
            });
            this.element('[data-action="save-settings"]').addEventListener('click', () => this.saveSettings());
            this.element('[data-action="chapter-all"]').addEventListener('click', () => this.setAllChapters(true));
            this.element('[data-action="chapter-none"]').addEventListener('click', () => this.setAllChapters(false));
            this.element('[data-action="start-chapters"]').addEventListener('click', () => {
                const indexes = [...this.shadow.querySelectorAll('[data-chapter-list] input:checked')]
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
                const item = event.target instanceof Element ? event.target.closest('[data-history-id]') : null;
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
                this.option('includeAnalysis').checked = false;
            }
            const bankEnabled = bank.checked;
            if (bankEnabled) {
                this.option('withAnswers').checked = false;
                this.option('withWrong').checked = false;
                this.option('shuffle').checked = false;
                this.option('includeAnalysis').checked = false;
            }
            // 无参考答案时强制取消勾选（覆盖历史记录恢复的偏好）
            if (!this.answerOptionsAvailable) {
                this.option('withAnswers').checked = false;
                this.option('withWrong').checked = false;
            }
            this.option('withAnswers').disabled = bankEnabled;
            this.option('withWrong').disabled = bankEnabled;
            this.option('shuffle').disabled = bankEnabled;
            this.option('includeAnalysis').disabled =
                bankEnabled || !(this.option('withAnswers').checked || this.option('withWrong').checked);
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
            this.option('withAnswers').closest('.cwe-check')?.classList.toggle('cwe-hidden', !available);
            this.option('withWrong').closest('.cwe-check')?.classList.toggle('cwe-hidden', !available);
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
                return this.element(`[data-action="${action}"]`).dataset.hasResult !== 'true';
            }
            return false;
        }
        saveSettings() {
            if (!this.settings)
                return;
            const rememberPanelPosition = this.element('[data-setting="rememberPanelPosition"]').checked;
            const settings = {
                ...this.settings,
                theme: this.element('[data-setting="theme"]').value,
                shortcut: (0, shortcut_2.parseShortcut)(this.element('[data-setting="shortcut"]').value, this.settings.shortcut),
                hideShortcut: (0, shortcut_2.parseShortcut)(this.element('[data-setting="hideShortcut"]').value, this.settings.hideShortcut),
                enableDrag: this.element('[data-setting="enableDrag"]').checked,
                rememberPanelPosition,
                autoExtractOnLoad: this.element('[data-setting="autoExtractOnLoad"]').checked,
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
define("application/app-controller", ["require", "exports", "infrastructure/clipboard-service", "infrastructure/download-service", "infrastructure/history-repository", "infrastructure/logger", "infrastructure/settings-repository", "utils/async", "utils/hash", "utils/dom", "utils/shortcut", "utils/text"], function (require, exports, clipboard_service_1, download_service_1, history_repository_1, logger_2, settings_repository_1, async_4, hash_6, dom_6, shortcut_3, text_8) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.AppController = void 0;
    class AppController {
        constructor(view, extractionService, chapterService, exportService, frameBridge, settingsRepository = new settings_repository_1.SettingsRepository(), historyRepository = new history_repository_1.HistoryRepository(), downloadService = new download_service_1.DownloadService(), clipboardService = new clipboard_service_1.ClipboardService()) {
            this.view = view;
            this.extractionService = extractionService;
            this.chapterService = chapterService;
            this.exportService = exportService;
            this.frameBridge = frameBridge;
            this.settingsRepository = settingsRepository;
            this.historyRepository = historyRepository;
            this.downloadService = downloadService;
            this.clipboardService = clipboardService;
            this.currentResult = null;
            this.observer = null;
            this.logger = new logger_2.Logger();
            this.refreshChaptersDebounced = (0, async_4.debounce)(() => this.refreshChapterAvailability(), 700);
            this.onShortcut = (event) => {
                if ((0, dom_6.isEditableTarget)(event.target))
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
                const requestedAt = Date.now();
                this.frameBridge.requestRefresh();
                let result = this.extractionService.extractAccessibleDocuments();
                if (!result) {
                    result = await this.extractionService.waitForChangedResult({
                        afterTimestamp: requestedAt,
                        timeoutMs: 4500,
                        intervalMs: 300,
                    });
                }
                if (!result || result.questions.length === 0) {
                    throw new Error('当前页面未识别到题目，请确认题目已经加载完成');
                }
                this.acceptResult(result);
            }
            catch (error) {
                this.logger.warn('Extraction failed', error);
                this.view.setStatus((0, text_8.humanizeError)(error), 'error');
            }
            finally {
                this.view.setBusy(null);
                this.refreshChapterAvailability();
            }
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
                this.view.setStatus((0, text_8.humanizeError)(error), 'error');
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
            this.view.applyExportOptions({ ...options, filename: (0, text_8.sanitizeFilename)(result.title) });
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
                let message;
                if (options.format === 'word') {
                    const warnMsg = failedImages > 0 ? `（${failedImages} 张图片加载失败）` : '';
                    const withAnswers = options.bankImport || options.withAnswers;
                    const withWrong = !options.bankImport && options.withWrong;
                    message =
                        (options.bankImport
                            ? '题库导入格式已下载'
                            : 'Word 试卷' + (withAnswers ? '（含答案）' : '') + (withWrong ? '（含错题）' : '') + '已下载') +
                            warnMsg;
                }
                else if (artifacts.length > 1) {
                    message = `已下载 ${artifacts.length} 个章节文件`;
                }
                else {
                    message = '文件已下载';
                }
                this.view.setStatus(message, failedImages > 0 ? 'warning' : 'success');
            }
            catch (error) {
                this.logger.error('Export failed', error);
                const prefix = options.format === 'word' ? 'Word 导出失败: ' : '导出失败：';
                this.view.setStatus(`${prefix}${(0, text_8.humanizeError)(error)}`, 'error');
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
                this.view.setStatus(`复制失败：${(0, text_8.humanizeError)(error)}`, 'error');
            }
            finally {
                this.view.setBusy(null);
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
                this.view.setStatus(`重新下载失败：${(0, text_8.humanizeError)(error)}`, 'error');
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
define("main", ["require", "exports", "application/app-controller", "application/chapter-extraction-service", "application/extraction-service", "exporters/export-service", "extractors/composite-extractor", "infrastructure/frame-bridge", "ui/panel-view"], function (require, exports, app_controller_1, chapter_extraction_service_1, extraction_service_1, export_service_1, composite_extractor_3, frame_bridge_2, panel_view_1) {
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
        const extractionService = new extraction_service_1.ExtractionService(new composite_extractor_3.CompositeExtractor(), frameBridge);
        const chapterService = new chapter_extraction_service_1.ChapterExtractionService(extractionService);
        const controller = new app_controller_1.AppController(new panel_view_1.PanelView(), extractionService, chapterService, new export_service_1.ExportService(), frameBridge);
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
