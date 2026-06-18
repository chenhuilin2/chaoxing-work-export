// ==UserScript==
// @name         学习通作业提取器
// @license      GPL-3.0
// @version      1.4
// @description  一键提取学习通作业题目，支持 Word/TXT/MD 导出，答案/错题收集，题目随机打乱，暗色模式，快捷键
// @author       huilin
// @icon         http://pan-yz.chaoxing.com/favicon.ico
// @match        *://*.chaoxing.com/*
// @match        *://*.edu.cn/*
// @require      https://unpkg.com/docx@8.5.0/build/index.umd.js
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // ==================== 样式注入 ====================
  const css = `
/* ===== 触发按钮 ===== */
#xxt-panel-btn {
  position: fixed; top: 120px; right: 0; z-index: 99999;
  width: 38px; padding: 14px 7px; background: #1e88e5;
  color: #fff; border: none; border-radius: 8px 0 0 8px;
  font-size: 13px; font-weight: 600; cursor: pointer;
  writing-mode: vertical-rl; letter-spacing: 3px;
  box-shadow: -3px 3px 12px rgba(30,136,229,0.25);
  transition: 0.25s cubic-bezier(.4,0,.2,1);
  font-family: "Microsoft YaHei","微软雅黑",sans-serif;
}
#xxt-panel-btn:hover {
  background: #1565c0; width: 42px;
  box-shadow: -4px 4px 16px rgba(30,136,229,0.35);
}

/* ===== 面板主体 ===== */
#xxt-panel {
  position: fixed; top: 80px; right: -400px; z-index: 99998;
  width: 380px; background: #fff;
  border-radius: 12px 0 0 12px;
  box-shadow: -4px 4px 20px rgba(0,0,0,0.08), -2px 2px 6px rgba(0,0,0,0.04);
  padding: 20px 18px 18px;
  font-family: "PingFang SC","Microsoft YaHei","微软雅黑",-apple-system,sans-serif;
  font-size: 13px; color: #333;
  transition: right 0.35s ease;
  max-height: 86vh; overflow-y: auto;
}
#xxt-panel::-webkit-scrollbar { width: 4px; }
#xxt-panel::-webkit-scrollbar-thumb { background: #d0d7de; border-radius: 4px; }
#xxt-panel.open { right: 0; }

/* ===== 标题栏 ===== */
#xxt-panel .xxt-header {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 14px; padding-bottom: 12px;
  border-bottom: 1px solid #eee;
}
#xxt-panel .xxt-header h3 {
  font-size: 15px; font-weight: 700; color: #222;
  margin: 0;
}
#xxt-panel .xxt-close-btn {
  width: 24px; height: 24px; border: none; background: #f0f0f0;
  border-radius: 50%; font-size: 15px; color: #999;
  cursor: pointer; display: flex; align-items: center; justify-content: center;
  transition: all 0.2s ease; line-height: 1;
}
#xxt-panel .xxt-close-btn:hover { background: #e0e0e0; color: #555; }

/* ===== 提取按钮 ===== */
#xxt-panel .xxt-btn-extract {
  display: block; width: 100%; padding: 11px; border: none;
  border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer;
  margin-bottom: 10px; transition: all 0.2s ease;
  background: #1e88e5; color: #fff;
  box-shadow: 0 2px 6px rgba(30,136,229,0.2);
  letter-spacing: 1px;
}
#xxt-panel .xxt-btn-extract:hover { background: #1565c0; box-shadow: 0 3px 10px rgba(30,136,229,0.28); }
#xxt-panel .xxt-btn-extract:active { transform: scale(0.985); }
#xxt-panel .xxt-btn-extract:disabled {
  background: #ccc; box-shadow: none; cursor: not-allowed; transform: none;
}

/* ===== 状态提示 ===== */
#xxt-panel .xxt-status {
  text-align: center; padding: 9px 12px; font-size: 12.5px;
  margin-bottom: 10px; border-radius: 6px; font-weight: 500;
}
.xxt-status-ok { color: #16a34a; background: #f0fdf4; border: 1px solid #bbf7d0; }
.xxt-status-err { color: #dc2626; background: #fef2f2; border: 1px solid #fecaca; }
.xxt-status-warn { color: #d97706; background: #fffbeb; border: 1px solid #fde68a; }

/* ===== 统计卡片 ===== */
#xxt-panel .xxt-stat {
  display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px;
  margin: 12px 0 14px;
}
#xxt-panel .xxt-stat-item {
  text-align: center; padding: 10px 4px 8px;
  background: #fafbfc; border-radius: 8px;
  border: 1px solid #eef0f2;
  transition: border-color 0.2s;
}
#xxt-panel .xxt-stat-item:hover { border-color: #1e88e5; }
#xxt-panel .xxt-stat-item .xxt-num {
  font-size: 22px; font-weight: 800; color: #1e88e5; line-height: 1.2;
}
#xxt-panel .xxt-stat-item .xxt-label {
  font-size: 11px; color: #999; margin-top: 2px; font-weight: 500;
}

/* ===== 分隔区块 ===== */
#xxt-panel .xxt-section {
  background: #fafbfc; border-radius: 8px; padding: 12px 14px;
  border: 1px solid #eef0f2; margin-top: 10px;
}

/* ===== 文件名输入 ===== */
#xxt-panel .xxt-filename {
  width: 100%; padding: 9px 12px; border: 1.5px solid #ddd;
  border-radius: 6px; font-size: 12.5px; color: #333;
  box-sizing: border-box; outline: none; transition: all 0.2s ease;
  background: #fff;
}
#xxt-panel .xxt-filename:focus {
  border-color: #1e88e5; box-shadow: 0 0 0 3px rgba(30,136,229,0.08);
}

/* ===== 格式选择 ===== */
#xxt-panel .xxt-format-row {
  display: flex; align-items: center; gap: 8px; margin-top: 10px;
  font-size: 12.5px; color: #666; font-weight: 500;
}
#xxt-panel .xxt-format-row label {
  cursor: pointer; display: inline-flex; align-items: center; gap: 5px;
  padding: 5px 14px; border-radius: 18px; border: 1.5px solid #ddd;
  background: #fff; transition: all 0.2s ease; font-size: 12px;
}
#xxt-panel .xxt-format-row label:hover { border-color: #bbb; background: #f5f5f5; }
#xxt-panel .xxt-format-row input[type="radio"] {
  width: 14px; height: 14px; cursor: pointer; accent-color: #1e88e5;
}
#xxt-panel .xxt-format-row label:has(input:checked) {
  color: #1e88e5; border-color: #1e88e5; background: #e8f4fd;
}

/* ===== 操作按钮组 ===== */
#xxt-panel .xxt-actions {
  display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 12px;
}
#xxt-panel .xxt-actions .xxt-btn {
  padding: 10px; border-radius: 8px; font-size: 13px; font-weight: 600;
  margin-bottom: 0; transition: all 0.2s ease;
}
.xxt-btn-outline {
  background: #fff; color: #1e88e5; border: 1.5px solid #1e88e5 !important;
}
.xxt-btn-outline:hover {
  background: #1e88e5; color: #fff; border-color: transparent !important;
}

/* ===== 开关选项（复选框） ===== */
#xxt-panel .xxt-toggle {
  display: flex; align-items: flex-start; gap: 10px; margin-top: 12px;
  padding: 10px 12px; border-radius: 8px;
  background: #fafbfc; border: 1px solid #eef0f2;
  cursor: pointer; transition: all 0.2s ease;
}
#xxt-panel .xxt-toggle:hover { background: #f0f2f5; }
#xxt-panel .xxt-toggle .xxt-checkbox-wrap {
  flex-shrink: 0; width: 18px; height: 18px; margin-top: 1px;
  border: 2px solid #ccc; border-radius: 4px;
  display: flex; align-items: center; justify-content: center;
  transition: all 0.2s ease; background: #fff;
}
#xxt-panel .xxt-toggle input { display: none; }
#xxt-panel .xxt-toggle input:checked + .xxt-checkbox-wrap {
  background: #1e88e5; border-color: #1e88e5;
}
#xxt-panel .xxt-toggle input:checked + .xxt-checkbox-wrap::after {
  content: ''; width: 5px; height: 9px;
  border: solid #fff; border-width: 0 2px 2px 0;
  transform: rotate(45deg) translateY(-1px);
}
#xxt-panel .xxt-toggle span {
  font-size: 12.5px; color: #444; line-height: 1.5; font-weight: 500;
}

/* ===== 错题提示 ===== */
#xxt-panel .xxt-wrong-hint {
  font-size: 11.5px; color: #dc2626; margin-top: 8px;
  padding: 8px 12px; background: #fef2f2; border-radius: 6px;
  border: 1px solid #fecaca; font-weight: 500;
  animation: xxt-shake 0.4s ease-in-out;
}
@keyframes xxt-shake {
  0%,100% { transform: translateX(0); }
  25% { transform: translateX(-3px); }
  50% { transform: translateX(3px); }
  75% { transform: translateX(-2px); }
}

.xxt-hidden { display: none !important; }

/* ===== 设置面板 ===== */
#xxt-panel .xxt-settings {
  border-top: 1px solid #eee; margin-top: 14px; padding-top: 12px;
}
#xxt-panel .xxt-settings-title {
  font-size: 12px; color: #999; font-weight: 600;
  margin-bottom: 8px; letter-spacing: 1px;
}
#xxt-panel .xxt-setting-row {
  display: flex; align-items: center; justify-content: space-between;
  padding: 8px 12px; border-radius: 8px;
  background: #fafbfc; border: 1px solid #eef0f2;
  margin-bottom: 8px; transition: background 0.2s;
  font-size: 12.5px; color: #444;
}
#xxt-panel .xxt-setting-row:hover { background: #f0f2f5; }
#xxt-panel .xxt-setting-label {
  font-weight: 500; white-space: nowrap;
}
#xxt-panel .xxt-theme-group {
  display: flex; gap: 4px;
}
#xxt-panel .xxt-theme-btn {
  padding: 3px 10px; border-radius: 14px; border: 1.5px solid #ddd;
  background: #fff; color: #666; font-size: 11px; cursor: pointer;
  transition: all 0.2s ease; font-weight: 500;
}
#xxt-panel .xxt-theme-btn:hover { border-color: #bbb; color: #333; }
#xxt-panel .xxt-theme-btn.active {
  background: #e8f4fd; color: #1e88e5; border-color: #1e88e5;
}
#xxt-panel .xxt-shortcut-display {
  display: flex; align-items: center; gap: 6px;
  padding: 4px 10px; border-radius: 6px;
  background: #f0f0f0; border: 1px solid #ddd;
  cursor: pointer; transition: all 0.2s ease; font-size: 12px;
}
#xxt-panel .xxt-shortcut-display:hover { border-color: #1e88e5; }
#xxt-panel .xxt-shortcut-keys {
  color: #1e88e5; font-weight: 600; font-family: monospace;
  padding: 2px 6px; background: #e8f4fd; border-radius: 4px;
}
#xxt-panel .xxt-shortcut-hint {
  color: #999; font-size: 11px;
}

/* ===== 响应式微调 ===== */
@media screen and (max-height: 700px) {
  #xxt-panel { top: 40px; max-height: 92vh; }
  #xxt-panel .xxt-stat { gap: 4px; }
  #xxt-panel .xxt-stat-item { padding: 6px 2px 4px; }
  #xxt-panel .xxt-stat-item .xxt-num { font-size: 18px; }
}
  `;
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  // ==================== 暗色模式 CSS ====================
  const darkCSS = `
[data-xxt-theme="dark"] #xxt-panel-btn {
  background: #3b82f6;
  box-shadow: -3px 3px 12px rgba(59,130,246,0.3);
}
[data-xxt-theme="dark"] #xxt-panel-btn:hover {
  background: #2563eb;
  box-shadow: -4px 4px 16px rgba(59,130,246,0.4);
}
[data-xxt-theme="dark"] #xxt-panel {
  background: #1e1e2e; color: #cdd6f4;
  box-shadow: -4px 4px 20px rgba(0,0,0,0.3), -2px 2px 6px rgba(0,0,0,0.2);
}
[data-xxt-theme="dark"] #xxt-panel::-webkit-scrollbar-thumb { background: #45475a; }
[data-xxt-theme="dark"] #xxt-panel .xxt-header {
  border-bottom-color: #313244;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-header h3 { color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-panel .xxt-close-btn {
  background: #313244; color: #a6adc8;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-close-btn:hover { background: #45475a; color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-panel .xxt-btn-extract {
  background: #3b82f6; color: #fff;
  box-shadow: 0 2px 6px rgba(59,130,246,0.25);
}
[data-xxt-theme="dark"] #xxt-panel .xxt-btn-extract:hover {
  background: #2563eb;
  box-shadow: 0 3px 10px rgba(59,130,246,0.35);
}
[data-xxt-theme="dark"] #xxt-panel .xxt-btn-extract:disabled { background: #45475a; }
[data-xxt-theme="dark"] .xxt-status-ok { color: #4ade80; background: #052e16; border-color: #166534; }
[data-xxt-theme="dark"] .xxt-status-err { color: #f87171; background: #450a0a; border-color: #991b1b; }
[data-xxt-theme="dark"] .xxt-status-warn { color: #fbbf24; background: #451a03; border-color: #92400e; }
[data-xxt-theme="dark"] #xxt-panel .xxt-stat-item {
  background: #181825; border-color: #313244;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-stat-item:hover { border-color: #3b82f6; }
[data-xxt-theme="dark"] #xxt-panel .xxt-stat-item .xxt-num { color: #89b4fa; }
[data-xxt-theme="dark"] #xxt-panel .xxt-stat-item .xxt-label { color: #a6adc8; }
[data-xxt-theme="dark"] #xxt-panel .xxt-section {
  background: #181825; border-color: #313244;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-filename {
  border-color: #45475a; color: #cdd6f4; background: #11111b;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-filename:focus {
  border-color: #3b82f6;
  box-shadow: 0 0 0 3px rgba(59,130,246,0.12);
}
[data-xxt-theme="dark"] #xxt-panel .xxt-format-row { color: #a6adc8; }
[data-xxt-theme="dark"] #xxt-panel .xxt-format-row label {
  border-color: #45475a; background: #181825;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-format-row label:hover { border-color: #585b70; background: #1e1e2e; }
[data-xxt-theme="dark"] #xxt-panel .xxt-format-row label:has(input:checked) {
  color: #89b4fa; border-color: #3b82f6; background: #1e1e3e;
}
[data-xxt-theme="dark"] .xxt-btn-outline {
  background: #181825; color: #89b4fa; border-color: #3b82f6 !important;
}
[data-xxt-theme="dark"] .xxt-btn-outline:hover {
  background: #3b82f6; color: #fff;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-toggle {
  background: #181825; border-color: #313244;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-toggle:hover { background: #1e1e2e; }
[data-xxt-theme="dark"] #xxt-panel .xxt-toggle .xxt-checkbox-wrap {
  border-color: #585b70; background: #11111b;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-toggle span { color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-panel .xxt-wrong-hint {
  color: #f87171; background: #450a0a; border-color: #991b1b;
}
/* ===== 设置面板暗色 ===== */
[data-xxt-theme="dark"] #xxt-panel .xxt-settings {
  border-top-color: #313244;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-setting-row {
  background: #181825; border-color: #313244;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-setting-row:hover { background: #1e1e2e; }
[data-xxt-theme="dark"] #xxt-panel .xxt-theme-btn {
  background: #313244; color: #a6adc8; border-color: #45475a;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-theme-btn:hover { border-color: #585b70; color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-panel .xxt-theme-btn.active {
  background: #1e3a5f; color: #89b4fa; border-color: #3b82f6;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-shortcut-display {
  background: #11111b; border-color: #45475a; color: #cdd6f4;
}
[data-xxt-theme="dark"] #xxt-panel .xxt-shortcut-keys {
  color: #89b4fa; background: #1e1e3e;
}
  `;
  const darkStyle = document.createElement('style');
  darkStyle.textContent = darkCSS;
  document.head.appendChild(darkStyle);

  // ==================== 设置存储 ====================
  const SETTINGS_KEY = 'xxt_settings';
  const DEFAULT_SETTINGS = {
    theme: 'auto',           // 'auto' | 'light' | 'dark'
    shortcut: {
      ctrl: true, shift: true, alt: false, key: 'e'
    }
  };

  function loadSettings() {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        return {
          theme: saved.theme || DEFAULT_SETTINGS.theme,
          shortcut: { ...DEFAULT_SETTINGS.shortcut, ...(saved.shortcut || {}) }
        };
      }
    } catch (e) { /* ignore */ }
    return { ...DEFAULT_SETTINGS, shortcut: { ...DEFAULT_SETTINGS.shortcut } };
  }

  function saveSettings(settings) {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) { /* ignore */ }
  }

  let currentSettings = loadSettings();

  // ==================== 主题切换 ====================
  function applyTheme(theme) {
    const root = document.documentElement;
    if (theme === 'dark') {
      root.setAttribute('data-xxt-theme', 'dark');
    } else if (theme === 'light') {
      root.removeAttribute('data-xxt-theme');
    } else {
      // auto: 跟随系统
      if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
        root.setAttribute('data-xxt-theme', 'dark');
      } else {
        root.removeAttribute('data-xxt-theme');
      }
    }
  }

  applyTheme(currentSettings.theme);

  // 监听系统主题变化（仅在 auto 模式下生效）
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
    if (currentSettings.theme === 'auto') {
      applyTheme('auto');
    }
  });

  // ==================== 快捷键 ====================
  function formatShortcutLabel(sc) {
    const parts = [];
    if (sc.ctrl) parts.push('Ctrl');
    if (sc.shift) parts.push('Shift');
    if (sc.alt) parts.push('Alt');
    parts.push(sc.key.toUpperCase());
    return parts.join(' + ');
  }

  function isShortcutMatch(e, sc) {
    return e.ctrlKey === sc.ctrl &&
           e.shiftKey === sc.shift &&
           e.altKey === sc.alt &&
           e.key.toLowerCase() === sc.key.toLowerCase();
  }

  document.addEventListener('keydown', (e) => {
    // 如果正在录制快捷键，忽略
    if (window.__xxt_recording) return;
    if (!isShortcutMatch(e, currentSettings.shortcut)) return;

    e.preventDefault();
    const panel = document.getElementById('xxt-panel');
    const btnEl = document.getElementById('xxt-panel-btn');
    const btnExtract = document.getElementById('xxt-btnExtract');

    if (!panel) return;

    // 面板未打开：打开面板
    if (!panel.classList.contains('open')) {
      panel.classList.add('open');
    }

    // 触发提取
    if (btnExtract && !btnExtract.disabled) {
      btnExtract.click();
    }
  });

  // ==================== 提取逻辑 ====================
  const TYPE_MAP = {
    '单选题': '单选', '选择题': '单选', '多选题': '多选',
    '填空题': '填空', '判断题': '判断', '简答题': '简答',
    '论述题': '简答', '问答题': '简答',
  };

  function detectTypeFromText(text) {
    for (const [label, type] of Object.entries(TYPE_MAP)) {
      if (text.includes(label)) return type;
    }
    return null;
  }

  function cleanHtml(html) {
    const div = document.createElement('div');
    div.innerHTML = html;
    return div.textContent || div.innerText || '';
  }

  function extractCorrectAnswer(qLi) {
    const markAnswer = qLi.querySelector('.mark_answer');
    if (!markAnswer) return '';

    const markKey = markAnswer.querySelector('.mark_key');
    if (markKey) {
      const rightEl = markKey.querySelector('.rightAnswerContent');
      if (rightEl) return rightEl.textContent.trim();
      const rightEl2 = markKey.querySelector('.colorGreen .rightAnswerContent');
      if (rightEl2) return rightEl2.textContent.trim();
    }

    const greenFill = markAnswer.querySelector('.mark_fill.colorGreen');
    if (greenFill) {
      const dds = greenFill.querySelectorAll('dd.rightAnswerContent');
      if (dds.length > 0) {
        return Array.from(dds).map(dd => dd.textContent.trim()).join('；');
      }
    }
    return '';
  }

  function extractMyAnswer(qLi, qtype) {
    const markAnswer = qLi.querySelector('.mark_answer');
    if (!markAnswer) return '';

    // 单选/多选/判断：从 .mark_key 中提取
    const markKey = markAnswer.querySelector('.mark_key');
    if (markKey) {
      const mySpan = markKey.querySelector('.colorDeep .stuAnswerContent');
      if (mySpan) return mySpan.textContent.trim();
    }

    // 填空：从 .mark_fill.colorDeep 中提取
    const deepFill = markAnswer.querySelector('.mark_fill.colorDeep');
    if (deepFill) {
      const spans = deepFill.querySelectorAll('.stuAnswerContent');
      if (spans.length > 0) {
        return Array.from(spans).map(s => s.textContent.trim()).join('；');
      }
    }
    return '';
  }

  function isAnswerWrong(qLi, myAnswer, correctAnswer) {
    if (!myAnswer) return false; // 没有我的答案，不算错题
    if (!correctAnswer) return false;
    return myAnswer.trim() !== correctAnswer.trim();
  }

  // Fisher-Yates 洗牌算法，同类型题目内部打乱
  function shuffleQuestions(results, typeOrder) {
    const shuffled = {};
    for (const qtype of typeOrder) {
      const arr = [...(results[qtype] || [])];
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [arr[i], arr[j]] = [arr[j], arr[i]];
      }
      shuffled[qtype] = arr;
    }
    return shuffled;
  }

  function extract() {
    const results = { '单选': [], '多选': [], '填空': [], '判断': [], '简答': [] };
    const typeOrder = [];
    let wrongCount = 0;

    const markItems = document.querySelectorAll('.mark_item');
    if (markItems.length === 0) return { results, typeOrder, wrongCount, hasMyAnswer: false };

    let hasAnyMyAnswer = false;

    markItems.forEach(markItem => {
      const typeTit = markItem.querySelector('.type_tit');
      if (!typeTit) return;
      const sectionType = detectTypeFromText(typeTit.textContent || '');
      if (!sectionType) return;

      if (!typeOrder.includes(sectionType)) typeOrder.push(sectionType);

      const questionLis = markItem.querySelectorAll('.questionLi');
      questionLis.forEach(qLi => {
        const qtContent = qLi.querySelector('.qtContent');
        if (!qtContent) return;
        let stem = cleanHtml(qtContent.innerHTML).trim();
        if (!stem) return;

        const options = [];
        const markLetter = qLi.querySelector('.mark_letter');
        if (markLetter) {
          markLetter.querySelectorAll('li').forEach(li => {
            const text = li.textContent.trim();
            const match = text.match(/^([A-H])\.?\s*(.+)$/);
            if (match) options.push({ letter: match[1], text: match[2].trim() });
          });
        }

        const correctAnswer = extractCorrectAnswer(qLi);
        const myAnswer = extractMyAnswer(qLi, sectionType);
        if (myAnswer) hasAnyMyAnswer = true;
        const wrong = isAnswerWrong(qLi, myAnswer, correctAnswer);
        if (wrong) wrongCount++;

        results[sectionType].push({ stem, options, correctAnswer, myAnswer, isWrong: wrong });
      });
    });

    return { results, typeOrder, wrongCount, hasMyAnswer: hasAnyMyAnswer };
  }

  // ==================== TXT 格式化 ====================
  function formatOutput(results, typeOrder) {
    const typeLabels = {
      '单选': '单选题', '多选': '多选题', '填空': '填空题',
      '判断': '判断题', '简答': '简答题',
    };
    const typeNumbers = ['一', '二', '三', '四', '五', '六'];

    let output = '';
    let globalNum = 0;
    let sectionIdx = 0;

    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;

      const label = typeLabels[qtype];
      const num = typeNumbers[sectionIdx] || sectionIdx + 1;
      output += `${num}. ${label}（共${questions.length}题）\n`;

      for (const q of questions) {
        globalNum++;
        output += `${globalNum}. (${label})${q.stem}\n`;
        if (q.options && q.options.length > 0) {
          for (const opt of q.options) {
            output += `${opt.letter}. ${opt.text}\n`;
          }
        }
        output += '\n';
      }
      sectionIdx++;
    }
    return output.trim();
  }

  function formatAnswersTXT(results, typeOrder) {
    const typeLabels = { '单选': '单选题', '多选': '多选题', '填空': '填空题', '判断': '判断题', '简答': '简答题' };
    const typeNumbers = ['一', '二', '三', '四', '五', '六'];

    let output = '';
    let globalNum = 0;
    let sectionIdx = 0;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;

      const num = typeNumbers[sectionIdx] || sectionIdx + 1;
      output += `${num}、${typeLabels[qtype]}\n\n`;
      sectionIdx++;

      for (const q of questions) {
        globalNum++;
        const answer = q.correctAnswer || '（未找到答案）';
        if (qtype === '填空' && answer.includes('；')) {
          const parts = answer.split('；').map(p => p.trim().replace(/^\(\d+\)\s*/, ''));
          output += `${globalNum}. \n`;
          parts.forEach((part, i) => { output += `(${i + 1}) ${part}\n`; });
          output += '\n';
        } else {
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

    let wrongNum = 0;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;
      if (qtype === '简答') continue; // 简答题不算错题
      for (const q of questions) {
        if (!q.isWrong) continue;
        wrongNum++;
        const typeLabel = qtype === '填空' ? '填空题' : '题目';
        output += `${wrongNum}. (${typeLabel})${q.stem}\n`;
        output += `   我的答案: ${q.myAnswer || '无'}\n`;
        output += `   正确答案: ${q.correctAnswer || '（未找到答案）'}\n\n`;
      }
    }
    return output.replace(/\n+$/, '');
  }

  // ==================== Markdown 格式化 ====================
  function formatOutputMD(results, typeOrder) {
    const typeLabels = {
      '单选': '单选题', '多选': '多选题', '填空': '填空题',
      '判断': '判断题', '简答': '简答题',
    };
    const typeNumbers = ['一', '二', '三', '四', '五', '六'];

    let output = '';
    let globalNum = 0;
    let sectionIdx = 0;

    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;

      const label = typeLabels[qtype];
      const num = typeNumbers[sectionIdx] || sectionIdx + 1;
      output += `### ${num}、${label}（共${questions.length}题）\n\n`;

      for (const q of questions) {
        globalNum++;
        output += `**${globalNum}.** ${q.stem}\n\n`;
        if (q.options && q.options.length > 0) {
          for (const opt of q.options) {
            output += `- ${opt.letter}. ${opt.text}\n`;
          }
          output += '\n';
        } else {
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
      if (!questions || questions.length === 0) continue;
      const typeLabels = { '单选': '单选题', '多选': '多选题', '填空': '填空题', '判断': '判断题', '简答': '简答题' };
      output += `**${typeLabels[qtype]}**\n\n`;
      for (const q of questions) {
        globalNum++;
        const answer = q.correctAnswer || '（未找到答案）';
        if (qtype === '填空' && answer.includes('；')) {
          const parts = answer.split('；').map(p => p.trim().replace(/^\(\d+\)\s*/, ''));
          output += `${globalNum}.  \n`;
          parts.forEach((part, i) => { output += `    (${i + 1}) ${part}  \n`; });
          output += '\n';
        } else {
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

    let wrongNum = 0;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;
      if (qtype === '简答') continue; // 简答题不算错题
      for (const q of questions) {
        if (!q.isWrong) continue;
        wrongNum++;
        output += `**${wrongNum}.** ${q.stem}\n\n`;
        output += `- 我的答案: ${q.myAnswer || '无'}\n`;
        output += `- 正确答案: ${q.correctAnswer || '（未找到答案）'}\n\n`;
      }
    }
    return output.replace(/\n+$/, '');
  }

  // ==================== Word 文档生成（纯前端） ====================
  function normalizeStem(stem) {
    return stem.replace(/\(\s{3,}\)/g, '(   )');
  }

  async function generateWordBlob(results, typeOrder, title) {
    const { Document, Packer, Paragraph, TextRun,
            AlignmentType, convertMillimetersToTwip,
            TabStopType } = docx;

    const typeHeaders = {
      '单选': '一、单项选择题', '多选': '二、多项选择题',
      '填空': '三、填空题', '判断': '四、判断题',
      '简答': '五、简答题',
    };

    const children = [];

    // 标题
    children.push(new Paragraph({
      children: [new TextRun({ text: title || '试卷', font: "宋体", size: 32, bold: true, color: "000000" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 }
    }));

    let qNum = 0;

    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;

      const header = typeHeaders[qtype] || qtype;
      const count = questions.length;

      children.push(new Paragraph({
        children: [new TextRun({ text: `${header}（本大题共${count}小题）`, font: "宋体", size: 28, bold: true, color: "000000" })],
        spacing: { after: 120 }
      }));

      for (const q of questions) {
        qNum++;

        if (qtype === '单选' || qtype === '多选') {
          children.push(new Paragraph({
            children: [new TextRun({ text: `${qNum}. ${normalizeStem(q.stem)}`, font: "宋体", size: 24 })],
            spacing: { after: 40 }
          }));
          const options = q.options || [];
          if (options.length > 0) {
            // 制表位对齐：每行两个选项，Tab 对齐，无表格无边框
            for (let i = 0; i < options.length; i += 2) {
              const left = options[i];
              const right = options[i + 1];
              let text = `${left.letter}. ${left.text}`;
              if (right) text += `\t${right.letter}. ${right.text}`;
              children.push(new Paragraph({
                children: [new TextRun({ text, font: "宋体", size: 24 })],
                tabStops: [{ type: TabStopType.LEFT, position: 4500 }],
                spacing: { after: 40 }
              }));
            }
            children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
          }
        } else if (qtype === '填空') {
          children.push(new Paragraph({
            children: [new TextRun({ text: `${qNum}. ${normalizeStem(q.stem)}`, font: "宋体", size: 24 })],
            spacing: { after: 160 }
          }));
        } else if (qtype === '判断') {
          children.push(new Paragraph({
            children: [new TextRun({ text: `${qNum}. ${normalizeStem(q.stem)} ( )`, font: "宋体", size: 24 })],
            spacing: { after: 160 }
          }));
        } else if (qtype === '简答') {
          children.push(new Paragraph({
            children: [new TextRun({ text: `${qNum}. ${normalizeStem(q.stem)}`, font: "宋体", size: 24 })],
            spacing: { after: 40 }
          }));
          for (let i = 0; i < 8; i++) {
            children.push(new Paragraph({
              children: [new TextRun({ text: '', font: "宋体", size: 24 })],
              spacing: { after: 40 }
            }));
          }
          children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
        }
      }
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: convertMillimetersToTwip(25),
              bottom: convertMillimetersToTwip(25),
              left: convertMillimetersToTwip(25),
              right: convertMillimetersToTwip(25)
            }
          }
        },
        children
      }]
    });

    return await Packer.toBlob(doc);
  }

  // ==================== UI 创建 ====================
  let extractedData = null;

  function createPanel() {
    if (document.getElementById('xxt-panel-btn')) return;

    if (observer) { observer.disconnect(); observer = null; }

    const btn = document.createElement('button');
    btn.id = 'xxt-panel-btn';
    btn.textContent = '提取题目';
    btn.title = '学习通题目提取器';
    document.body.appendChild(btn);

    const panel = document.createElement('div');
    panel.id = 'xxt-panel';
    panel.innerHTML = `
      <div class="xxt-header">
        <h3>学习通题目提取器</h3>
        <button class="xxt-close-btn" id="xxt-closeBtn">&times;</button>
      </div>
      <button id="xxt-btnExtract" class="xxt-btn xxt-btn-extract">提取本页题目</button>
      <div id="xxt-status" class="xxt-hidden"></div>
      <div id="xxt-result" class="xxt-hidden">
        <div id="xxt-stat" class="xxt-stat"></div>
        <div class="xxt-section">
          <input type="text" id="xxt-filename" class="xxt-filename xxt-hidden" placeholder="文件名（默认自动生成）">
          <div class="xxt-format-row">
            <span>输出格式</span>
            <label><input type="radio" name="xxt-fmt" value="txt" checked> TXT</label>
            <label><input type="radio" name="xxt-fmt" value="md"> MD</label>
            <label><input type="radio" name="xxt-fmt" value="word"> Word 试卷</label>
          </div>
        </div>
        <div class="xxt-actions">
          <button id="xxt-btnDownload" class="xxt-btn xxt-btn-outline">&#8681; 下载</button>
          <button id="xxt-btnCopy" class="xxt-btn xxt-btn-outline">&#128203; 复制文本</button>
        </div>
        <label class="xxt-toggle">
          <input type="checkbox" id="xxt-chkAnswers">
          <div class="xxt-checkbox-wrap"></div>
          <span>附加答案</span>
        </label>
        <label class="xxt-toggle xxt-hidden" id="xxt-wrong-toggle">
          <input type="checkbox" id="xxt-chkWrong">
          <div class="xxt-checkbox-wrap"></div>
          <span>附加错题</span>
        </label>
        <label class="xxt-toggle" id="xxt-shuffle-toggle">
          <input type="checkbox" id="xxt-chkShuffle">
          <div class="xxt-checkbox-wrap"></div>
          <span>打乱题目顺序</span>
        </label>
        <div id="xxt-wrong-hint" class="xxt-wrong-hint xxt-hidden"></div>
        <div class="xxt-settings">
          <div class="xxt-settings-title">设置</div>
          <div class="xxt-setting-row">
            <span class="xxt-setting-label">主题</span>
            <div class="xxt-theme-group" id="xxt-theme-group">
              <button class="xxt-theme-btn" data-theme="auto">自动</button>
              <button class="xxt-theme-btn" data-theme="light">浅色</button>
              <button class="xxt-theme-btn" data-theme="dark">深色</button>
            </div>
          </div>
          <div class="xxt-setting-row">
            <span class="xxt-setting-label">快捷键</span>
            <div class="xxt-shortcut-display" id="xxt-shortcut-display">
              <span class="xxt-shortcut-keys" id="xxt-shortcut-keys"></span>
              <span class="xxt-shortcut-hint" id="xxt-shortcut-hint">点击修改</span>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(panel);

    const $ = (id) => document.getElementById(id);
    const els = {
      btn, panel,
      btnExtract: $('xxt-btnExtract'),
      status: $('xxt-status'),
      result: $('xxt-result'),
      stat: $('xxt-stat'),
      filename: $('xxt-filename'),
      btnDownload: $('xxt-btnDownload'),
      btnCopy: $('xxt-btnCopy'),
      chkAnswers: $('xxt-chkAnswers'),
      chkWrong: $('xxt-chkWrong'),
      wrongToggle: $('xxt-wrong-toggle'),
      wrongHint: $('xxt-wrong-hint'),
      chkShuffle: $('xxt-chkShuffle'),
      closeBtn: $('xxt-closeBtn'),
    };

    btn.addEventListener('click', () => { panel.classList.toggle('open'); });
    els.closeBtn.addEventListener('click', () => { panel.classList.remove('open'); });
    document.addEventListener('click', (e) => {
      if (!panel.contains(e.target) && e.target !== btn) panel.classList.remove('open');
    });

    // ==================== 设置面板事件 ====================
    // 初始化主题按钮状态
    const themeBtns = document.querySelectorAll('#xxt-theme-group .xxt-theme-btn');
    function updateThemeUI() {
      themeBtns.forEach(b => {
        b.classList.toggle('active', b.dataset.theme === currentSettings.theme);
      });
    }
    updateThemeUI();

    themeBtns.forEach(b => {
      b.addEventListener('click', () => {
        currentSettings.theme = b.dataset.theme;
        applyTheme(currentSettings.theme);
        saveSettings(currentSettings);
        updateThemeUI();
      });
    });

    // 初始化快捷键显示
    const shortcutDisplay = document.getElementById('xxt-shortcut-display');
    const shortcutKeysEl = document.getElementById('xxt-shortcut-keys');
    const shortcutHintEl = document.getElementById('xxt-shortcut-hint');
    function updateShortcutUI() {
      shortcutKeysEl.textContent = formatShortcutLabel(currentSettings.shortcut);
    }
    updateShortcutUI();

    // 快捷键录制
    shortcutDisplay.addEventListener('click', () => {
      shortcutHintEl.textContent = '请按下新快捷键...';
      shortcutKeysEl.textContent = '...';
      shortcutDisplay.style.borderColor = '#1e88e5';
      window.__xxt_recording = true;

      function onKeyDown(e) {
        e.preventDefault();
        e.stopPropagation();
        const key = e.key;
        // 忽略单独的修饰键
        if (['Control', 'Shift', 'Alt', 'Meta'].includes(key)) return;

        currentSettings.shortcut = {
          ctrl: e.ctrlKey, shift: e.shiftKey, alt: e.altKey, key: key.toLowerCase()
        };
        saveSettings(currentSettings);
        updateShortcutUI();
        shortcutHintEl.textContent = '点击修改';
        shortcutDisplay.style.borderColor = '';
        window.__xxt_recording = false;
        document.removeEventListener('keydown', onKeyDown, true);
      }
      document.addEventListener('keydown', onKeyDown, true);

      // 3 秒后自动取消录制
      setTimeout(() => {
        if (window.__xxt_recording) {
          window.__xxt_recording = false;
          document.removeEventListener('keydown', onKeyDown, true);
          updateShortcutUI();
          shortcutHintEl.textContent = '点击修改';
          shortcutDisplay.style.borderColor = '';
        }
      }, 5000);
    });

    // 格式切换时更新文件名，Word 格式隐藏答案/错题选项
    panel.addEventListener('change', (e) => {
      if (e.target.name === 'xxt-fmt' && extractedData) {
        updateFilename(els);
        const isWord = getFormat(els) === 'word';
        const chkAnswers = document.getElementById('xxt-chkAnswers');
        const chkWrong = document.getElementById('xxt-chkWrong');
        const wrongToggle = document.getElementById('xxt-wrong-toggle');
        const wrongHint = document.getElementById('xxt-wrong-hint');
        if (chkAnswers) chkAnswers.closest('.xxt-toggle').style.display = isWord ? 'none' : '';
        if (wrongToggle) wrongToggle.style.display = isWord ? 'none' : '';
        if (wrongHint) wrongHint.style.display = isWord ? 'none' : '';
      }
    });

    els.btnExtract.addEventListener('click', () => {
      els.btnExtract.disabled = true;
      els.btnExtract.textContent = '提取中...';
      els.status.className = 'xxt-hidden';
      els.result.classList.add('xxt-hidden');

      const { results, typeOrder, wrongCount, hasMyAnswer } = extract();
      const total = Object.values(results).reduce((s, a) => s + a.length, 0);

      if (total === 0) {
        showStatus(els, '未检测到题目，请确认已打开作业页面', 'warn');
        els.btnExtract.disabled = false;
        els.btnExtract.textContent = '提取本页题目';
        return;
      }

      const titleEl = document.querySelector('.mark_title');
      const title = titleEl ? titleEl.textContent.trim() : '';

      extractedData = {
        total, title, typeOrder, results, wrongCount, hasMyAnswer,
        breakdown: Object.fromEntries(Object.entries(results).map(([k, v]) => [k, v.length])),
        text: formatOutput(results, typeOrder),
        textWithAnswers: formatOutputWithAnswers(results, typeOrder),
        textWrong: formatWrongQuestionsTXT(results, typeOrder),
        textMD: formatOutputMD(results, typeOrder),
        textWithAnswersMD: formatOutputWithAnswersMD(results, typeOrder),
        textWrongMD: formatWrongQuestionsMD(results, typeOrder),
      };

      renderStats(els);
      updateFilename(els);
      els.result.classList.remove('xxt-hidden');

      // 错题按钮逻辑
      if (hasMyAnswer && wrongCount > 0) {
        els.wrongToggle.classList.remove('xxt-hidden');
        els.wrongHint.textContent = `检测到 ${wrongCount} 道错题，可勾选附加到输出末尾`;
        els.wrongHint.classList.remove('xxt-hidden');
        els.chkWrong.checked = false;
      } else {
        els.wrongToggle.classList.add('xxt-hidden');
        els.wrongHint.classList.add('xxt-hidden');
        els.chkWrong.checked = false;
      }

      showStatus(els, `成功提取 ${total} 道题目` + (wrongCount > 0 ? `，含 ${wrongCount} 道错题` : ''), 'ok');
      els.btnExtract.disabled = false;
      els.btnExtract.textContent = '重新提取';
    });

    els.btnDownload.addEventListener('click', async () => {
      if (!extractedData) return;
      const fmt = getFormat(els);

      if (fmt === 'word') {
        // Word 试卷导出
        els.btnDownload.disabled = true;
        els.btnDownload.textContent = '生成中...';
        try {
          const doShuffle = els.chkShuffle && els.chkShuffle.checked;
          const activeResults = doShuffle
            ? shuffleQuestions(extractedData.results, extractedData.typeOrder)
            : extractedData.results;
          const blob = await generateWordBlob(activeResults, extractedData.typeOrder, extractedData.title);
          const filename = (els.filename.value || '学习通试卷').replace(/\.(txt|md|docx)$/, '') + '.docx';
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url; a.download = filename; a.click();
          URL.revokeObjectURL(url);
          showStatus(els, 'Word 试卷已下载', 'ok');
        } catch (err) {
          showStatus(els, 'Word 导出失败: ' + err.message, 'err');
        }
        els.btnDownload.disabled = false;
        els.btnDownload.textContent = '\u{1F847} 下载';
        return;
      }

      const text = getOutputText(els);
      const filename = els.filename.value || '学习通题目.txt';
      const ext = fmt === 'md' ? '.md' : '.txt';
      const mime = fmt === 'md' ? 'text/markdown;charset=utf-8' : 'text/plain;charset=utf-8';
      const blob = new Blob([text], { type: mime });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename.endsWith(ext) ? filename : filename + ext;
      a.click();
      URL.revokeObjectURL(url);
    });

    els.btnCopy.addEventListener('click', async () => {
      if (!extractedData) return;
      const fmt = getFormat(els);
      if (fmt === 'word') {
        showStatus(els, 'Word 格式不支持复制，请使用下载', 'warn');
        return;
      }
      const text = getOutputText(els);
      try {
        await navigator.clipboard.writeText(text);
      } catch (err) {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
      }
      showStatus(els, '已复制到剪贴板', 'ok');
    });
  }

  function getFormat(els) {
    const checked = document.querySelector('input[name="xxt-fmt"]:checked');
    return checked ? checked.value : 'txt';
  }

  function showStatus(els, msg, type) {
    els.status.textContent = msg;
    els.status.className = `xxt-status xxt-status-${type}`;
    els.status.classList.remove('xxt-hidden');
  }

  function getOutputText(els) {
    if (!extractedData) return '';
    const fmt = getFormat(els);
    const withAnswers = els.chkAnswers.checked;
    const withWrong = els.chkWrong && els.chkWrong.checked;
    const doShuffle = els.chkShuffle && els.chkShuffle.checked;

    // 打乱时重新生成题目文本，答案和错题始终用原始顺序
    let base = '';
    if (doShuffle) {
      const shuffled = shuffleQuestions(extractedData.results, extractedData.typeOrder);
      if (fmt === 'md') {
        base = withAnswers ? formatOutputWithAnswersMD(shuffled, extractedData.typeOrder) : formatOutputMD(shuffled, extractedData.typeOrder);
      } else {
        base = withAnswers ? formatOutputWithAnswers(shuffled, extractedData.typeOrder) : formatOutput(shuffled, extractedData.typeOrder);
      }
    } else {
      if (fmt === 'md') {
        base = withAnswers ? extractedData.textWithAnswersMD : extractedData.textMD;
      } else {
        base = withAnswers ? extractedData.textWithAnswers : extractedData.text;
      }
    }

    if (withWrong) {
      const wrong = fmt === 'md' ? extractedData.textWrongMD : extractedData.textWrong;
      return base + wrong;
    }
    return base;
  }

  function updateFilename(els) {
    if (!extractedData) return;
    const title = extractedData.title || '学习通题目';
    const cleanTitle = title.replace(/[\\/:*?"<>|]/g, '_').substring(0, 60);
    const fmt = getFormat(els);
    const ext = fmt === 'md' ? '.md' : fmt === 'word' ? '.docx' : '.txt';
    els.filename.value = cleanTitle + ext;
    els.filename.classList.remove('xxt-hidden');
  }

  function renderStats(els) {
    const typeLabels = { '单选': '单选题', '多选': '多选题', '填空': '填空题', '判断': '判断题', '简答': '简答题' };
    const order = extractedData.typeOrder || Object.keys(typeLabels);
    let html = '';
    for (const key of order) {
      const count = extractedData.breakdown[key] || 0;
      if (count > 0) {
        html += `<div class="xxt-stat-item"><div class="xxt-num">${count}</div><div class="xxt-label">${typeLabels[key]}</div></div>`;
      }
    }
    els.stat.innerHTML = html;
  }

  // ==================== 初始化 ====================
  let initTimer = null;
  let observer = null;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', createPanel);
  } else {
    createPanel();
  }

  observer = new MutationObserver(() => {
    if (initTimer) clearTimeout(initTimer);
    initTimer = setTimeout(createPanel, 500);
  });
  observer.observe(document.body, { childList: true, subtree: true });

})();