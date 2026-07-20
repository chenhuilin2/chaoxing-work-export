// ==UserScript==
// @name         超星学习通作业/考试一键提取导出word文档
// @license      GPL-3.0
// @version      1.10.2
// @description  一键提取学习通作业题目，支持富文本（图文混排），Word/TXT/MD 导出，答案/错题收集，题库导入格式，暗色模式，快捷键，iframe 提取
// @author       huilin
// @icon         http://pan-yz.chaoxing.com/favicon.ico
// @match        *://*.chaoxing.com/*
// @match        *://*.edu.cn/*
// @require      https://unpkg.com/docx@8.5.0/build/index.umd.js
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // ==================== iframe 跨窗口通信 ====================
  // 存储从 iframe 通过 postMessage 发来的题目数据
  window.__xxt_iframe_data = null;
  // 父窗口自动提取回调（由 createPanel 设置）
  window.__xxt_auto_extract = null;

  // 监听来自 iframe 的题目提取结果（跨域 iframe 通过 postMessage 回传）
  window.addEventListener('message', function(e) {
    if (e.data && e.data.type === 'xxt-iframe-result' && e.data.data) {
      window.__xxt_iframe_data = e.data.data;
      // 章节切换时自动触发提取
      if (window.__xxt_auto_extract) window.__xxt_auto_extract();
    }
  });

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
/* 题库导入勾选时禁用打乱/答案选项，不隐藏 */
#xxt-panel .xxt-toggle.xxt-disabled {
  opacity: 0.45; pointer-events: none; user-select: none;
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

/* ===== 设置齿轮图标 ===== */
#xxt-panel .xxt-settings-btn {
  width: 24px; height: 24px; border: none; background: transparent;
  border-radius: 50%; cursor: pointer; display: flex; align-items: center;
  justify-content: center; transition: all 0.2s ease; margin-right: 4px;
  color: #999; padding: 0;
}
#xxt-panel .xxt-settings-btn:hover { background: #f0f0f0; color: #555; }
#xxt-panel .xxt-settings-btn svg {
  width: 16px; height: 16px; fill: currentColor;
}

/* ===== 设置弹窗遮罩 ===== */
#xxt-settings-modal {
  position: fixed; inset: 0; z-index: 100000;
  background: rgba(0,0,0,0.35);
  display: flex; align-items: center; justify-content: center;
  opacity: 0; pointer-events: none;
  transition: opacity 0.2s ease;
}
#xxt-settings-modal.open {
  opacity: 1; pointer-events: auto;
}

/* ===== 设置弹窗主体 ===== */
#xxt-settings-modal .xxt-modal-box {
  background: #fff; border-radius: 14px; width: 340px;
  padding: 24px; box-shadow: 0 8px 32px rgba(0,0,0,0.15);
  transform: translateY(12px); transition: transform 0.25s ease;
}
#xxt-settings-modal.open .xxt-modal-box {
  transform: translateY(0);
}
#xxt-settings-modal .xxt-modal-header {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 18px; padding-bottom: 14px;
  border-bottom: 1px solid #eee;
}
#xxt-settings-modal .xxt-modal-header h3 {
  font-size: 15px; font-weight: 700; color: #222; margin: 0;
}
#xxt-settings-modal .xxt-modal-close {
  width: 24px; height: 24px; border: none; background: #f0f0f0;
  border-radius: 50%; font-size: 15px; color: #999; cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  transition: all 0.2s ease;
}
#xxt-settings-modal .xxt-modal-close:hover { background: #e0e0e0; color: #555; }

/* ===== 设置弹窗内容行 ===== */
#xxt-settings-modal .xxt-setting-row {
  display: flex; align-items: center; justify-content: space-between;
  padding: 10px 12px; border-radius: 8px;
  background: #fafbfc; border: 1px solid #eef0f2;
  margin-bottom: 10px; font-size: 13px; color: #444;
}
#xxt-settings-modal .xxt-setting-label {
  font-weight: 500; white-space: nowrap;
}
#xxt-settings-modal .xxt-theme-group {
  display: flex; gap: 4px;
}
#xxt-settings-modal .xxt-theme-btn {
  padding: 4px 12px; border-radius: 14px; border: 1.5px solid #ddd;
  background: #fff; color: #666; font-size: 11px; cursor: pointer;
  transition: all 0.2s ease; font-weight: 500;
}
#xxt-settings-modal .xxt-theme-btn:hover { border-color: #bbb; color: #333; }
#xxt-settings-modal .xxt-theme-btn.active {
  background: #e8f4fd; color: #1e88e5; border-color: #1e88e5;
}
#xxt-settings-modal .xxt-shortcut-display {
  display: flex; align-items: center; gap: 6px;
  padding: 4px 10px; border-radius: 6px;
  background: #f0f0f0; border: 1px solid #ddd;
  cursor: pointer; transition: all 0.2s ease; font-size: 12px;
}
#xxt-settings-modal .xxt-shortcut-display:hover { border-color: #1e88e5; }
#xxt-settings-modal .xxt-shortcut-keys {
  color: #1e88e5; font-weight: 600; font-family: monospace;
  padding: 2px 6px; background: #e8f4fd; border-radius: 4px;
}
#xxt-settings-modal .xxt-shortcut-hint {
  color: #999; font-size: 11px;
}

/* ===== 历史记录弹窗 ===== */
#xxt-history-modal {
  position: fixed; inset: 0; z-index: 100001;
  background: rgba(0,0,0,0.35);
  display: flex; align-items: center; justify-content: center;
  opacity: 0; pointer-events: none;
  transition: opacity 0.2s ease;
}
#xxt-history-modal.open {
  opacity: 1; pointer-events: auto;
}
#xxt-history-modal .xxt-modal-box {
  background: #fff; border-radius: 14px; width: 400px; max-height: 520px;
  padding: 24px; box-shadow: 0 8px 32px rgba(0,0,0,0.15);
  transform: translateY(12px); transition: transform 0.25s ease;
  display: flex; flex-direction: column;
}
#xxt-history-modal.open .xxt-modal-box {
  transform: translateY(0);
}
#xxt-history-modal .xxt-modal-header {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 14px; padding-bottom: 14px;
  border-bottom: 1px solid #eee; flex-shrink: 0;
}
#xxt-history-modal .xxt-modal-header h3 {
  font-size: 15px; font-weight: 700; color: #222; margin: 0;
}
#xxt-history-modal .xxt-modal-close {
  width: 24px; height: 24px; border: none; background: #f0f0f0;
  border-radius: 50%; font-size: 15px; color: #999; cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  transition: all 0.2s ease;
}
#xxt-history-modal .xxt-modal-close:hover { background: #e0e0e0; color: #555; }
#xxt-history-modal .xxt-history-list {
  flex: 1; overflow-y: auto;
}
#xxt-history-modal .xxt-history-empty {
  text-align: center; color: #bbb; padding: 40px 0; font-size: 13px;
}
#xxt-history-modal .xxt-history-item {
  display: flex; align-items: center; justify-content: space-between;
  padding: 12px 14px; border-radius: 10px;
  background: #fafbfc; border: 1px solid #eef0f2;
  margin-bottom: 8px; transition: background 0.2s;
  cursor: pointer;
}
#xxt-history-modal .xxt-history-item:hover { background: #f0f2f5; }
#xxt-history-modal .xxt-history-info { flex: 1; min-width: 0; }
#xxt-history-modal .xxt-history-title {
  font-size: 13px; font-weight: 600; color: #333;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
#xxt-history-modal .xxt-history-meta {
  font-size: 11px; color: #999; margin-top: 3px;
}
#xxt-history-modal .xxt-history-delete {
  width: 26px; height: 26px; border: none; background: transparent;
  border-radius: 50%; cursor: pointer; color: #ccc; font-size: 16px;
  display: flex; align-items: center; justify-content: center;
  transition: all 0.2s ease; flex-shrink: 0; margin-left: 8px;
}
#xxt-history-modal .xxt-history-delete:hover { background: #fee2e2; color: #ef4444; }

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
/* ===== 设置弹窗暗色 ===== */
[data-xxt-theme="dark"] #xxt-panel .xxt-settings-btn { color: #a6adc8; }
[data-xxt-theme="dark"] #xxt-panel .xxt-settings-btn:hover { background: #313244; color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-modal-box {
  background: #1e1e2e; box-shadow: 0 8px 32px rgba(0,0,0,0.4);
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-modal-header {
  border-bottom-color: #313244;
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-modal-header h3 { color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-modal-close {
  background: #313244; color: #a6adc8;
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-modal-close:hover { background: #45475a; color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-setting-row {
  background: #181825; border-color: #313244; color: #cdd6f4;
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-theme-btn {
  background: #313244; color: #a6adc8; border-color: #45475a;
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-theme-btn:hover { border-color: #585b70; color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-theme-btn.active {
  background: #1e3a5f; color: #89b4fa; border-color: #3b82f6;
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-shortcut-display {
  background: #11111b; border-color: #45475a; color: #cdd6f4;
}
[data-xxt-theme="dark"] #xxt-settings-modal .xxt-shortcut-keys {
  color: #89b4fa; background: #1e1e3e;
}

/* ===== 历史记录弹窗暗色 ===== */
[data-xxt-theme="dark"] #xxt-history-modal .xxt-modal-box {
  background: #1e1e2e; box-shadow: 0 8px 32px rgba(0,0,0,0.4);
}
[data-xxt-theme="dark"] #xxt-history-modal .xxt-modal-header {
  border-bottom-color: #313244;
}
[data-xxt-theme="dark"] #xxt-history-modal .xxt-modal-header h3 { color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-modal-close {
  background: #313244; color: #a6adc8;
}
[data-xxt-theme="dark"] #xxt-history-modal .xxt-modal-close:hover { background: #45475a; color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-empty { color: #585b70; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-item {
  background: #181825; border-color: #313244;
}
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-item:hover { background: #1e1e2e; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-title { color: #cdd6f4; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-meta { color: #a6adc8; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-delete { color: #585b70; }
[data-xxt-theme="dark"] #xxt-history-modal .xxt-history-delete:hover { background: #450a0a; color: #f87171; }
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
    },
    hideShortcut: {
      ctrl: true, shift: true, alt: false, key: 'h'
    }
  };

  function loadSettings() {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        return {
          theme: saved.theme || DEFAULT_SETTINGS.theme,
          shortcut: { ...DEFAULT_SETTINGS.shortcut, ...(saved.shortcut || {}) },
          hideShortcut: { ...DEFAULT_SETTINGS.hideShortcut, ...(saved.hideShortcut || {}) }
        };
      }
    } catch (e) { /* ignore */ }
    return {
      ...DEFAULT_SETTINGS,
      shortcut: { ...DEFAULT_SETTINGS.shortcut },
      hideShortcut: { ...DEFAULT_SETTINGS.hideShortcut }
    };
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

  // ==================== 历史记录 ====================
  const HISTORY_KEY = 'xxt_history';
  const MAX_HISTORY = 10;

  function loadHistory() {
    try {
      const raw = localStorage.getItem(HISTORY_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) { return []; }
  }

  function saveHistory(history) {
    try {
      localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
    } catch (e) { /* ignore */ }
  }

  function addToHistory(extractedData, fmt, withAnswers, withWrong, shuffle, bankImport, outputText) {
    const history = loadHistory();
    let totalQ = 0;
    for (const qtype of extractedData.typeOrder) {
      totalQ += (extractedData.results[qtype] || []).length;
    }
    const entry = {
      id: Date.now(),
      title: extractedData.title || '未命名',
      date: new Date().toLocaleString('zh-CN'),
      format: fmt,
      withAnswers: withAnswers,
      withWrong: withWrong,
      shuffle: shuffle,
      bankImport: bankImport || false,
      totalQuestions: totalQ,
      typeOrder: extractedData.typeOrder,
      results: extractedData.results,
      wrongCount: extractedData.wrongCount,
      hasMyAnswer: extractedData.hasMyAnswer,
      outputText: outputText || ''
    };
    // 去重：相同标题+相同格式的旧记录替换
    const filtered = history.filter(h => !(h.title === entry.title && h.format === entry.format));
    filtered.unshift(entry);
    if (filtered.length > MAX_HISTORY) filtered.length = MAX_HISTORY;
    saveHistory(filtered);
    return filtered;
  }

  function deleteHistoryById(id) {
    const history = loadHistory().filter(h => h.id !== id);
    saveHistory(history);
    return history;
  }

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

  // 一键隐藏/显示浮窗及按钮快捷键
  document.addEventListener('keydown', (e) => {
    if (window.__xxt_recording) return;
    if (!isShortcutMatch(e, currentSettings.hideShortcut)) return;

    e.preventDefault();
    const panel = document.getElementById('xxt-panel');
    const btn = document.getElementById('xxt-panel-btn');
    if (!panel && !btn) return;

    // 面板打开时需先关闭再隐藏
    if (panel) panel.classList.remove('open');
    const hidden = panel ? panel.style.display === 'none' : (btn ? btn.style.display === 'none' : false);
    if (panel) panel.style.display = hidden ? '' : 'none';
    if (btn) btn.style.display = hidden ? '' : 'none';
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

  // ==================== 富文本提取基础设施 ====================
  // 块级标签：遍历子节点后自动追加换行，保证图文混排时段落结构正确
  const RICH_BLOCK_TAGS = new Set(['P', 'DIV', 'LI', 'DD', 'DT', 'TR', 'TABLE', 'SECTION']);

  // 多属性 fallback 解析图片 URL，兼容学习通各种图片加载方式
  function resolveImageUrl(img) {
    if (!img) return '';
    const attrs = ['src', 'data-src', 'data-original', 'origin-src', 'fileid'];
    for (const attr of attrs) {
      const value = img.getAttribute(attr);
      if (value && value.trim()) return value.trim();
    }
    return img.currentSrc || img.src || '';
  }

  // 规范化富文本数组：合并相邻文本节点、去首尾空行、压缩空白
  function normalizeRichContent(parts) {
    const normalized = [];
    const pushText = (text) => {
      if (!text) return;
      const value = text.replace(/\u00a0/g, ' ').replace(/[ \t\r\f]+/g, ' ');
      if (!value) return;
      const last = normalized[normalized.length - 1];
      if (last && last.type === 'text') last.text += value;
      else normalized.push({ type: 'text', text: value });
    };
    const pushBreak = () => {
      const last = normalized[normalized.length - 1];
      if (!last || last.type !== 'break') normalized.push({ type: 'break' });
    };

    for (const part of parts || []) {
      if (!part) continue;
      if (part.type === 'text') {
        pushText(part.text);
      } else if (part.type === 'image' && part.url) {
        normalized.push(part);
      } else if (part.type === 'break') {
        pushBreak();
      }
    }

    while (normalized[0] && normalized[0].type === 'break') normalized.shift();
    while (normalized[normalized.length - 1] && normalized[normalized.length - 1].type === 'break') normalized.pop();
    return normalized;
  }

  // 按 DOM 顺序提取富文本内容：保留文字、图片和换行的原始顺序
  function extractRichContent(el) {
    const parts = [];
    if (!el) return parts;

    const walk = (node) => {
      if (!node) return;
      if (node.nodeType === Node.TEXT_NODE) {
        parts.push({ type: 'text', text: node.nodeValue || '' });
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;

      const tag = node.tagName;
      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT') return;
      if (tag === 'IMG') {
        const url = resolveImageUrl(node);
        if (url) parts.push({ type: 'image', url, alt: node.getAttribute('alt') || '' });
        return;
      }
      if (tag === 'BR') {
        parts.push({ type: 'break' });
        return;
      }

      const beforeLen = parts.length;
      node.childNodes.forEach(walk);
      if (RICH_BLOCK_TAGS.has(tag) && parts.length > beforeLen) parts.push({ type: 'break' });
    };

    el.childNodes.forEach(walk);
    return normalizeRichContent(parts);
  }

  // 将富文本数组转为纯文本，图片通过回调格式化
  function richContentToText(content, imageFormatter) {
    let text = '';
    for (const part of content || []) {
      if (part.type === 'text') text += part.text;
      else if (part.type === 'image') text += imageFormatter ? imageFormatter(part.url, part) : '';
      else if (part.type === 'break') text += '\n';
    }
    return text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  // 从富文本首段剥离选项字母（A~H），返回字母和剩余内容
  function stripOptionPrefix(content) {
    const parts = (content || []).map(part => part.type === 'text' ? { ...part } : part);
    let letter = '';
    for (const part of parts) {
      if (part.type !== 'text') continue;
      if (!part.text.trim()) continue;
      const match = part.text.match(/^\s*([A-H])\s*[\.、．]?\s*/i);
      if (match) {
        letter = match[1].toUpperCase();
        part.text = part.text.slice(match[0].length);
      }
      break;
    }
    return { letter, content: normalizeRichContent(parts) };
  }

  // 剥离题干前缀：题号（如"1."）和题型标签（如"（单选题）"）
  function stripQuestionPrefix(content) {
    const parts = (content || []).map(part => part.type === 'text' ? { ...part } : part);
    let strippedNumber = false;
    let strippedType = false;
    for (const part of parts) {
      if (part.type !== 'text') continue;
      // 去除末尾的分数标记，如 (2.0)
      part.text = part.text.replace(/\s*\(\d+\.\d+\)\s*$/, '');
      if (!strippedNumber) {
        const before = part.text;
        part.text = part.text.replace(/^\s*\d+\.\s*/, '');
        if (part.text !== before || part.text.trim()) strippedNumber = true;
      }
      if (!strippedType) {
        const before = part.text;
        // 剥离已知题型标签（支持圆括号、方头括号、半角括号），避免误伤选项字母
        part.text = part.text.replace(/^\s*[（【(](单选题|多选题|填空题|判断题|简答题|论述题|问答题|选择题)[）】)]\s*/, '');
        if (part.text !== before || part.text.trim()) strippedType = true;
      }
    }
    return normalizeRichContent(parts);
  }

  // 剥离答案标签前缀（如"正确答案："）
  function stripAnswerLabel(content) {
    const parts = (content || []).map(part => part.type === 'text' ? { ...part } : part);
    for (const part of parts) {
      if (part.type !== 'text') continue;
      if (!part.text.trim()) continue;
      part.text = part.text.replace(/^\s*(正确答案|参考答案|答案)\s*[:：]?\s*/, '');
      break;
    }
    return normalizeRichContent(parts);
  }

  // 判断富文本是否有实质内容（文本或图片）
  function hasRichContent(content) {
    return (content || []).some(part => part.type === 'image' || (part.type === 'text' && part.text.trim()));
  }

  // 提取正确答案为富文本数组，保留图片和公式
  function extractCorrectAnswerContent(qLi) {
    const markAnswer = qLi.querySelector('.mark_answer');
    if (!markAnswer) return [];

    const contents = [];
    markAnswer.querySelectorAll('.rightAnswerContent').forEach(el => {
      const content = stripAnswerLabel(extractRichContent(el));
      if (hasRichContent(content)) {
        if (contents.length > 0) contents.push({ type: 'text', text: '；' });
        contents.push(...content);
      }
    });
    if (contents.length > 0) return normalizeRichContent(contents);

    const markKey = markAnswer.querySelector('.mark_key');
    if (markKey) {
      const rightEl = markKey.querySelector('.colorGreen');
      if (rightEl) {
        const content = stripAnswerLabel(extractRichContent(rightEl));
        if (hasRichContent(content)) return content;
      }
    }

    const greenFill = markAnswer.querySelector('.mark_fill.colorGreen');
    if (greenFill) {
      const dds = greenFill.querySelectorAll('dd');
      const fillContents = [];
      dds.forEach(dd => {
        const content = stripAnswerLabel(extractRichContent(dd));
        if (hasRichContent(content)) {
          if (fillContents.length > 0) fillContents.push({ type: 'text', text: '；' });
          fillContents.push(...content);
        }
      });
      if (fillContents.length > 0) return normalizeRichContent(fillContents);
    }

    const fallbackSelectors = ['.mark_key .colorGreen', '.mark_fill.colorGreen'];
    for (const selector of fallbackSelectors) {
      const el = markAnswer.querySelector(selector);
      if (!el) continue;
      const content = stripAnswerLabel(extractRichContent(el));
      if (hasRichContent(content)) {
        return content;
      }
    }
    return [];
  }

  function extractCorrectAnswer(qLi) {
    return richContentToText(extractCorrectAnswerContent(qLi), () => '').replace(/\s+/g, ' ').trim();
  }

  // 适配器：兼容新旧数据格式，统一返回富文本数组
  function questionContent(q) {
    if (hasRichContent(q.stemContent)) return q.stemContent;
    const parts = q.stem ? [{ type: 'text', text: q.stem }] : [];
    if (q.images && q.images.length) {
      q.images.forEach(url => parts.push({ type: 'break' }, { type: 'image', url }));
    }
    return normalizeRichContent(parts);
  }

  function optionContent(opt) {
    if (hasRichContent(opt.content)) return opt.content;
    return opt.text ? [{ type: 'text', text: opt.text }] : [];
  }

  function answerContent(q) {
    if (hasRichContent(q.correctAnswerContent)) return q.correctAnswerContent;
    return q.correctAnswer ? [{ type: 'text', text: q.correctAnswer }] : [];
  }

  // Markdown URL 转义，防止特殊字符破坏图片语法
  function escapeMarkdownAlt(text) {
    return (text || '图片').replace(/[\[\]\n\r]/g, ' ').trim() || '图片';
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

  // 从旧版 li 元素提取选项（富文本）
  function extractOptionFromLegacyLi(li, index) {
    const parsed = stripOptionPrefix(extractRichContent(li));
    const letter = parsed.letter || String.fromCharCode(65 + index);
    const text = richTextOnly(parsed.content);
    if (letter && (text || hasRichContent(parsed.content))) {
      return { letter, text, content: parsed.content };
    }
    return null;
  }

  // 从新版 .answerBg 元素提取选项（富文本）
  function extractOptionFromAnswerBg(bg, index) {
    const numOption = bg.querySelector('.num_option');
    const answerP = bg.querySelector('.answer_p');
    if (numOption && answerP) {
      const letter = (numOption.getAttribute('data') || numOption.textContent.trim() || String.fromCharCode(65 + index)).trim().toUpperCase();
      const content = extractRichContent(answerP);
      const text = richTextOnly(content);
      if (letter && (text || hasRichContent(content))) return { letter, text, content };
    }
    return null;
    return null;
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

  function isAnswerWrong(qLi, myAnswer, correctAnswer, qtype) {
    if (qtype === '简答') return false; // 简答题不纳入错题统计
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
    // 优先从当前文档提取
    const topResult = extractFromRoot(document);
    if (hasQuestions(topResult)) return topResult;

    // 检查 iframe 通过 postMessage 发来的最新数据（跨域 iframe 回传）
    const iframeData = window.__xxt_iframe_data;
    window.__xxt_iframe_data = null; // 每次提取后清空，确保下次拿到最新数据
    if (iframeData && hasQuestions(iframeData)) {
      return iframeData;
    }

    // 学生学习页面：题目在多层嵌套 iframe 中，递归查找（同源 iframe）
    const iframeResult = extractFromIframesRecursive(document);
    if (iframeResult) return iframeResult;

    // 兜底：返回空结果
    return { results: { '单选': [], '多选': [], '填空': [], '判断': [], '简答': [] }, typeOrder: [], wrongCount: 0, hasMyAnswer: false, hasCorrectAnswer: false };
  }

  // 检查提取结果是否包含题目
  function hasQuestions(result) {
    return Object.values(result.results).some(arr => arr.length > 0);
  }

  // 从章节测验/考试页面（.TiMu.newTiMu 结构）提取题目
  function extractFromTiMuRoot(root) {
    const bottom = root.querySelector('#ZyBottom');
    if (!bottom) return null;
    const areas = bottom.querySelectorAll('.aiArea');
    if (areas.length === 0) return null;

    const results = { '单选': [], '多选': [], '填空': [], '判断': [], '简答': [] };
    const typeOrder = [];
    let wrongCount = 0;
    let hasAnyMyAnswer = false;
    let hasCorrectAnswer = false;
    let currentType = null;

    // 提取章节测验页面中我的答案
    function extractMyAnswerTiMu(qDiv) {
      const myAnswerBx = qDiv.querySelector('.newAnswerBx .myAnswerBx');
      if (!myAnswerBx) return '';
      const con = myAnswerBx.querySelector('.myAnswer .answerCon');
      return con ? con.textContent.trim() : '';
    }

    // 提取章节测验页面中正确答案的富文本内容
    function extractCorrectAnswerContentTiMu(qDiv) {
      const correctBx = qDiv.querySelector('.newAnswerBx .correctAnswerBx');
      if (!correctBx) return [];
      const con = correctBx.querySelector('.correctAnswer .answerCon');
      if (!con) return [];
      const content = stripAnswerLabel(extractRichContent(con));
      return hasRichContent(content) ? content : [];
    }

    // 判断章节测验页面中是否答错
    function isAnswerWrongTiMu(qDiv, myAnswer, correctAnswer, qtype) {
      if (qtype === '简答') return false;
      if (!myAnswer) return false;
      // 有对勾图标则正确，有叉图标则错误
      if (qDiv.querySelector('.marking_dui')) return false;
      if (qDiv.querySelector('.marking_cuo')) return true;
      if (!correctAnswer) return false;
      return myAnswer.trim() !== correctAnswer.trim();
    }

    areas.forEach(area => {
      // 题型标题通常在 aiArea 之前，同一层级检测
      const typeTit = area.querySelector('.newTestType');
      if (typeTit) {
        currentType = detectTypeFromText(typeTit.textContent || '');
        if (currentType && !typeOrder.includes(currentType)) typeOrder.push(currentType);
      }

      const qDiv = area.querySelector('.TiMu.newTiMu');
      if (!qDiv) return;

      const titleDiv = qDiv.querySelector('.Zy_TItle');
      if (!titleDiv) return;

      const numEl = titleDiv.querySelector('i.fl');
      const qnum = numEl ? parseInt(numEl.textContent.trim(), 10) || 0 : 0;

      const qtContent = titleDiv.querySelector('.qtContent');
      if (!qtContent) return;
      const stemContent = stripQuestionPrefix(extractRichContent(qtContent));
      if (!hasRichContent(stemContent)) return;
      const stem = formatRichForText(stemContent);

      // 题型识别：优先使用最近遇到的题型标题，其次从题干标签识别
      let sectionType = currentType;
      if (!sectionType) {
        const typeLabel = qtContent.querySelector('.newZy_TItle');
        sectionType = detectTypeFromText(typeLabel ? typeLabel.textContent : '');
      }
      if (!sectionType) return;
      if (!typeOrder.includes(sectionType)) typeOrder.push(sectionType);

      // 选项（单选/多选），判断题通常无选项
      const options = [];
      qDiv.querySelectorAll('.Zy_ulTop.qtDetail > li').forEach((li, i) => {
        const opt = extractOptionFromLegacyLi(li, i);
        if (opt) options.push(opt);
      });

      const myAnswer = extractMyAnswerTiMu(qDiv);
      if (myAnswer) hasAnyMyAnswer = true;

      const correctAnswerContent = extractCorrectAnswerContentTiMu(qDiv);
      const correctAnswer = richContentToText(correctAnswerContent, () => '').replace(/\s+/g, ' ').trim();
      if (correctAnswer) hasCorrectAnswer = true;

      const wrong = isAnswerWrongTiMu(qDiv, myAnswer, correctAnswer, sectionType);
      if (wrong) wrongCount++;

      results[sectionType].push({
        qnum, stem, stemContent, options,
        correctAnswer, correctAnswerContent,
        myAnswer, isWrong: wrong
      });
    });

    if (!hasQuestions({ results })) return null;
    return { results, typeOrder, wrongCount, hasMyAnswer: hasAnyMyAnswer, hasCorrectAnswer };
  }

  // 从指定文档/根节点提取题目（支持 .mark_item、.questionLi、.TiMu.newTiMu 三种结构，富文本版本）
  function extractFromRoot(root) {
    // 优先章节测验/考试页面结构
    const tiMuResult = extractFromTiMuRoot(root);
    if (tiMuResult) return tiMuResult;

    const results = { '单选': [], '多选': [], '填空': [], '判断': [], '简答': [] };
    const typeOrder = [];

    // 优先 .mark_item 结构（带答案的作业详情页）
    const markItems = root.querySelectorAll('.mark_item');
    if (markItems.length > 0) {
      let wrongCount = 0;
      let hasAnyMyAnswer = false;
      let hasCorrectAnswer = false;

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
          // 富文本提取题干
          const stemContent = stripQuestionPrefix(extractRichContent(qtContent));
          if (!hasRichContent(stemContent)) return;
          const stem = formatRichForText(stemContent);

          const options = [];
          const markLetter = qLi.querySelector('.mark_letter');
          if (markLetter) {
            markLetter.querySelectorAll('li').forEach((li, i) => {
              const opt = extractOptionFromLegacyLi(li, i);
              if (opt) options.push(opt);
            });
          }

          const correctAnswerContent = extractCorrectAnswerContent(qLi);
          const correctAnswer = richContentToText(correctAnswerContent, () => '').replace(/\s+/g, ' ').trim();
          const myAnswer = extractMyAnswer(qLi, sectionType);
          if (myAnswer) hasAnyMyAnswer = true;
          if (correctAnswer) hasCorrectAnswer = true;
          const wrong = isAnswerWrong(qLi, myAnswer, correctAnswer, sectionType);
          if (wrong) wrongCount++;

          results[sectionType].push({
            stem, stemContent,
            options, correctAnswer, correctAnswerContent,
            myAnswer, isWrong: wrong
          });
        });
      });

      return { results, typeOrder, wrongCount, hasMyAnswer: hasAnyMyAnswer, hasCorrectAnswer };
    }

    // 新版页面 .questionLi 结构（学习页面，通常无答案）
    const questionLis = root.querySelectorAll('.questionLi');
    if (questionLis.length > 0) {
      questionLis.forEach(qLi => {
        const typeName = qLi.getAttribute('typeName') || '';
        const sectionType = detectTypeFromText(typeName);
        if (!sectionType) return;

        if (!typeOrder.includes(sectionType)) typeOrder.push(sectionType);

        const qData = extractFromQuestionLi(qLi, sectionType);
        if (qData) results[sectionType].push(qData);
      });

      return { results, typeOrder, wrongCount: 0, hasMyAnswer: false, hasCorrectAnswer: false };
    }

    return { results, typeOrder, wrongCount: 0, hasMyAnswer: false, hasCorrectAnswer: false };
  }

  // 从单个 .questionLi 提取题目（富文本版本）
  function extractFromQuestionLi(qLi, sectionType) {
    const markName = qLi.querySelector('.mark_name');
    if (!markName) return null;
    const stemContent = stripQuestionPrefix(extractRichContent(markName));
    if (!hasRichContent(stemContent)) return null;
    const stem = formatRichForText(stemContent);

    const options = [];
    const answerBgs = qLi.querySelectorAll('.answerBg');
    answerBgs.forEach((bg, i) => {
      const opt = extractOptionFromAnswerBg(bg, i);
      if (opt) options.push(opt);
    });
    options.sort((a, b) => a.letter.localeCompare(b.letter));

    return {
      stem, stemContent, options,
      correctAnswer: '', correctAnswerContent: [],
      myAnswer: '', isWrong: false
    };
  }

  // 递归遍历所有 iframe（含嵌套 iframe）查找题目
  function extractFromIframesRecursive(root) {
    const iframes = root.querySelectorAll('iframe');
    for (const iframe of iframes) {
      let doc;
      try {
        doc = iframe.contentDocument || (iframe.contentWindow && iframe.contentWindow.document);
      } catch (e) { continue; }
      if (!doc) continue;

      const result = extractFromRoot(doc);
      if (hasQuestions(result)) return result;

      const nested = extractFromIframesRecursive(doc);
      if (nested) return nested;
    }
    return null;
  }

  // ==================== TXT 格式化（富文本） ====================
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
        output += `${globalNum}. ${formatRichForText(questionContent(q))}\n`;
        output += `${globalNum}. ${formatRichForText(questionContent(q))}\n`;
        if (q.options && q.options.length > 0) {
          for (const opt of q.options) {
            output += `${opt.letter}. ${formatRichForText(optionContent(opt))}\n`;
            output += `${opt.letter}. ${formatRichForText(optionContent(opt))}\n`;
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
        const answer = formatRichForText(answerContent(q)) || '（未找到答案）';
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

    let globalNum = 0;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;
      if (qtype === '简答') { globalNum += questions.length; continue; }
      for (const q of questions) {
        globalNum++;
        if (!q.isWrong) continue;
        const typeLabel = qtype === '填空' ? '填空题' : '题目';
        output += `${globalNum}. (${typeLabel})${formatRichForText(questionContent(q))}\n`;
        output += `${globalNum}. (${typeLabel})${formatRichForText(questionContent(q))}\n`;
        output += `   我的答案: ${q.myAnswer || '无'}\n`;
        output += `   正确答案: ${formatRichForText(answerContent(q)) || '（未找到答案）'}\n\n`;
      }
    }
    return output.replace(/\n+$/, '');
  }

  // ==================== Markdown 格式化（富文本） ====================
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
        output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
        output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
        if (q.options && q.options.length > 0) {
          for (const opt of q.options) {
            output += `- ${opt.letter}. ${formatRichForMD(optionContent(opt))}\n`;
            output += `- ${opt.letter}. ${formatRichForMD(optionContent(opt))}\n`;
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
        const answer = formatRichForMD(answerContent(q)) || '（未找到答案）';
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

    let globalNum = 0;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;
      if (qtype === '简答') { globalNum += questions.length; continue; }
      for (const q of questions) {
        globalNum++;
        if (!q.isWrong) continue;
        output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
        output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
        output += `- 我的答案: ${q.myAnswer || '无'}\n`;
        output += `- 正确答案: ${formatRichForMD(answerContent(q)) || '（未找到答案）'}\n\n`;
      }
    }
    return output.replace(/\n+$/, '');
  }

  // ==================== Word 文档生成（富文本） ====================
  async function fetchImageAsset(url) {
    if (!url) return null;
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000);
      const resp = await fetch(url, { mode: 'cors', signal: controller.signal });
      clearTimeout(timeoutId);
      if (!resp.ok) return null;
      const blob = await resp.blob();
      if (!blob.type.startsWith('image/')) return null;
      const data = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(blob);
      });
      if (!data) return null;
      const size = await new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.naturalWidth || 300, height: img.naturalHeight || 200 });
        img.onerror = () => resolve({ width: 300, height: 200 });
        img.src = data;
      });
      const typeMap = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp' };
      const type = typeMap[blob.type] || 'png';
      return { data, type, width: size.width, height: size.height };
    } catch (e) {
      return null;
    }
  }

  async function buildImageRun(url) {
    if (!url) return null;
    const base64 = await fetchImageAsset(url);
    if (base64) {
      return new docx.ImageRun({
        data: base64,
        transformation: { width: 300, height: 200 },
        type: 'png'
      });
    }
    return null;
  }

  async function buildRichRuns(content, prefix = '') {
    const { TextRun } = docx;
    const runs = [];
    if (prefix) runs.push(new TextRun({ text: prefix, font: "宋体", size: 24 }));
    for (const part of content || []) {
      if (part.type === 'text') {
        const normalized = part.text.replace(/\n+/g, ' ');
        if (normalized) runs.push(new TextRun({ text: normalized, font: "宋体", size: 24 }));
      } else if (part.type === 'image') {
        if (runs.length > 0) runs.push(new TextRun({ text: ' ', font: "宋体", size: 24 }));
        const imgRun = await buildImageRun(part.url);
        if (imgRun) runs.push(imgRun);
        runs.push(new TextRun({ text: ' ', font: "宋体", size: 24 }));
      } else if (part.type === 'break') {
        runs.push(new TextRun({ text: '\n', break: 1, font: "宋体", size: 24 }));
      }
    }
    return runs.length ? runs : [new TextRun({ text: prefix, font: "宋体", size: 24 })];
  }

  // 将富文本内容构建为带段后间距的 Paragraph 数组（Word 导出用）
  async function buildRichParagraphs(content, prefix = '', spacing = 0) {
    const { Paragraph } = docx;
    const runs = await buildRichRuns(content, prefix);
    return [new Paragraph({
      children: runs,
      spacing: { after: spacing }
    })];
  }

  // 获取题目的规范化富文本题干（Word 导出判断题等场景使用）
  function normalizedQuestionContent(q) {
    return questionContent(q);
  }

  async function generateWordBlob(results, typeOrder, title, withAnswers, withWrong, bankImport) {
    const { Document, Packer, Paragraph, TextRun, ImageRun,
            AlignmentType, convertMillimetersToTwip,
            TabStopType, PageBreak } = docx;

    const typeHeaders = {
      '单选': '一、单项选择题', '多选': '二、多项选择题',
      '填空': '三、填空题', '判断': '四、判断题',
      '简答': '五、简答题',
    };

    const bankTypeLabels = {
      '单选': '【单选题】', '多选': '【多选题】', '填空': '【填空题】',
      '判断': '【判断题】', '简答': '【简答题】',
    };

    // 题库导入格式：生成学习通智能导入兼容的 Word 文档
    if (bankImport) {
      const children = [];
      children.push(new Paragraph({
        children: [new TextRun({ text: title || '题库导入', font: "宋体", size: 32, bold: true, color: "000000" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 }
      }));

      let qNum = 0;

      for (const qtype of typeOrder) {
        const questions = results[qtype];
        if (!questions || questions.length === 0) continue;

        const prefix = bankTypeLabels[qtype];

        for (const q of questions) {
          qNum++;
          // 题干（题号 + 题型标签 + 题干内容），括号归一化
          const stemText = formatRichForText(questionContent(q));
          const stem = prefix + stemText.replace(/\(\s{2,}\)/g, '（ ）').replace(/（\s{2,}）/g, '（ ）');
          children.push(new Paragraph({
            children: [new TextRun({ text: `${qNum}.${stem}`, font: "宋体", size: 24 })],
            spacing: { after: 40 }
          }));
          // 题干中的图片
          const stemContent = questionContent(q);
          for (const part of stemContent) {
            if (part.type === 'image') {
              const imgRun = await buildImageRun(part.url);
              if (imgRun) {
                children.push(new Paragraph({
                  children: [imgRun],
                  spacing: { after: 80 }
                }));
              }
            }
          }

          // 选项
          const options = q.options || [];
          for (const opt of options) {
            const runs = await buildRichRuns(optionContent(opt), `${opt.letter}. `);
            children.push(new Paragraph({
              children: runs,
              spacing: { after: 40 }
            }));
          }

          // 答案
          const answer = formatRichForText(answerContent(q)).trim();
          if (answer) {
            let formattedAnswer = answer;
            if (qtype === '多选') {
              const parts = answer.replace(/\s+/g, '').split('');
              formattedAnswerParts = [{ type: 'text', text: parts.join('，') }];
            } else if (!hasAnswerImage && qtype === '判断') {
              if (/^[√✓Tt]|正确|True|TRUE/.test(answer)) {
                formattedAnswerParts = [{ type: 'text', text: '对' }];
              } else {
                formattedAnswerParts = [{ type: 'text', text: '错' }];
              }
            }
            children.push(...await buildRichParagraphs(formattedAnswerParts, '答案：', 120));
          } else {
            children.push(new Paragraph({
              children: [new TextRun({ text: '', font: "宋体", size: 24 })],
              spacing: { after: 120 }
            }));
          }
        }
      }

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              margin: {
                top: convertMillimetersToTwip(20),
                bottom: convertMillimetersToTwip(20),
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
        const stemContent = questionContent(q);

        if (qtype === '单选' || qtype === '多选') {
          children.push(new Paragraph({
            children: await buildRichRuns(stemContent, `${qNum}. `),
            spacing: { after: 40 }
          }));
          const options = q.options || [];
          if (options.length > 0) {
            const maxLen = Math.max(...options.map(o => (o.text || '').length));
            const useVertical = maxLen > 25;

            if (useVertical) {
              for (const opt of options) {
                children.push(new Paragraph({
                  children: await buildRichRuns(optionContent(opt), `${opt.letter}. `),
                  spacing: { after: 40 }
                }));
              }
            } else {
              for (let i = 0; i < options.length; i += 2) {
                const left = options[i];
                const right = options[i + 1];
                const leftRuns = await buildRichRuns(optionContent(left), `${left.letter}. `);
                if (right) {
                  const rightRuns = await buildRichRuns(optionContent(right), `${right.letter}. `);
                  const tabRun = new TextRun({ text: '\t', font: "宋体", size: 24 });
                  children.push(new Paragraph({
                    children: [...leftRuns, tabRun, ...rightRuns],
                    tabStops: [{ type: TabStopType.LEFT, position: 4500 }],
                    spacing: { after: 40 }
                  }));
                } else {
                  children.push(new Paragraph({
                    children: leftRuns,
                    spacing: { after: 40 }
                  }));
                }
              }
            }
            children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
          }
        } else if (qtype === '填空') {
          children.push(new Paragraph({
            children: await buildRichRuns(stemContent, `${qNum}. `),
            spacing: { after: 40 }
          }));
          children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
        } else if (qtype === '判断') {
          // 输出题干，题干末尾追加判断括号，避免重复输出
          const judgeRuns = await buildRichRuns(stemContent, `${qNum}. `);
          judgeRuns.push(new TextRun({ text: '（  ）', font: "宋体", size: 24 }));
          children.push(new Paragraph({
            children: judgeRuns,
            spacing: { after: 120 }
          }));
        } else if (qtype === '简答') {
          children.push(new Paragraph({
            children: await buildRichRuns(stemContent, `${qNum}. `),
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

    // 答案页
    if (withAnswers) {
      children.push(new Paragraph({
        children: [new PageBreak()],
        spacing: { after: 0 }
      }));
      children.push(new Paragraph({
        children: [new TextRun({ text: '参考答案', font: "宋体", size: 32, bold: true, color: "000000" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 }
      }));
      let aNum = 0;
      for (const qtype of typeOrder) {
        const questions = results[qtype];
        if (!questions || questions.length === 0) continue;
        const header = typeHeaders[qtype] || qtype;
        children.push(new Paragraph({
          children: [new TextRun({ text: header, font: "宋体", size: 28, bold: true, color: "000000" })],
          spacing: { after: 120 }
        }));
        for (const q of questions) {
          aNum++;
          const answerRich = answerContent(q);
          const answerRuns = await buildRichRuns(answerRich);
          if (answerRuns.length === 0) {
            answerRuns.push(new TextRun({ text: '（未找到答案）', font: "宋体", size: 24 }));
          }
          children.push(new Paragraph({
            children: [new TextRun({ text: `${aNum}. `, font: "宋体", size: 24 }), ...answerRuns],
            spacing: { after: 60 }
          }));
        }
      }
    }

    // 错题汇总（Word 试卷）
    if (withWrong) {
      let hasWrong = false;
      for (const qtype of typeOrder) {
        const questions = results[qtype];
        if (!questions) continue;
        for (const q of questions) {
          if (q.isWrong) { hasWrong = true; break; }
        }
        if (hasWrong) break;
      }

      if (hasWrong) {
        children.push(new Paragraph({
          children: [new PageBreak()],
          spacing: { after: 0 }
        }));
        children.push(new Paragraph({
          children: [new TextRun({ text: '错题汇总', font: "宋体", size: 32, bold: true, color: "000000" })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 240 }
        }));

        let globalNum = 0;
        for (const qtype of typeOrder) {
          const questions = results[qtype];
          if (!questions || questions.length === 0) continue;
          if (qtype === '简答') { globalNum += questions.length; continue; }

          const header = typeHeaders[qtype] || qtype;
          let sectionHasWrong = false;
          for (const q of questions) {
            if (q.isWrong) { sectionHasWrong = true; break; }
          }
          if (!sectionHasWrong) { globalNum += questions.length; continue; }

          children.push(new Paragraph({
            children: [new TextRun({ text: header, font: "宋体", size: 28, bold: true, color: "000000" })],
            spacing: { after: 120 }
          }));

          for (const q of questions) {
            globalNum++;
            if (!q.isWrong) continue;
            children.push(new Paragraph({
              children: await buildRichRuns(questionContent(q), `${globalNum}. `),
              spacing: { after: 40 }
            }));
            const options = q.options || [];
            if (options.length > 0) {
              for (const opt of options) {
                children.push(new Paragraph({
                  children: await buildRichRuns(optionContent(opt), `${opt.letter}. `),
                  spacing: { after: 40 }
                }));
              }
            }
            children.push(new Paragraph({
              children: [new TextRun({ text: `我的答案: ${q.myAnswer || '无'}`, font: "宋体", size: 24, color: "DC2626" })],
              spacing: { after: 40 }
            }));
            const correctRuns = await buildRichRuns(answerContent(q));
            if (correctRuns.length === 0) {
              correctRuns.push(new TextRun({ text: '（未找到答案）', font: "宋体", size: 24, color: "16A34A" }));
            }
            children.push(new Paragraph({
              children: [new TextRun({ text: '正确答案: ', font: "宋体", size: 24, color: "16A34A" }), ...correctRuns],
              spacing: { after: 120 }
            }));
          }
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

  let _creatingPanel = false;

  function createPanel() {
    if (_creatingPanel) return;
    if (document.getElementById('xxt-panel-btn')) return;

    _creatingPanel = true;
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
        <div style="display:flex;align-items:center;gap:4px;">
          <button class="xxt-settings-btn" id="xxt-historyBtn" title="历史记录">
            <svg viewBox="0 0 24 24"><path d="M13 3a9 9 0 0 0-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42A8.954 8.954 0 0 0 13 21a9 9 0 0 0 0-18zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
          </button>
          <button class="xxt-settings-btn" id="xxt-settingsBtn" title="设置">
            <svg viewBox="0 0 24 24"><path d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.61l-1.92-3.32a.49.49 0 0 0-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54a.484.484 0 0 0-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.07.62-.07.94s.02.64.07.94l-2.03 1.58a.49.49 0 0 0-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/></svg>
          </button>
          <button class="xxt-close-btn" id="xxt-closeBtn">&times;</button>
        </div>
      </div>
      <button id="xxt-btnExtract" class="xxt-btn xxt-btn-extract">提取本页题目</button>
      <div id="xxt-status" class="xxt-hidden"></div>
      <div id="xxt-result" class="xxt-hidden">
        <div id="xxt-stat" class="xxt-stat"></div>
        <div class="xxt-section">
          <input type="text" id="xxt-filename" class="xxt-filename xxt-hidden" placeholder="文件名（默认自动生成）">
          <div class="xxt-format-row">
            <span>输出格式</span>
            <label><input type="radio" name="xxt-fmt" value="word" checked> Word 试卷</label>
            <label><input type="radio" name="xxt-fmt" value="txt"> TXT</label>
            <label><input type="radio" name="xxt-fmt" value="md"> MD</label>
          </div>
        </div>
        <label class="xxt-toggle" id="xxt-ans-toggle">
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
        <label class="xxt-toggle" id="xxt-bank-import-toggle">
          <input type="checkbox" id="xxt-chkBankImport">
          <div class="xxt-checkbox-wrap"></div>
          <span>题库导入格式（学习通智能导入兼容）</span>
        </label>
        <div class="xxt-actions">
          <button id="xxt-btnDownload" class="xxt-btn xxt-btn-outline">&#8681; 下载</button>
          <button id="xxt-btnCopy" class="xxt-btn xxt-btn-outline">&#128203; 复制文本</button>
        </div>
        <div id="xxt-wrong-hint" class="xxt-wrong-hint xxt-hidden"></div>
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
      chkBankImport: $('xxt-chkBankImport'),
      bankImportToggle: $('xxt-bank-import-toggle'),
      closeBtn: $('xxt-closeBtn'),
    };

    btn.addEventListener('click', () => { panel.classList.toggle('open'); });
    els.closeBtn.addEventListener('click', () => { panel.classList.remove('open'); });
    // 点击面板外部关闭面板，但不影响设置弹窗和历史记录弹窗内的操作
    document.addEventListener('click', (e) => {
      if (!panel.contains(e.target) && e.target !== btn
        && !modal.contains(e.target) && !historyModal.contains(e.target)) {
        panel.classList.remove('open');
      }
    });

    // ==================== 设置弹窗 ====================
    // 创建设置弹窗
    const modal = document.createElement('div');
    modal.id = 'xxt-settings-modal';
    modal.innerHTML = `
      <div class="xxt-modal-box">
        <div class="xxt-modal-header">
          <h3>设置</h3>
          <button class="xxt-modal-close" id="xxt-modal-close">&times;</button>
        </div>
        <div class="xxt-setting-row">
          <span class="xxt-setting-label">主题</span>
          <div class="xxt-theme-group" id="xxt-theme-group">
            <button class="xxt-theme-btn" data-theme="auto">自动</button>
            <button class="xxt-theme-btn" data-theme="light">浅色</button>
            <button class="xxt-theme-btn" data-theme="dark">深色</button>
          </div>
        </div>
        <div class="xxt-setting-row">
          <span class="xxt-setting-label">提取快捷键</span>
          <div class="xxt-shortcut-display" id="xxt-shortcut-display">
            <span class="xxt-shortcut-keys" id="xxt-shortcut-keys"></span>
            <span class="xxt-shortcut-hint" id="xxt-shortcut-hint">点击修改</span>
          </div>
        </div>
        <div class="xxt-setting-row">
          <span class="xxt-setting-label">隐藏浮窗快捷键</span>
          <div class="xxt-shortcut-display" id="xxt-hide-shortcut-display">
            <span class="xxt-shortcut-keys" id="xxt-hide-shortcut-keys"></span>
            <span class="xxt-shortcut-hint" id="xxt-hide-shortcut-hint">点击修改</span>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(modal);

    // 初始化主题按钮状态
    const themeBtns = modal.querySelectorAll('.xxt-theme-btn');
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
    const shortcutDisplay = modal.querySelector('#xxt-shortcut-display');
    const shortcutKeysEl = modal.querySelector('#xxt-shortcut-keys');
    const shortcutHintEl = modal.querySelector('#xxt-shortcut-hint');
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

    // 隐藏浮窗快捷键初始化与录制
    const hideShortcutDisplay = modal.querySelector('#xxt-hide-shortcut-display');
    const hideShortcutKeysEl = modal.querySelector('#xxt-hide-shortcut-keys');
    const hideShortcutHintEl = modal.querySelector('#xxt-hide-shortcut-hint');
    function updateHideShortcutUI() {
      hideShortcutKeysEl.textContent = formatShortcutLabel(currentSettings.hideShortcut);
    }
    updateHideShortcutUI();

    hideShortcutDisplay.addEventListener('click', () => {
      hideShortcutHintEl.textContent = '请按下新快捷键...';
      hideShortcutKeysEl.textContent = '...';
      hideShortcutDisplay.style.borderColor = '#1e88e5';
      window.__xxt_recording = true;

      function onKeyDown(e) {
        e.preventDefault();
        e.stopPropagation();
        const key = e.key;
        if (['Control', 'Shift', 'Alt', 'Meta'].includes(key)) return;

        currentSettings.hideShortcut = {
          ctrl: e.ctrlKey, shift: e.shiftKey, alt: e.altKey, key: key.toLowerCase()
        };
        saveSettings(currentSettings);
        updateHideShortcutUI();
        hideShortcutHintEl.textContent = '点击修改';
        hideShortcutDisplay.style.borderColor = '';
        window.__xxt_recording = false;
        document.removeEventListener('keydown', onKeyDown, true);
      }
      document.addEventListener('keydown', onKeyDown, true);

      setTimeout(() => {
        if (window.__xxt_recording) {
          window.__xxt_recording = false;
          document.removeEventListener('keydown', onKeyDown, true);
          updateHideShortcutUI();
          hideShortcutHintEl.textContent = '点击修改';
          hideShortcutDisplay.style.borderColor = '';
        }
      }, 5000);
    });

    // 弹窗打开/关闭
    const settingsBtn = document.getElementById('xxt-settingsBtn');
    const modalClose = modal.querySelector('#xxt-modal-close');

    settingsBtn.addEventListener('click', () => {
      updateThemeUI();
      updateShortcutUI();
      modal.classList.add('open');
    });
    modalClose.addEventListener('click', () => { modal.classList.remove('open'); });
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.classList.remove('open');
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && modal.classList.contains('open')) {
        modal.classList.remove('open');
      }
    });

    // ==================== 历史记录弹窗 ====================
    const historyModal = document.createElement('div');
    historyModal.id = 'xxt-history-modal';
    historyModal.innerHTML = `
      <div class="xxt-modal-box">
        <div class="xxt-modal-header">
          <h3>历史记录</h3>
          <button class="xxt-modal-close" id="xxt-history-close">&times;</button>
        </div>
        <div class="xxt-history-list" id="xxt-history-list">
          <div class="xxt-history-empty">暂无历史记录</div>
        </div>
      </div>
    `;
    document.body.appendChild(historyModal);

    const historyListEl = historyModal.querySelector('#xxt-history-list');
    const historyCloseBtn = historyModal.querySelector('#xxt-history-close');

    function renderHistoryList() {
      const history = loadHistory();
      if (history.length === 0) {
        historyListEl.innerHTML = '<div class="xxt-history-empty">暂无历史记录</div>';
        return;
      }
      const fmtNames = { 'txt': 'TXT', 'md': 'MD', 'word': 'Word' };
      historyListEl.innerHTML = history.map(h => {
        const flags = [];
        if (h.withAnswers) flags.push('含答案');
        if (h.withWrong) flags.push('含错题');
        if (h.shuffle) flags.push('打乱');
        if (h.bankImport) flags.push('题库导入');
        const flagStr = flags.length > 0 ? ' · ' + flags.join('、') : '';
        return `
          <div class="xxt-history-item" data-id="${h.id}">
            <div class="xxt-history-info">
              <div class="xxt-history-title">${escapeHtml(h.title)}</div>
              <div class="xxt-history-meta">${h.date} · ${h.totalQuestions}题 · ${fmtNames[h.format] || h.format}${flagStr}</div>
            </div>
            <button class="xxt-history-delete" data-id="${h.id}" title="删除">✕</button>
          </div>
        `;
      }).join('');

      // 点击历史条目 → 加载
      historyListEl.querySelectorAll('.xxt-history-item').forEach(item => {
        item.addEventListener('click', (e) => {
          if (e.target.closest('.xxt-history-delete')) return;
          const id = Number(item.dataset.id);
          const entry = loadHistory().find(h => h.id === id);
          if (!entry) return;
          // 恢复提取数据
          const typeOrder = entry.typeOrder;
          const results = entry.results;
          const total = entry.totalQuestions;
          const hasCorrectAnswer = entry.hasCorrectAnswer !== false;
          extractedData = {
            total, title: entry.title, typeOrder, results,
            wrongCount: entry.wrongCount, hasMyAnswer: entry.hasMyAnswer, hasCorrectAnswer,
            breakdown: Object.fromEntries(Object.entries(results).map(([k, v]) => [k, v.length])),
            text: formatOutput(results, typeOrder),
            textWithAnswers: formatOutputWithAnswers(results, typeOrder),
            textWrong: formatWrongQuestionsTXT(results, typeOrder),
            textMD: formatOutputMD(results, typeOrder),
            textWithAnswersMD: formatOutputWithAnswersMD(results, typeOrder),
            textWrongMD: formatWrongQuestionsMD(results, typeOrder),
          };
          // 恢复格式选项
          const fmtRadio = document.querySelector(`input[name="xxt-fmt"][value="${entry.format}"]`);
          if (fmtRadio) fmtRadio.checked = true;
          // 恢复复选框
          if (els.chkAnswers) {
            els.chkAnswers.checked = hasCorrectAnswer ? entry.withAnswers : false;
            els.chkAnswers.disabled = !hasCorrectAnswer;
          }
          const ansToggle = document.getElementById('xxt-ans-toggle');
          if (ansToggle) ansToggle.classList.toggle('xxt-disabled', !hasCorrectAnswer);
          if (els.chkShuffle) els.chkShuffle.checked = entry.shuffle;
          if (els.chkWrong) els.chkWrong.checked = entry.withWrong;
          if (els.chkBankImport) els.chkBankImport.checked = entry.bankImport || false;
          // 根据格式调整 UI
          const isWord = entry.format === 'word';
          if (els.wrongToggle) els.wrongToggle.style.display = isWord ? 'none' : '';
          if (els.wrongHint) els.wrongHint.style.display = isWord ? 'none' : '';
          // 题库导入格式选项仅在 Word 格式下显示
          if (els.bankImportToggle) els.bankImportToggle.style.display = isWord ? '' : 'none';
          // 题库导入格式：勾选时禁用打乱和答案选项（不隐藏，维持高度稳定）
          const isBankImport = entry.bankImport || false;
          const shuffleToggle = document.getElementById('xxt-shuffle-toggle');
          if (ansToggle) {
            ansToggle.classList.toggle('xxt-disabled', isBankImport || !hasCorrectAnswer);
            if (ansToggle.querySelector('input')) ansToggle.querySelector('input').disabled = isBankImport || !hasCorrectAnswer;
          }
          if (shuffleToggle) {
            shuffleToggle.classList.toggle('xxt-disabled', isBankImport);
            if (shuffleToggle.querySelector('input')) shuffleToggle.querySelector('input').disabled = isBankImport;
          }
          if (els.wrongToggle) {
            els.wrongToggle.classList.toggle('xxt-disabled', isBankImport);
            if (els.wrongToggle.querySelector('input')) els.wrongToggle.querySelector('input').disabled = isBankImport;
          }
          // 题库导入格式下，错题提示不显示（选项已禁用）
          if (entry.hasMyAnswer && entry.wrongCount > 0 && !isBankImport) {
            els.wrongToggle.classList.remove('xxt-hidden');
            els.wrongHint.textContent = `检测到 ${entry.wrongCount} 道错题，可勾选附加到输出末尾`;
            els.wrongHint.classList.remove('xxt-hidden');
          } else {
            els.wrongToggle.classList.add('xxt-hidden');
            els.wrongHint.classList.add('xxt-hidden');
          }
          renderStats(els);
          updateFilename(els);
          els.result.classList.remove('xxt-hidden');
          showStatus(els, `已加载历史记录：${entry.title} (${total}题)`, 'ok');
          historyModal.classList.remove('open');
          els.btnExtract.textContent = '重新提取';
        });
      });

      // 删除按钮
      historyListEl.querySelectorAll('.xxt-history-delete').forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          const id = Number(btn.dataset.id);
          deleteHistoryById(id);
          renderHistoryList();
        });
      });
    }

    function escapeHtml(str) {
      const div = document.createElement('div');
      div.textContent = str;
      return div.innerHTML;
    }

    // 历史记录按钮 → 打开弹窗
    const historyBtn = document.getElementById('xxt-historyBtn');
    historyBtn.addEventListener('click', () => {
      renderHistoryList();
      historyModal.classList.add('open');
    });
    historyCloseBtn.addEventListener('click', () => { historyModal.classList.remove('open'); });
    historyModal.addEventListener('click', (e) => {
      if (e.target === historyModal) historyModal.classList.remove('open');
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && historyModal.classList.contains('open')) {
        historyModal.classList.remove('open');
      }
    });

    // 格式/选项切换时更新文件名
    panel.addEventListener('change', (e) => {
      if (!extractedData) return;
      if (e.target.name === 'xxt-fmt') {
        updateFilename(els);
        const isWord = getFormat(els) === 'word';
        // 题库导入格式选项仅在 Word 格式下显示
        if (els.bankImportToggle) els.bankImportToggle.style.display = isWord ? '' : 'none';
        if (!isWord && els.chkBankImport) els.chkBankImport.checked = false;
      }
      // 题库导入格式切换：勾选时禁用打乱和答案选项（不隐藏，维持高度稳定）
      if (e.target === els.chkBankImport) {
        updateFilename(els);
        const checked = els.chkBankImport.checked;
        const hasCorrectAnswer = extractedData.hasCorrectAnswer !== false;
        const ansToggle = document.getElementById('xxt-ans-toggle');
        const shuffleToggle = document.getElementById('xxt-shuffle-toggle');
        const wrongToggle = document.getElementById('xxt-wrong-toggle');
        if (ansToggle) {
          ansToggle.classList.toggle('xxt-disabled', checked || !hasCorrectAnswer);
          if (ansToggle.querySelector('input')) ansToggle.querySelector('input').disabled = checked || !hasCorrectAnswer;
        }
        if (shuffleToggle) {
          shuffleToggle.classList.toggle('xxt-disabled', checked);
          if (shuffleToggle.querySelector('input')) shuffleToggle.querySelector('input').disabled = checked;
        }
        if (wrongToggle) {
          wrongToggle.classList.toggle('xxt-disabled', checked);
          if (wrongToggle.querySelector('input')) wrongToggle.querySelector('input').disabled = checked;
        }
        if (checked) {
          if (els.chkAnswers) els.chkAnswers.checked = false;
          if (els.chkShuffle) els.chkShuffle.checked = false;
          if (els.chkWrong) els.chkWrong.checked = false;
        }
      }
      // 勾选/取消附加答案时更新文件名
      if (e.target === els.chkAnswers) {
        updateFilename(els);
      }
    });

    els.btnExtract.addEventListener('click', () => {
      els.btnExtract.disabled = true;
      els.btnExtract.textContent = '提取中...';
      els.status.className = 'xxt-hidden';
      els.result.classList.add('xxt-hidden');

      const { results, typeOrder, wrongCount, hasMyAnswer, hasCorrectAnswer } = extract();
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
        total, title, typeOrder, results, wrongCount, hasMyAnswer, hasCorrectAnswer,
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

      // 没有正确答案时禁用“附加答案”
      if (els.chkAnswers) {
        els.chkAnswers.checked = hasCorrectAnswer ? els.chkAnswers.checked : false;
        els.chkAnswers.disabled = !hasCorrectAnswer;
      }
      const ansToggle = document.getElementById('xxt-ans-toggle');
      if (ansToggle) ansToggle.classList.toggle('xxt-disabled', !hasCorrectAnswer);

      // 错题按钮逻辑（题库导入格式下不显示）
      const isBankImport = els.chkBankImport && els.chkBankImport.checked;
      if (hasMyAnswer && wrongCount > 0 && !isBankImport) {
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
          const isBankImport = els.chkBankImport && els.chkBankImport.checked;
          // 题库导入格式：不打乱、始终含答案
          const doShuffle = !isBankImport && els.chkShuffle && els.chkShuffle.checked;
          const withWrong = !isBankImport && els.chkWrong && els.chkWrong.checked;
          const activeResults = doShuffle
            ? shuffleQuestions(extractedData.results, extractedData.typeOrder)
            : extractedData.results;
          const blob = await generateWordBlob(activeResults, extractedData.typeOrder, extractedData.title, isBankImport || els.chkAnswers.checked, withWrong, isBankImport);
          const filename = (els.filename.value || '学习通试卷').replace(/\.(txt|md|docx)$/, '') + '.docx';
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url; a.download = filename; a.click();
          URL.revokeObjectURL(url);
          showStatus(els, isBankImport ? '题库导入格式已下载' : 'Word 试卷' + (els.chkAnswers.checked ? '（含答案）' : '') + (withWrong ? '（含错题）' : '') + '已下载', 'ok');
          addToHistory(extractedData, 'word', isBankImport || els.chkAnswers.checked, withWrong, doShuffle, isBankImport, '');
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
      addToHistory(extractedData, fmt, els.chkAnswers.checked, els.chkWrong && els.chkWrong.checked, els.chkShuffle && els.chkShuffle.checked, false, text);
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

    // 设置自动提取回调：iframe 发来新数据时自动触发提取
    window.__xxt_auto_extract = function() {
      const btn = document.getElementById('xxt-btnExtract');
      if (btn && !btn.disabled) btn.click();
    };
    _creatingPanel = false;
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
    // 题库导入格式用"（题库导入）"后缀，附加答案用"（含答案）"
    const isBankImport = els.chkBankImport && els.chkBankImport.checked;
    const suffix = isBankImport ? '（题库导入）' : (els.chkAnswers.checked ? '（含答案）' : '');
    els.filename.value = cleanTitle + suffix + ext;
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
  // 检测是否在 iframe 中运行
  let isInIframe = false;
  try { isInIframe = window.top !== window.self; } catch(e) { isInIframe = true; }

  let initTimer = null;
  let observer = null;

  if (isInIframe) {
    // 在 iframe 中：提取题目并通过 postMessage 回传给父窗口，不创建 UI
    let iframeDebounceTimer = null;
    function iframeExtract() {
      const result = extractFromRoot(document);
      if (!hasQuestions(result)) return;
      // 获取标题（兼容作业详情页与章节测验页）
      const titleEl = document.querySelector('.mark_title, .newTestTitle, .TestTitle_name');
      const title = titleEl ? titleEl.textContent.trim() : '';
      // 直接发送最新数据，父窗口以最后收到的为准
      window.parent.postMessage({
        type: 'xxt-iframe-result',
        data: { ...result, title }
      }, '*');
    }

    // 监听 DOM 变化，章节切换时自动重新提取并回传
    // 始终挂载在 document.body 上，避免 #ZyBottom 被替换后 observer 失效
    function observeForQuestions() {
      const mo = new MutationObserver(() => {
        if (iframeDebounceTimer) clearTimeout(iframeDebounceTimer);
        iframeDebounceTimer = setTimeout(iframeExtract, 300);
      });
      mo.observe(document.body, { childList: true, subtree: true });
    }

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => {
        iframeExtract();
        observeForQuestions();
      });
    } else {
      iframeExtract();
      observeForQuestions();
    }
    // 延迟重试，应对 iframe 内容动态加载
    setTimeout(iframeExtract, 1500);
    setTimeout(iframeExtract, 3000);
    // 定时轮询兜底，应对 SPA 切换章节时 MutationObserver 遗漏
    setInterval(iframeExtract, 5000);
  } else {
    // 在顶层窗口中：正常创建 UI 并监听 iframe 回传数据
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
  }

})();