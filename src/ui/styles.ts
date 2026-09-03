export const PANEL_STYLES = `
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
.cwe-subtitle { margin-top: 2px; color: var(--cwe-text-muted); font-size: 11px; }
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
.cwe-check { min-height: 34px; display: flex; align-items: center; gap: 8px; padding: 7px 8px; border-radius: 8px; color: var(--cwe-text-secondary); cursor: pointer; font-size: 11.5px; }
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

.cwe-list { display: grid; gap: 8px; }
.cwe-list-empty { padding: 30px 12px; color: var(--cwe-text-muted); text-align: center; font-size: 12px; }
.cwe-history-item { padding: 11px; border: 1px solid var(--cwe-border); border-radius: 11px; background: var(--cwe-bg-soft); }
.cwe-history-title { overflow: hidden; color: var(--cwe-text); font-size: 12px; font-weight: 800; text-overflow: ellipsis; white-space: nowrap; }
.cwe-history-meta { margin-top: 4px; color: var(--cwe-text-muted); font-size: 10.5px; }
.cwe-history-actions { display: flex; gap: 6px; margin-top: 9px; }
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
