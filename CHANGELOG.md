# Changelog

## 3.0.4

### 功能

- 文件名预览优化：勾选「附加答案」时，文件名预览区自动追加「（含答案）」后缀；取消勾选或勾选「题库导入」时自动移除/切换，与主脚本 v1.7 行为一致。`updateFilenamePreview()` 在选项变化与提取完成时刷新，历史恢复后按恢复的勾选状态重新计算。
- 导出格式为 TXT/MD 时，「题库导入」选项保持始终禁用（灰显）并自动取消勾选，切回 Word 格式恢复可用。
- 历史记录弹窗 UI 对齐主脚本：弹窗改为 400px 白卡片（浅色）/ Catppuccin 深色卡片，列表项显示「时间 · 题数 · 格式 · 含答案/含错题/打乱/题库导入」标志，右侧提供圆形「重新下载 ⤓」「删除 ✕」图标按钮；点击条目本身恢复记录，点击图标分别触发重新下载/删除（与主脚本交互一致）。`panel-view.ts` 模板与 `renderHistory()` 重写，`styles.ts` 新增 `cwe-history-*` 样式。
- 历史记录弹窗宽度增至 460px。

### 新增

- 「打开窗口自动提取」设置（默认关闭）：开启后脚本加载时延迟 800ms 自动执行一次提取，后续页面变化不再自动提取。`UserSettings` 新增 `autoExtractOnLoad` 字段，`settings-repository` 兼容旧配置（缺省回落为关闭），设置弹窗新增开关行。

## 3.0.3

### 修复

- 修复 Word 导出时题目与选项之间出现多余空行的问题（与主脚本 v2.0.2 同步修复）：学习通题干 HTML 末尾常残留「换行 + 纯空白文本」，`normalizeRichContent` 只修剪首尾换行、漏掉被空白文本挡住的换行，Word 中渲染为段内空行；现改为首尾的换行与纯空白文本一并修剪（图片前的换行仍保留）。`extractors/rich-content.ts` 与 `exporters/legacy-bridge.ts` 两处同步修改，并新增单元测试覆盖。

## 3.0.2

### 修复

- 提取结果不含参考答案时，导出选项区自动隐藏「附加参考答案」和「附加错题汇总」两项，并强制取消勾选（含从历史记录恢复偏好的场景）；重新提取到含答案的题目后恢复显示。

## 3.0.1

### 重构

- 导出与下载逻辑改为直接移植主脚本 `chaoxing-work-export.user.js` 的实现，输出与主脚本完全一致：
  - TXT/Markdown 格式化（含题干与选项的双行输出、答案汇总、错题汇总）逐行移植。
  - Word 生成（`generateWordBlob` 全套：题库智能导入格式、试卷排版、答案页、错题汇总、图片按比例缩放与失败计数）逐行移植。
  - 打乱逻辑改回原版 `shuffleQuestions`（题型内 Fisher-Yates 打乱；题目与答案随打乱顺序生成，错题汇总始终用原始顺序）。
  - 下载方式对齐原版（临时 `<a>` 点击后立即释放 URL，多文件间隔 300ms）。
  - 下载/复制状态提示文案对齐原版（如「题库导入格式已下载」「（N 张图片加载失败）」「Word 格式不支持复制，请使用下载」）。
- 新增 `exporters/legacy-bridge.ts`：主脚本数据结构定义、领域模型到原版结构的转换，以及原版公共辅助函数（`questionContent`/`optionContent`/`answerContent`/`formatRichForText`/`formatRichForMD` 等）。
- 删除 `exporters/result-transformer.ts`（被原版打乱逻辑替代）。

### 修复

- 修复 `tools/test.mjs` 在 Windows 下调用 `tsc.cmd` 触发 EINVAL 的问题（改为 node 执行 tsc，与 build.mjs 一致）。
- 补充根 `tsconfig.json`，修复 ESLint projectService 无法定位项目导致 lint 失败的问题。

## 3.0.0

### 重构

- 将单文件 JavaScript 重构为严格 TypeScript 多模块工程。
- 建立 domain、extractors、application、exporters、infrastructure、ui 六层结构。
- 构建阶段打包为单个可安装 userscript。
- 使用 Shadow DOM 隔离界面样式。
- 将 iframe 通信升级为带协议版本和来源校验的结构化消息。
- 将章节发现、章节遍历与单页解析拆分。
- 将 Word、TXT、Markdown 统一到同一题目领域模型。
- 增加版本化设置与历史仓储以及旧数据迁移。

### 功能

- 增加答案解析提取和导出。
- 增加通用练习页面兜底适配器。
- 增加按章节拆分文件。
- 增加 CSS 背景图片和文本填空输入提取。
- 增加图片请求缓存、超时和失败统计。
- 保留参考答案、错题汇总、题型内随机排序和题库智能导入格式。

### 工程化

- TypeScript strict 配置。
- ESLint、Prettier、EditorConfig。
- Node.js 单元测试。
- GitHub Actions 校验与构建产物上传。
- Dependabot 依赖更新。
- 架构、贡献、解析器扩展和安全文档。
