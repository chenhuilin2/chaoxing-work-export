# Changelog

## 3.0.9

### 修复

- **标题仍显示不出章节名**：上版只认了 `#prev_title`（id），而真实 DOM 用的是类名 `<div class="prev_title" title="毛泽东思想的主要内容">`，且该节点在外层壳文档里，题目却位于嵌套的知识卡片 / 答题 iframe —— 嵌套文档只能解析出自己的标题（「章节测验 待完成」）。现在：
  - `resolvePrevTitle()` 同时匹配 `#prev_title` 与 `.prev_title`（title 属性优先，回退正文）；
  - `ExtractionService` 在选出候选结果后用外层壳的章节名覆盖标题（`applyScopeTitle()`），因此**题目在嵌套 iframe 里也能显示章节名**，并直接决定导出文件名与 Word 试卷标题。
- **标题不再参与题目指纹**（`ExtractionService.fingerprint()`）：章节名的更新早于题目 DOM，纳入指纹会把上一章残留的题目误判成本章内容。

### 优化

- **章节切换提速**：去掉「切一次卡 + 固定长等待」的节奏（旧流程先等任务卡栏 ≤2.5s，再固定等 14–20s），改为采样等待：
  - `ExtractionService.settleQuestions()` 每 120ms 采样一次跨层提取，**题目一出现立即返回**，检测延迟收敛到一个采样间隔；
  - 新增 `QuestionTabGuard`：每轮补切一次题目卡（单章最多 4 次、间隔 ≥400ms），解决「章节切换后卡片页重建、默认回到视频卡，只切一次的任务卡随旧页面失效」的问题；
  - 单次提取同样改为采样等待，并且不再先空等 4.5s 才去尝试切卡；页面确实没有题目卡时只等 4.5s。
- 移除已被取代的 `ExtractionService.waitForChangedResult()`（其「固定间隔长等待」语义与旧标题指纹规则已不再需要）。

### 验证

- 新增单元测试 `tests/extraction-settle.test.cjs`（6 例）：嵌套 iframe 场景下标题取外层壳章节名、无 `.prev_title` 时行为不变、题目就绪时采样立即返回（<300ms）、无题目时按超时结束、**章节切换后自动补切任务卡并秒级拿到题目**、章节已激活时不空等。`tests/task-tab-locator.test.cjs` 增补补切节流用例（限次 + 冷却）。
- 单测 **55 → 64 例，全绿**；typecheck 与 ESLint（`src`）0 问题；构建产物 434,555 → **437,908 字节**。
- 真实快照端到端验证（`.workbuddy/verify-prev-title.cjs`）：`html/章节测验.html` 提取标题由「章节测验 待完成」变为 **「毛泽东思想的主要内容」**（题数 5 不变）；`作业页面.html`（44 题，标题「章节测验 待完成」）与 `考试页面.html`（56 题，标题「学习通题目」）**无回归**；一次采样提取耗时 56–92ms（合并后的 11.3MB 巨型快照，真实页面各层文档更小）。

## 3.0.8

### 修复

- **不再需要手动切到「章节测验」**：知识卡片页默认停在「视频」卡，而任务卡内容是懒加载的 —— 卡片内部 iframe 的真实地址写在 `_src` 上，只有点击该卡触发页面自身的 `changeDisplayContent()` 之后才会写回 `src` 并加载答题页。因此不切卡时页面里根本不存在题目 DOM。现在提取前会自动定位并激活承载题目的任务卡（章节测验 / 作业 / 考试 / 练习），单次提取与批量提取都生效。
- **修复「一键提取多个章节」切不动章节**：`chapter-locator.activate()` 原先对容器 `div.posCatalog_select` 派发点击，但学习通把 `onclick` 挂在内部的 `span.posCatalog_name` 上，事件不会向下传递，章节实际没有切换，每章只能空等超时。现在优先选中带 `onclick` 的节点。
- **修复章节列表把分组标题当章节**：`.posCatalog_select` 会把第一级分组标题（`.firstLayer` / `.posCatalog_title`）一起收进来（真实快照 64 项里 10 项是分组标题）。现在优先取 `span.posCatalog_name` 并显式排除分组标题，快照实测 64 → 54 项。
- 章节名兜底补充 `textContent`：章节名节点不带 `title` 属性时不再退化成「第 N 章」。

### 新增

- 新增 `application/task-tab-locator.ts`：任务卡（任务点）定位器。跨同源 iframe 找到 `#prev_tab`，按文案匹配题目类任务卡，返回 `active` / `switched` / `missing` 三态（已激活 / 本次切换 / 页面无可用任务卡）。
- `utils/dom.ts` 新增 `collectAccessibleDocuments()`：跨层收集可同步访问的文档（含同源 iframe）。章节与任务卡定位、`ExtractionService` 共用同一套 iframe 遍历规则（后者改为复用，行为不变）。
- 批量提取流程升级为「切章节 → 等任务卡栏重建 → 切题目卡 → 等题目 → 提取」，收尾按「先还原章节、再还原任务卡」恢复用户原有浏览状态；任务卡栏只在知识卡片页存在，因此批量开始前探测一次，避免逐章空等。
- 批量结果标题规则明确为：单章节直接用章节名，多章节用「页面标题 - N 个章节」（页面标题已内含 `#prev_title` 优先级）。

### 验证

- 新增单元测试 `tests/chapter-locator.test.cjs`（5 例）与 `tests/task-tab-locator.test.cjs`（7 例）：覆盖分组标题排除、激活态识别、点击目标必须是带 `onclick` 的节点、任务卡三态、文案回退（`title` → `spanText` → 去序号的正文）、同源 iframe 内定位。
- 单测 **43 → 55 例，全绿**；typecheck 与 ESLint（`src`）0 问题；构建产物 421,223 → **434,555 字节**。
- 真实快照 `html/章节测验.html`（11.3 MB，多层 iframe 合并）端到端验证（`.workbuddy/verify-chapter-flow.cjs`）：目录 64 个容器 → 精确识别 54 个章节（等于 `span.posCatalog_name` 节点数，差额正好是 10 个分组标题）；3 个任务卡栏全部解析正确（默认停在视频卡的栏返回 `switched`，已停在章节测验的栏返回 `active`）；`activate()` 的点击目标确认为 `SPAN.posCatalog_name`（带 `onclick`），不再是容器 div。
- 待真机确认：`changeDisplayContent` 的函数定义不在快照中（外链 JS），「切卡把 `_src` 写回 `src`」这条链路依据的是页面自身的懒加载约定与 `#frame_content` 的实际取值，需在真实页面最终确认。

## 3.0.7

### 新增

- 学生学习页面（章节练习 / 章节测验 / 做作业）提取的标题优先取自 `#prev_title` 容器：先取其 `title` 属性，缺失时回退到其正文。该容器承载章节/测验的真实名称，比浏览器默认的 `document.title`（多为课程名或「超星学习通」）更准确，导出文件名与 Word 试卷标题据此生成。无 `#prev_title` 时仍走原有选择器与 `document.title` 兜底，其他页面行为不变。

### 验证

- 新增单元测试 `tests/page-context.test.cjs`（3 例）：覆盖「`#prev_title` 的 title 属性优先」「无 title 属性回退正文」「无 `#prev_title` 时原有逻辑不变」。
- 单测 **38 → 41 例，全绿**；typecheck 与 ESLint 0 问题；构建产物已更新。

## 3.0.6

### 修复

- 修复「章节测验 / 作业页面」题干与选项导出为乱码的问题。学习通对答题页启用了 `font-cxsecret` 字体反爬：把部分汉字换成「备用码位」，再用内联字体把该码位画成真字的轮廓，因此人眼所见正常而 `textContent` 读出乱码（如「中悢共产党悤史上」）。原先提取器读到什么就导出什么，导出的题干、选项文字不可用。诊断与验证过程见 `docs/chapter-quiz-diagnosis.md`。

### 新增

- 新增 `extractors/cxsecret-font.ts`：字体反爬解码的纯逻辑层。利用「子集字体的 `glyf` 表逐字节保留了源字体字形数据」这一特性，用
  `sha1(glyf[loca[gid] : loca[gid + 1]])[:4]` 反查真实汉字码位。内含自实现的同步 SHA-1（只取摘要首 4 字节）与最小 sfnt 解析（`head`/`maxp`/`loca`/`glyf`/`cmap`，支持 cmap format 4 与 12、支持 TTC）。
- 新增 `extractors/cxsecret-table.ts`（生成物）：`sha1[:4] → BMP 码位` 解码表，覆盖 U+4E00–U+9FA5 共 20902 字，4 字节键经生成器校验**零冲突**，二进制紧凑存放（122.5 KB 裸数据 / base64 163.3 KB），运行时二分查找。
- 新增 `extractors/cxsecret-decoder.ts`：DOM 接入层。从文档的 `<style>`（及同源 `styleSheets`）中收集内联 `data:` 字体，构建映射并按文档缓存；**只还原处于 cxsecret 作用域内的文本**（自身或祖先类名含 `cxsecret`），避免把作用域外的同码位正常文字改坏；页面若把该类名改掉则退化为整页还原。家族名不匹配时回退尝试全部内联字体，且要求命中数达标，避免误伤普通内嵌字体。
- 新增 `tools/gen-cxsecret-table.py`：解码表生成器，支持两条路径 —— `--font <字节兼容的源字体>`（只依赖 OFL 字体，推荐）与 `--pkl <目录>`（用预生成哈希表重建），生成时校验键冲突。
- `extractors/rich-content.ts` 接入解码：文本节点按其所属元素作用域还原，`aria-label` 等属性兜底路径同样还原（真实页面的选项文字同时存在于这两处）。
- 新增「提取多个章节」相关的解码覆盖：多层 iframe 各自持有独立文档与内联字体，解码器按文档缓存，互不干扰。
- 新增单元测试 `tests/cxsecret-decoder.test.cjs`（11 例）与字体夹具 `tests/fixtures/cxsecret-quiz.ttf`（真实章节测验页内联字体，20,756 字节）：覆盖 SHA-1 与 node crypto 一致、内置表条目与冲突、base64 解码、真实字体 83 条映射、作用域内外差异、无字体/字体损坏/类名被改的退化行为、`extractRichContent` 与 `TiMuExtractor` 端到端还原。

### 验证

- 单测 **27 → 38 例，全绿**；typecheck 与 ESLint 0 问题。
- 解码正确性以两条**互相独立**的方法交叉验证：字形哈希反查 vs 栅格形状匹配，`sha1[:4]` 结果与栅格真值 **83/83、154/154 完全一致**。
- 三份真实快照端到端：`html/章节测验.html`（四层嵌套）解出 83 条映射、提取 **5 题**且题干选项可读；`html/作业页面.html` 解出 154 条映射、提取 **44 题**；`html/考试页面.html`（无混淆）正确判定为无需解码、提取 **56 题**。
- 构建产物 227,687 → **420,573 字节**（增量主要为 163 KB 解码表）。浏览器冒烟测试失败点与改动前基线相同（既有的 `navigator.clipboard` 环境问题），提取链路无回归。

## 3.0.5

### 修复

- 修复「做作业 / 章节测验」答题页提取不到题目（面板提示「当前页面未识别到题目」）的问题。题干容器选择器只匹配 `.qtContent / .question-content / .mark_name`，而 2026 版答题页的题干在 `.Zy_TItle > div.fontLabel`，选择器落空 → 题干为空 → `createQuestion` 据此判定整题无效并丢弃 → 四类提取器全部返回 0 题。诊断过程与页面结构证据见 `docs/chapter-quiz-diagnosis.md`。

### 新增

- 新增 `extractors/question-stem.ts`：题干**分层定位**，取代四个提取器各自硬编码的单一选择器列表，以应对学习通频繁改版：
  1. 已知题干类名（新版 → 历史版本，集中在 `STEM_CONTAINER_SELECTORS` 一处维护）
  2. 题型标记的父元素（题干与「【单选题】」同级，题干容器改名也不受影响）
  3. 兜底：剔除题号、选项列表与答案区后，取标题容器 / 整题容器的剩余内容
  日后改版导致题干丢失时，只需追加一个类名，无需改动任何提取流程代码。定位结果同时返回「已剥离前缀的题干」与「原始文本」，题号解析结果与改动前一致。
- `mark-item-extractor` / `question-li-extractor` / `generic-exercise-extractor` 同步改用该定位器，删除各自重复的选择器列表。
- 新增提取器单元测试 `tests/question-stem.test.cjs`、`tests/question-extraction.test.cjs`（共 13 例），覆盖 2026 版真实结构、题干类名被改、题型标记被改、兜底不混入选项文本、图片题干、空题干不产出无效题目等场景。
- 引入仅测试使用的 `linkedom`（devDependency）为 Node 单测提供 DOM 环境；`tsconfig.test-build.json` 的 include 扩展为 `src/extractors/**/*.ts`，使提取器首次进入单测覆盖范围。构建产物不受影响。

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
