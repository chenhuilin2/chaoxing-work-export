# 会话交接单（2026-09-17 ~ 09-18）

> 用途：本会话上下文将满，此文档用于让下一次会话**无需重读全部历史**即可接手。
> 新会话开始请先读：本文件 → `AGENTS.md` → `CHANGELOG.md`（最新条目 3.0.5）。

---

## 1. 本次会话干了什么

主题：**学生学习页面（章节练习/章节测验）提取失败的诊断与修复**。

分五个阶段：

1. **目录扫描**（09-17 下午）→ 摸清项目结构、目标页面 DOM、`chapter-locator` 的两处 P0 → `docs/chapter-exercise-scan.md`
2. **快照质量排雷**（09-17 晚）→ 用户第一次给的 `章节测验.html` 是空壳（无 AJAX DOM），无法诊断；指导用户重新采集 → `chapter-exercise-scan.md`
3. **根因确诊**（09-17 晚）→ 用户回传四层嵌套有效快照，定位 RC-1 与 RC-2 → `docs/chapter-quiz-diagnosis.md`
4. **实施 RC-1**（09-18）→ 按用户「要做改版兼容」的要求，新增题干分层定位模块并接入四个提取器，单测 14→27，真实快照验证 5/5 题 → **已完成并构建**
5. **实施 RC-2**（09-18）→ 破解 `font-cxsecret` 字体反爬（字形哈希解码表 + 作用域感知解码器），单测 27→38，三份真实快照端到端通过 → **已完成并构建**

---

## 2. 项目速览（接手用）

- 学习通作业/考试/章节练习题目提取导出油猴脚本，TypeScript 六层架构，`npm run validate` = typecheck + 单测 + 构建
- 构建产物：`dist/chaoxing-work-export.user.js`（当前 **227,687 字节**）
- 远程：`github.com/chenhuilin2/chaoxing-work-export`（远端主分支 `main`，本地 `master`）
- 版本记录：**新版本一律写 `CHANGELOG.md`**（`version.md` 是旧单文件脚本时代的历史档案）

### ⚠️ 本机环境注意事项（踩过的坑，务必先看）

| 问题 | 应对 |
|---|---|
| **`npm` 在本会话 shell 中不可用** | 不要用 `npm run xxx`。改为直接用 node 调底层脚本：<br>`& <node> "node_modules\typescript\bin\tsc" -p tsconfig.typecheck.json --noEmit`<br>`& <node> "node_modules\eslint\bin\eslint.js" src --max-warnings=0`<br>`& <node> "tools\test.mjs"` / `& <node> "tools\build.mjs"` |
| node 路径 | `C:\Users\chenh\.workbuddy\binaries\node\versions\22.22.2-3\node.exe` |
| Python 路径 | venv：`C:\Users\chenh\.workbuddy\binaries\python\envs\default\Scripts\python.exe`（已装 fonttools/Pillow/numpy） |
| **Bash 会吞反斜杠** | 不要在 Bash 里写 Windows 路径字面量，写成 `.py` 脚本再执行 |
| **Bash 缺常用命令** | 该环境 `mkdir` / `cp` / `head` / `dirname` 均不可用（`command not found`）。文件操作、取前 N 行等一律改用 Python 完成 |
| **PowerShell 输出捕获不稳** | `*>&1 \| Out-File -Encoding utf8 <log>` 写文件，再用 Read 读 |
| 浏览器冒烟测试 | `tools/browser-smoke.mjs` **基线即失败**（`Copy formatter failed`），原因是 headless 下 `navigator.clipboard` 无法被 fixture 覆盖。与提取改动无关，**判断回归时以「失败点是否与基线相同」为准**。`CHROME_BIN` 可用 `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe` |

---

## 3. 诊断结论（四条）

### RC-1 ✅ 已修复 —— 题干选择器落空

- 真实题干容器：`.Zy_TItle > div.fontLabel`
- 旧代码：`title.querySelector('.qtContent, .question-content, .mark_name')` → **0 命中**
- 断链：`stemElement=null` → `extractRichContent(null)=[]` → `question-factory.ts:28` `isRichContentEmpty(stem)` 直接 `return null` → 5 题全丢 → 面板报「未识别到题目」
- **与 iframe 嵌套层级无关**，`Zy_TItle`/`Zy_ulTop`/`newZy_TItle`/`i.fl` 本来都能命中
- 同一结构也存在于 `html/做作业.html` → 影响**整类答题页**，不只章节测验

### RC-2 ✅ 已实施 —— `font-cxsecret` 字体反爬

- 页面内联 `@font-face`（`font-family: 'font-cxsecret'`，`data:application/font-ttf;charset=utf-8;base64,...`），
  字体为 **Source Han Sans CN Normal v1.000 子集**（`name` 表可查证；20,756 B / 83 码位 / `indexToLocFormat=1` / cmap format 4）
- 服务端把真字换成「备用码位」、字形换成真字轮廓 → 肉眼正常，`textContent` 是乱码：
  `中悢共产党悤史上` ← 读到 / `中国共产党历史上` ← 实际显示
- **映射每页随机**：`章节测验.html` vs `作业页面.html` 46 个重叠码位 **0 一致**（U+51C4 在测验页=误、作业页=书）→ 静态对照表不可行
- **实施方式**（与最初方案的差异）：
  1. **解码表**：`sha1(glyf[loca[gid] : loca[gid+1]])[:4] → BMP 码位`，全量 20902 字，零冲突；
     以**二进制紧凑存放（每条 6 字节）+ 运行时二分查找**，而非 JSON/Map。生成物 `src/extractors/cxsecret-table.ts`（163.3 KB base64）。
  2. **运行时**：自实现同步 SHA-1（只取首 4 字节）+ 最小 sfnt 解析（`head`/`maxp`/`loca`/`glyf`/`cmap`，支持 cmap 4/12 与 TTC）。
     不依赖 `atob`（自带 base64 解码），避免油猴沙箱差异。
  3. **按作用域解码**：只还原「自身或祖先类名含 `cxsecret`」的文本；作用域外的同码位文字是**正常内容**，无条件替换会改坏它。
     找不到类名标记时退化为整页还原。家族名不匹配时回退尝试全部内联字体，且要求命中数 ≥ 4，避免误伤普通内嵌字体。
  4. **接入点只有一处**：`rich-content.ts` 的文本抽取出口（文本节点 + `aria-label` 兜底路径）。
     题干、选项、答案、解析全部经此路径，下游（`answer-comparison`、导出器）拿到的已是明文。
- **为何没用第三方预生成表**：参考仓库 `TellMeYourWish/chaoxing_solution_of_font_confusion` **无任何 LICENSE**（默认保留所有权利），
  把其派生数据打包进仓库不合规，因此只把它当作**分析阶段的交叉验证素材**；生成器提供 `--font`（只用 OFL 字体，推荐）与 `--pkl`（便利重建）两条路径。
- **已验证的其他路线（均失败，勿重复尝试）**：
  - 本机全部字体（Noto/思源/微软雅黑/等线/黑体/宋体）字形字节哈希 → **0 命中**
  - 字形轮廓归一化点序列精确匹配 → **0 命中**（页面字体经 OTF→TTF 转换，点数不同，如「中」24 vs 16）
  - `post` 表 format 3.0（无字形名）、`cmap` 无真字码位 → 无从旁路获取

### RC-3 ℹ️ 无需大改 —— 多层 iframe

`frame-bridge.ts` 的 `requestRefresh()` 已递归到 depth 5，`FrameAgent` 逐层 `postMessage` 给 parent，`@match` 覆盖四层。唯一风险：每层都会加载 `@require docx`。

### RC-4 ℹ️ 待补材料 —— 答案提取

当前快照 `<!--新的答案开始-->…<!--新的答案结束-->` 为空，解码后 `正确答案/我的答案/解析` 全 0 → 属**答题前**视图。要验证答案提取，需另采一份**已提交**状态的快照。

---

## 4. 已完成改动清单（RC-1）

### 新增 `src/extractors/question-stem.ts`

三层逐级兜底，**不写死任何单一选择器**：

1. **已知类名候选表** `STEM_CONTAINER_SELECTORS`（新版→历史版本排序，命中即用）
   = 四个提取器原有选择器的并集：`[data-role="stem"]` / `.question-stem` / `.questionStem` / `.question-content` / `.qtContent` / `.mark_name` / **`.fontLabel`** / `.subject-title` / `.subject`
   → **以后改版只需往这个列表追加一个类名，不改任何流程代码**
2. **题型标记的父元素** —— 找 `.newZy_TItle` / `[class*="TestType"]` 取父元素。这是**结构特征**：题干与「【单选题】」总在同一元素内，容器改名也不影响
3. **兜底** —— 克隆容器后剔除题号（`i` / `.fl` / `[class*="index"|"number"]` 中的纯数字块）、`ul/ol` 选项列表、答案区、题型标记，取剩余内容

返回 `{ content, rawText }`：`content` 为剥离前缀后的题干；`rawText` 为剥离前原文，供 `parseLeadingNumber` 取题号 → **题号解析行为与改动前完全一致**。

### 接入的四个提取器

`timu-extractor.ts` / `mark-item-extractor.ts` / `question-li-extractor.ts` / `generic-exercise-extractor.ts` —— 统一改用该定位器，删除各自重复的选择器列表与多余导入。

### 测试与工程配置

- 新增 devDependency **`linkedom`**（仅测试用，**不影响构建产物**；`tsconfig.test-build.json` 原本刻意排除提取器，就是因为 Node 无 DOM —— 这是 RC-1 一直没被单测覆盖的原因）
- 新增 `tests/dom-setup.cjs`：把 linkedom 的 `Node`/`Element`/`HTMLElement`/`HTMLImageElement` 挂到全局
- 新增 `tests/question-stem.test.cjs`（7 例）、`tests/question-extraction.test.cjs`（6 例）：覆盖 2026 真实结构、题干类名被改、题型标记被改、兜底不混入选项、图片题干、空题干不产出题目、旧版 `.qtContent` 行为不变
- `tsconfig.test-build.json`：`include` 中提取器显式条目改为 `src/extractors/**/*.ts`

### 验证数据

| 项目 | 结果 |
|---|---|
| 真实快照 L4 层（`html/章节测验.html`） | `TiMuExtractor` **5 题**、`CompositeExtractor` **5 题**（修复前 `null`） |
| typeMeta / 选项 | 【单选题】【多选题】正确，A/B/C/D 正文正确 |
| 单测 | **14 → 27 例，全绿** |
| typecheck / ESLint / build | 0 错误 / 0 警告 / 223,463 → **227,687 字节** |
| 浏览器冒烟 | 失败点与基线完全相同 → 无回归 |

---

## 5. 已知但未处理的问题

0. **【本轮新发现】考试页面题干混入分数残片** —— `html/考试页面.html`（56 题，`questionLi` 结构，无字体反爬）提取出的题干是
   `, 1.0 分)\n算法分析通常采用( )?`。原因：题干容器候选里的 `.mark_name` 命中的是**整个 `<h3>`**，它把
   `<span class="colorShallow">(单选题, 1.0 分)</span>` 一并包了进去；`stripQuestionPrefix` 只能剥掉
   **结尾**的 `（…分）`，位于中间的这段剥不掉；题型标签正则要求 `单选题` 后紧跟括号，而实际是 `,`，故也失配。
   修法（未实施）：让题干定位优先取 `.mark_name` 内除去 `.colorShallow` 后的内容，或把分数括号的剥离范围从「结尾」放宽到「题号之后、题型标记之后」。
1. **`tools/browser-smoke.mjs` 基线即失败** —— `Copy formatter failed`。断言在「提取成功」检查之后，故提取路径本身正常；根因是 headless 下 `navigator.clipboard` 不可被 fixture 的 `Object.defineProperty` 覆盖，`window.__copied` 始终为空。修好可恢复仓库自带验证闭环
2. **`parseLeadingNumber` 不认 `<i class="fl">1</i>`** —— 它要求 `/^\s*(\d+)\s*[.、．]/`（带分隔符），真实页面题号无分隔符 → `question.number` 为 `undefined`（导出时按序号兜底）。属既有行为，本次刻意保持
3. **`chapter-locator.ts` 两处 P0（尚未动）**：
   - `activate()`（:47-60）对**容器 div** 调 `.click()`，但 `onclick="getTeacherAjax(...)"` **只挂在子节点 `.posCatalog_name`（span）上** → 章节切换实际不生效，每章空等 14 秒超时（旧版脚本同缺陷）
   - `:18-25` 选 `.posCatalog_select` 命中 **200 个**直接返回 = 28 个 `.firstLayer` 分组标题 + **172 个**真实章节点 → 28 个伪章节必然失败，点「全选」放大失败率
   - 次要：遍历无「有无题目」预判（172 章最坏 >40min 且无取消）、章节弹窗无编号/分组/搜索、失败清单不落盘、`#iframe` 内 `.ans-job-icon` 测验任务卡入口未打通

---

## 6. 可复用的分析脚本（`.workbuddy/`，未纳入 git）

| 脚本 | 用途 |
|---|---|
| `split_layers.py` | 切分四层嵌套文档，定位每层行边界 |
| `extract_font.py` | 提取页面内联 base64 字体为 TTF（产出 `font_cache/*.ttf`） |
| `parse_font.py` | 解析子集字体 cmap / glyf 结构 |
| `probe_glyfdata.py` | 探测 `fontTools` 的 `Glyph.data` 与原始 loca 切片的字节关系 |
| `verify_convention.py` | **钉死运行时字节约定**：`sha1(raw glyf[loca[i]:loca[i+1]])` 命中参考表 83/83、154/154 |
| `raster_match.py` | 栅格形状匹配反推真实汉字（独立验证方法，产出 `font_map.json`） |
| `compare_maps.py` | 比对两页映射是否一致（证明随机性，产出 `maps.json`） |
| `validate_table.py` | 用公开哈希表解码并与栅格结果交叉验证 |
| `decode_page.py` | 整页解码，产出 `章节测验.decoded.html` |
| `build_table.py` | 实测解码表体积与键冲突（`sha1[:3..8]` / `md5`） |
| `test_source_font.py` | 检查本机字体能否复现参考表哈希（结论：**不能**，0 命中） |
| `test_outline.py` | 检查字形轮廓点序列能否跨字体精确匹配（结论：**不能**，点数不同） |
| `probe_scope.py` / `probe_three.py` | 普查三份快照的混淆作用域标记与题目结构 |
| `verify-js.cjs` | 用编译产物解码真实字体，与栅格真值逐条比对（**83/83、154/154**）+ SHA-1 vs node crypto |
| `verify-real-decoded.cjs` | **端到端**：三份真实快照 → 解码器 + `CompositeExtractor`，打印题干/选项并检测生僻字残留 |
| `ref_glyfSearch.py` / `ref_README.md` | 参考实现的算法与说明（仅作对照，其数据未进入仓库） |

产物：`章节测验.decoded.html`（可读全文）、`章节测验.l4.html`、`maps.json`（栅格真值，**是验证解码正确性的裁判**）、`font-cache/`

---

## 7. 下一步建议（按优先级）

1. **考试页面题干残片**（第 5 节第 0 条）—— 56 题的题干都带 `, 1.0 分)` 残片，影响导出可读性；原因已定位到 `.mark_name` 命中整个 `<h3>`，改动小，`html/考试页面.html` 可直接回归
2. **`chapter-locator` 两处 P0** —— 章节切换点击目标 + 28 个分组标题被当章节。改动小，可直接补单测（需重新采集章节树快照，或按已记录的 DOM 结构写 fixture）
3. **`browser-smoke` 修复** —— 让仓库自带验证闭环恢复（headless 下 `navigator.clipboard` 覆盖失败）
4. **补「已提交」状态快照** —— 用于验证答案/解析提取（RC-4）
5. **`parseLeadingNumber` 兼容无分隔符题号**（`<i class="fl">1</i>`）
6. **可选：解码表瘦身** —— 当前全量 20902 字 / 163 KB；若在意体积可只收常用 7000 字（约 70 KB），代价是生僻字不解码

---

## 8. 给下次会话的开场提示

> 读 `docs/session-handoff-2026-09-18.md` 接手。RC-1（题干分层定位）与 RC-2（`font-cxsecret` 字体反爬解码）均已完成并构建：单测 38 例全绿，三份真实快照端到端通过（章节测验 5 题 / 作业页面 44 题 / 考试页面 56 题），dist 420,573 字节。
> 当前待办：**考试页面题干混入 `, 1.0 分)` 残片**、`chapter-locator` 两处 P0、`browser-smoke` 修复、补「已提交」快照验证答案提取。
> 注意本机 `npm` 不可用（用 node 直接调 `node_modules` 下的 CLI），Bash 缺 `mkdir`/`cp`/`head` 等命令（用 Python 做文件操作），PowerShell 需写文件再读。
