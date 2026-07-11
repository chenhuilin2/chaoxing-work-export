## `v1.10.0` - 最新版本

### iframe 内题目提取
- 移除 `@noframes` 指令，脚本现在也可以在 iframe 内运行
- 在 iframe 中运行时，自动提取题目并通过 `postMessage` 回传给父窗口，不在 iframe 内创建 UI
- 父窗口通过 `message` 事件监听接收 iframe 数据，`extract()` 函数自动合并 iframe 结果
- 延迟重试机制（1.5s / 3s），应对 iframe 内容动态加载场景
- 同源 iframe 仍通过 `extractFromIframesRecursive` 直接递归查找

### 章节测验/考试页面提取
- 新增对 `.TiMu.newTiMu` 结构的提取，覆盖 `#ZyBottom > .aiArea` 章节测验与考试页面
- 新增 `extractFromTiMuRoot` 函数，识别 `.newTestType` 题型标题、`.Zy_TItle .qtContent` 题干、`.Zy_ulTop.qtDetail > li` 选项
- 支持从 `.newAnswerBx` 中提取「我的答案」「正确答案」及对错状态
- `stripQuestionPrefix` 支持剥离 `【单选题】` 等方头括号题型标签


## `v1.9.0`

### 富文本提取（图文混排）
- 新增完整的富文本提取基础设施，按 DOM 顺序提取题目内容，保留文字、图片和换行的原始顺序
- 新增 `extractRichContent` 函数，遍历 DOM 节点树提取文本、图片（`<img>`）和换行（`<br>`、块级标签）
- 新增 `RICH_BLOCK_TAGS` 常量，定义块级标签集合，自动在块级元素后追加换行
- 新增 `resolveImageUrl` 函数，多属性 fallback 解析图片 URL（`src`、`data-src`、`data-original`、`origin-src`、`fileid`）
- 新增 `normalizeRichContent` 规范化函数，合并相邻文本节点、去首尾空行、压缩空白
- 新增 `stripOptionPrefix` 剥离选项字母前缀（A~H），`stripQuestionPrefix` 剥离题干题号和题型标签，`stripAnswerLabel` 剥离答案标签
- 新增 `extractCorrectAnswerContent` 以富文本格式提取正确答案，保留答案中的图片和公式
- 新增 `extractOptionFromLegacyLi` 和 `extractOptionFromAnswerBg`，分别从旧版和新版页面结构提取选项富文本
- 新增 `questionContent`、`optionContent`、`answerContent` 三个适配器，统一兼容新旧数据格式

### 格式化输出改造
- TXT 格式化：题干和选项使用 `formatRichForText` 输出，图片以 `[图片: url]` 格式保留
- MD 格式化：题干和选项使用 `formatRichForMD` 输出，图片以 `![alt](url)` 格式保留
- 修复 `escapeMarkdownUrl` 括号编码问题，防止 URL 中的 `()` 破坏 Markdown 图片语法

### Word 导出改造
- 新增 `fetchImageAsset` 函数，带 10 秒超时的 AbortController，避免图片加载卡死
- 新增 `buildImageRun` 函数，将图片 URL 转为 docx `ImageRun`
- 新增 `buildRichRuns` 函数，将富文本内容数组转为 Word `TextRun` 数组，图文混排嵌入文档
- 题干、选项、答案部分均使用 `buildRichRuns` 渲染，支持内联图片和换行
- 题库导入格式同步适配：选项使用富文本渲染，题干图片独立渲染

### 代码清理
- 移除不再使用的 `cleanHtml` 函数
- 移除不再使用的 `extractImagesFromElement` 函数
- 移除不再使用的 `formatImagesForText`、`formatImagesForMD`、`normalizeStem`、`normalizeStemForBank` 函数
- 移除不再使用的 `fetchImageAsBase64` 函数（由 `fetchImageAsset` 替代）


## `v1.8.0`

### 题库导入格式
- Word 试卷导出下新增"题库导入格式"复选框，生成学习通智能导入兼容的 Word 文档
- 格式遵循学习通智能导入规范：`题号.【题型】题干` → 选项 → `答案：XXX`
- 多选题答案使用中文逗号分隔，判断题答案统一为"对"/"错"
- 勾选后自动禁用"打乱题目顺序"、"附加答案"、"附加错题"（灰显不隐藏，维持 UI 高度稳定）
- 题干括号归一化：宽括号 `(   )` 自动转为 `（ ）`

### Word 试卷附加错题
- Word 试卷导出模式下，"附加错题"复选框不再隐藏，用户可自由勾选
- 勾选后，试卷末尾新增错题汇总页（分页符分隔），包含题干、选项、我的答案（红色）和正确答案（绿色）
- 错题按题型分组展示，保留原题号，便于对照定位

### 错题汇总序号优化
- TXT、MD、Word 三种导出格式的错题汇总，序号统一改为保留原题号（全局题号），不再使用从 1 开始的独立序号，便于用户在试卷中快速定位错题


## `v1.7.1`

### 一键隐藏/显示快捷键
- 新增 `Ctrl+Shift+H` 快捷键，一键隐藏或显示浮窗面板和悬浮按钮
- 隐藏状态下再次按下快捷键即可恢复，无需刷新页面
- 支持在设置弹窗中自定义快捷键


## `v1.7`

### 新页面结构支持
- 新增对"做作业"页面的题目提取支持（原仅支持"查看答案"页面）
- 自动识别页面 DOM 结构：无 `.mark_item` 时，通过 `.questionLi` 的 `typeName` 属性识别题型
- 题干从 `h3.mark_name` 提取，选项从 `.answerBg` 容器提取（`.num_option` 取字母、`.answer_p` 取文本）
- 做作业页面不显示正确答案，提取结果中答案字段为空，不影响导出

### UI 交互优化
- Word 试卷调整为默认输出格式（排在第一位），用户无需手动切换
- "附加答案"、"附加错题"、"打乱题目顺序"三个复选框移至下载/复制按钮上方，形成"选格式 → 勾选项 → 点下载"的自然操作顺序
- 修复设置弹窗和历史记录弹窗内操作时浮窗被意外关闭的问题
- 勾选"附加答案"时，文件名预览区自动追加"（含答案）"后缀

### 代码优化
- 新增 `extractFromQuestionLi` 函数，与原有 `extract` 函数并行，按页面结构自动分支


## `v1.6`

### 下载历史记录
- 每次下载文件时会自动保存记录到 localStorage（最多保留 10 条）
- 记录包含：标题、日期、题目数、导出格式（TXT/MD/Word）、下载配置（含答案/含错题/打乱）
- 相同标题 + 相同格式的旧记录自动去重替换
- 面板头部新增历史记录按钮（时钟图标），点击弹出历史记录弹窗
- 弹窗中可查看所有历史下载，点击任意条目恢复当时的提取数据和格式设置
- 支持单条删除，点击遮罩层或 Esc 关闭弹窗

### Word 导出附加答案
- Word 试卷导出模式下，"附加答案"复选框不再隐藏，用户可自由勾选
- 勾选后，试卷末尾新增一页（分页符），按题型分组列出所有题目的参考答案
- 完成状态提示显示"Word 试卷（含答案）已下载"

### 选项排列优化
- 选择题选项智能排版：任意选项超过 25 字时自动改为单列竖排，避免选项拥挤

### 代码优化
- 修复 `Paragraph is not defined` 错误，将 `buildImageParagraphs` 移入 `generateWordBlob` 作用域
- 新增 `PageBreak` 导入，支持 Word 文档分页


## `v1.5`

### 图片保留
- 提取题目时同步提取题干预的图片，支持 `<img>` 标签和背景图片（`answerBg` 类）
- **Word 导出**：图片通过 base64 编码嵌入文档（`ImageRun`），无需外部依赖，离线可用
- **TXT/MD 导出**：保留图片原始 URL，格式为 `[图片: url]`
- 图片获取失败时显示灰色占位文本，不影响正常导出流程


## `v1.4`

### 脚本设置
- 面板头部新增设置按钮（齿轮图标），点击弹出设置弹窗

#### 快捷键支持
- 默认绑定 `Ctrl+Shift+E`，按下后自动打开面板并触发提取
- 支持自定义快捷键，点击设置弹窗中的快捷键区域进入录制模式
- 录制时按下目标组合键即可绑定，支持 Ctrl/Shift/Alt 修饰键
- 5 秒超时自动取消录制

#### 暗色模式适配
- 支持三种模式：**自动（跟随系统）**、**浅色**、**深色**
- 默认为自动模式，实时监听系统主题变化
- 深色模式下所有 UI 元素（面板、按钮、弹窗、统计卡片等）均适配深色配色
- 设置通过 `localStorage` 持久化存储


## `v1.3`

### 题目随机排序
- 面板新增"打乱题目顺序"复选框
- 勾选后导出时使用 Fisher-Yates 洗牌算法对题目重新排序
- 同一题型的题目仍然保持在一起，仅在题型内部打乱顺序
- 答案和错题部分始终使用原始顺序，不受打乱影响
- 适用于 TXT、MD、Word 三种导出格式