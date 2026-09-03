# 架构设计

## 1. 设计目标

本次重构遵循以下约束：

1. 最终仍是脚本猫或 Tampermonkey 可直接安装的单文件 JavaScript。
2. DOM 结构变化不应迫使开发者修改 UI、导出器或持久化代码。
3. Word、TXT、Markdown 必须消费同一份规范化题目模型，避免三套格式各自解析 DOM。
4. iframe、章节切换和异步加载属于应用编排问题，不进入单题解析器。
5. 业务模型不依赖浏览器 UI，纯函数能够在 Node.js 中测试。
6. 新版能够读取旧版设置和下载历史，升级不应无故清空用户数据。

## 2. 分层与依赖方向

```text
main
  └─ application
       ├─ extractors
       ├─ exporters
       ├─ infrastructure
       ├─ ui
       └─ domain

extractors ────────> domain
exporters ─────────> domain
infrastructure ────> domain
ui ────────────────> domain
utils 不持有业务状态
```

关键规则：

- `domain` 不导入 UI、DOM 解析器或 localStorage。
- `extractors` 只负责把一个 `Document` 转换成规范化题目，不负责下载与界面更新。
- `exporters` 不访问原页面 DOM，只处理 `ExtractionResult`。
- `PanelView` 不决定提取或导出流程，只负责展示状态、收集输入并触发回调。
- `AppController` 是组合根之外唯一同时协调 UI、提取、导出和存储的模块。

## 3. 核心领域模型

所有页面适配器统一输出：

```ts
interface Question {
  id: string;
  number?: number;
  type: QuestionType;
  stem: RichContent;
  options: readonly QuestionOption[];
  correctAnswer: RichContent;
  userAnswer: RichContent;
  analysis: RichContent;
  isWrong: boolean;
  source: QuestionSource;
  chapterId?: string;
  chapterTitle?: string;
}
```

`RichContent` 是按 DOM 顺序排列的文本、图片和换行节点。文本节点可携带粗体、斜体、上标和下标信息。这样，提取阶段只关心“内容是什么”，各导出器分别决定如何呈现。

`ExtractionResult` 是一次不可变快照，包含标题、题目、题型顺序、统计信息、来源、提取时间和可选章节列表。下载历史保存的也是该快照，而不是 UI 内部状态。

## 4. 提取流程

```text
用户触发提取
    │
    ▼
ExtractionService
    ├─ 当前 Document
    ├─ 可访问的同源 iframe（递归，最大 5 层）
    └─ FrameBridge 最近一次跨子域 iframe 结果
           │
           ▼
CompositeExtractor
    ├─ TiMuExtractor            置信度 110
    ├─ MarkItemExtractor        置信度 100
    ├─ QuestionLiExtractor      置信度 90
    └─ GenericExerciseExtractor 置信度 40
           │
           ▼
QuestionFactory → ExtractionResult
```

`CompositeExtractor` 优先选择最专用、置信度最高且成功返回题目的适配器，题目数量仅作为同置信度下的次级条件。该策略避免通用选择器因匹配到父子重复节点而覆盖专用解析结果。

### 富文本规范化

`rich-content.ts` 负责：

- 递归遍历文本、图片、换行和块级元素。
- 合并样式相同的相邻文本。
- 去除首尾与重复换行。
- 解析懒加载图片属性与 CSS 背景图。
- 剥离题号、题型、选项字母和答案标签。
- 保留粗体、斜体、上标、下标。

### 题型与错题

题型优先读取明确的中文标签或属性，兼容旧版数字代码；缺少标签时，根据 radio、checkbox、textarea 和判断题选项进行推断。

错题状态由两类证据合并：

1. 页面显式错误标记。
2. 对可确定题型的用户答案与正确答案规范化比较。

简答题不做字符串相等判定，避免把等价表述误判为错误。

## 5. iframe 协议

脚本不使用 `@noframes`。在 iframe 中运行时只启动 `FrameAgent`，不创建浮窗。

`FrameAgent`：

- 页面加载后立即尝试提取。
- 使用 `MutationObserver` 监听异步加载。
- 对频繁变化进行防抖。
- 在 1 秒、3 秒后再次重试。
- 将结构化 `ExtractionResult` 通过带协议版本的 `postMessage` 发给父窗口。

顶层窗口的 `FrameBridge` 仅接受当前来源或 `chaoxing.com` 子域发送的、结构形状正确的消息。跨域 iframe 不需要父窗口读取其 DOM。

## 6. 多章节提取

`ChapterLocator` 只负责发现章节与触发点击；`ChapterExtractionService` 负责顺序执行：

1. 获取选中章节和当前活动章节。
2. 逐个激活章节。
3. 等待新 iframe 消息或题目指纹变化。
4. 将章节 ID、标题写入每道题。
5. 汇总成功章节，记录进度。
6. 在结束或异常时恢复原章节。

已经处于活动状态的章节直接读取当前结果，不强制等待“内容变化”，避免首章超时。

## 7. 导出流程

```text
ExtractionResult
    │
    ├─ ResultTransformer（可选题型内随机排序）
    │
    ├─ TextFormatter
    ├─ MarkdownFormatter
    └─ WordExporter
          ├─ A4 页面与统一样式
          ├─ 图文混排
          ├─ 参考答案页
          ├─ 错题汇总页
          ├─ 答案解析
          └─ 题库智能导入布局
```

`ExportService` 统一管理扩展名、MIME 类型、文件名清洗以及按章节拆分。下载由 `DownloadService` 负责，导出器不会直接创建 `<a>` 标签。

Word 图片由带超时和缓存的 `ImageRepository` 加载，按原比例限制在 420×260 像素内。单张图片失败不会中断整份文档，导出结果会报告失败数量。

## 8. UI 与状态

UI 挂载在开放的 Shadow DOM 中，避免学习通页面样式覆盖脚本，也避免脚本 CSS 污染原页面。

`PanelView` 管理：

- 面板、弹窗、状态和统计的 DOM 表现。
- 导出选项读取与联动。
- 历史列表和章节选择列表渲染。
- 主题、拖动和位置表现。

`AppController` 管理：

- 当前提取快照。
- 异步操作状态和错误处理。
- 设置、历史记录的读取与写入。
- 键盘快捷键。
- 提取、复制、导出和重新下载用例。

## 9. 持久化与迁移

新版使用版本化 key：

```text
chaoxing-work-export:settings:v3
chaoxing-work-export:history:v3
```

首次读取不到新版数据时，仓储会尝试迁移：

```text
xxt_settings
xxt_export_config
xxt_export_options
xxt_history
```

历史仓储最多保留 10 条。localStorage 空间不足时，从最旧记录开始缩减；即使存储不可用，当前页面提取与导出仍可使用。

## 10. 构建

TypeScript 源码编译为一个 AMD 文件。`tools/build.mjs` 注入一个小型同步 AMD 运行时，然后执行 `main` 模块，并在前面拼接 userscript 元数据。

该方案的特点：

- 不依赖 Rollup、Webpack 或 Vite 的运行时行为。
- 构建配置小，输出可审计。
- 开发代码仍保持 ES 模块边界。
- 最终只有一个 `.user.js`，符合用户脚本安装方式。

## 11. 质量保障

- TypeScript `strict`、`noUncheckedIndexedAccess`、`noUnusedLocals` 等严格选项。
- ESLint type-aware strict/stylistic 配置。
- Prettier 与 EditorConfig。
- Node.js 内置测试覆盖题型映射、答案规范化、富文本规范化、随机排序和文本导出。
- `node --check` 可校验最终 JavaScript 语法。
- GitHub Actions 在 push 与 pull request 时执行 lint、类型检查、测试和构建，并上传 userscript 产物；Prettier 配置用于开发者本地统一格式。
