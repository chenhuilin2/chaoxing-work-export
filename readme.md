# Chaoxing Work Export — 学习通作业（考试）一键导出题目，快速生成一份 Word 试卷

一个使用 TypeScript 重构的超星学习通用户脚本，用于从已加载的作业、考试与章节练习页面中提取题目，并导出为 Word、TXT 或 Markdown。

这个用户脚本可以让你一键提取学习通作业题目并导出 **Word 试卷**（.docx），无需手动复制粘贴，就可以保存试题，方便后续进行期末复习。

支持的导出格式：**Word 试卷** / TXT / Markdown。
功能：支持答案/错题收集、图文混排、题库导入格式

本分支以“可维护、可测试、可扩展”为目标，将原先集中在单个 JavaScript 文件中的页面识别、状态管理、UI、历史记录和导出逻辑拆分为独立模块。构建后仍产出一个可直接安装到脚本猫、Tampermonkey 等管理器中的 `.user.js` 文件。

## 功能

- 识别 `.mark_item`、`.questionLi`、`.TiMu.newTiMu` 等学习通页面结构。
- 通用练习页兜底解析，覆盖部分多次提交在线练习页面。
- 单选、多选、填空、判断、简答题统一领域模型。
- 保留题干、选项、答案及解析中的文字、图片、换行、粗体、斜体、上下标。
- 提取正确答案、用户答案、错题状态和答案解析。
- 支持同源 iframe 递归提取与跨子域 iframe `postMessage` 回传。
- 支持选择多个章节，自动遍历并合并题目。
- 导出 Word、TXT、Markdown；支持参考答案、错题汇总、答案解析和题型内随机排序。
- 支持学习通题库智能导入 Word 格式。
- 多章节结果可合并导出，也可按章节拆分为多个文件。
- 下载历史、历史记录恢复与重新下载。
- 浅色、深色、跟随系统主题；快捷键、面板拖动和位置记忆。
- 自动迁移旧版 `xxt_settings`、`xxt_history` 等本地数据。

## 安装构建产物

直接安装：

```text
dist/chaoxing-work-export.user.js
```

脚本依赖固定版本的 `docx@8.5.0`，由 userscript 的 `@require` 在运行时加载。


### scriptcat 安装

1. 确保已安装脚本管理器（[脚本猫](https://docs.scriptcat.org/)）
2. 打开安装页面 ([超星学习通作业/考试一键提取导出word文档](https://scriptcat.org/zh-CN/script-show-page/6713))


## 开发环境

要求 Node.js 20 或更高版本。

```bash
npm install
npm run typecheck
npm run test
npm run build
```

完整校验：

```bash
npm run validate
```

构建结果：

```text
dist/chaoxing-work-export.user.js
dist/chaoxing-work-export.meta.js
```

项目使用 TypeScript 的 AMD 输出能力和仓库内置的小型模块运行时完成零额外打包器构建。源码仍按 ES 模块组织，最终产物为单文件 JavaScript userscript。

## 常用命令

| 命令 | 作用 |
| --- | --- |
| `npm run typecheck` | 严格 TypeScript 类型检查 |
| `npm run test` | 执行 Node.js 内置测试 |
| `npm run test:browser` | 使用本机 Chromium/Chrome 执行可选浏览器冒烟测试 |
| `npm run lint` | ESLint 严格规则检查 |
| `npm run format` | 使用 Prettier 格式化工程 |
| `npm run build` | 生成可安装的 `.user.js` 与更新元数据 |
| `npm run validate` | 类型检查、单元测试和构建 |

## 目录结构

```text
src/
├── application/      # 用例编排：提取、章节遍历、控制器
├── domain/           # Question、ExtractionResult、Settings 等稳定模型
├── exporters/        # Word/TXT/Markdown 与导出策略
├── extractors/       # 各页面结构适配器与富文本解析
├── infrastructure/   # localStorage、iframe、下载、剪贴板、日志
├── types/            # 外部运行时声明
├── ui/               # Shadow DOM 面板和样式
├── utils/            # 无业务状态的通用函数
├── app-config.ts
├── main.ts
└── userscript.meta.txt
```

详细设计见 [ARCHITECTURE.md](ARCHITECTURE.md)，扩展解析器的方法见 [docs/adding-extractor.md](docs/adding-extractor.md)。

## 快捷键

- `Ctrl + Shift + E`：打开面板并提取当前页面。
- `Ctrl + Shift + H`：隐藏或恢复整个浮窗。

快捷键可在设置中修改。输入框、文本域和可编辑区域获得焦点时，脚本不会拦截快捷键。

## 隐私与使用边界

脚本在浏览器本地解析当前用户已经打开并加载的页面，不向项目维护者上传题目、答案或课程信息。Word 图片会在导出时由浏览器请求并嵌入文件；受跨域策略限制的图片会使用占位提示代替。

请仅在你有权访问的课程和页面中使用，并遵守学校、课程和学习通平台的规则。

## 贡献

提交问题时，请提供页面类型、脚本版本、可脱敏的 DOM 结构和复现步骤，不要公开账号、Cookie、课程邀请码或未授权题库内容。参见 [CONTRIBUTING.md](CONTRIBUTING.md)。


## 欢迎提出问题、需求

如有问题以及需求，欢迎在 issue 页面中提出，我会及时修复更新。

谢谢以下开源贡献者～
<a href=" ">  
 <img src="https://contrib.rocks/image?repo=chenhuilin2/chaoxing-work-parser" width="100"/>
</a >


## License

GPL-3.0-only。原项目作者及贡献者署名保留在 userscript 元数据中。
