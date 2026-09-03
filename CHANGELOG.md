# Changelog

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
