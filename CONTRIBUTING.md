# 贡献指南

## 开发流程

1. 从 `main` 创建主题分支。
2. 将页面差异限制在对应解析器中，避免扩大跨层依赖。
3. 为纯逻辑添加测试。
4. 运行 `npm run lint`、`npm run format:check` 和 `npm run validate`。
5. 提交 pull request，并说明页面类型、改动范围和回归结果。

## 提交问题

一个可处理的解析问题应包含：

- userscript 版本。
- 浏览器与脚本管理器版本。
- 页面类型，例如“查看答案”“做作业”“章节测验”“考试”“多次提交练习”。
- 预期结果与实际结果。
- 控制台错误。
- 经脱敏的最小 DOM 片段或截图。

不得提交 Cookie、账号密码、课程邀请码、真实学生信息或无权公开的完整题库。

## 代码原则

- TypeScript 严格模式下不得出现未解释的类型逃逸。
- DOM selector 应集中在 extractor 或 locator 中。
- 新导出格式必须消费领域模型，不得二次读取页面。
- 异步循环必须有超时、失败隔离或取消路径。
- localStorage 读写必须通过仓储与 `SafeStorage`。
- 用户可见错误使用中文；调试细节通过 `Logger` 输出。

## Commit 建议

推荐使用清晰的动词前缀：

```text
feat: support a new exercise DOM structure
fix: preserve images in fill-in answers
refactor: isolate chapter navigation
 test: cover answer normalization
 docs: explain extractor confidence
```
