# 重构说明

## 重构前的主要问题

- 页面样式、DOM 选择器、提取流程、Word 构造、历史记录和事件绑定集中在一个大型 userscript 中。
- 数据结构依赖中文题型分组和大量隐式字段，跨函数修改风险高。
- 页面适配逻辑与导出逻辑耦合，增加新页面结构时容易影响所有格式。
- iframe、章节遍历依赖共享全局变量和时序回调，难以独立验证。
- localStorage key 与数据结构没有版本边界。
- 缺少类型检查、测试、lint、格式化和 CI。

## 重构后的对应关系

| 原职责 | 新位置 |
| --- | --- |
| 题型与题目对象 | `src/domain/question.ts` |
| `.mark_item` 解析 | `src/extractors/mark-item-extractor.ts` |
| `.questionLi` 解析 | `src/extractors/question-li-extractor.ts` |
| `.TiMu.newTiMu` 解析 | `src/extractors/timu-extractor.ts` |
| 未知练习页 | `src/extractors/generic-exercise-extractor.ts` |
| 富文本、图片、前缀清洗 | `src/extractors/rich-content.ts` |
| iframe 通信 | `src/infrastructure/frame-bridge.ts` |
| 多章节遍历 | `src/application/chapter-extraction-service.ts` |
| TXT / Markdown | `src/exporters/*-formatter.ts` |
| Word | `src/exporters/word-exporter.ts` |
| localStorage | `src/infrastructure/*-repository.ts` |
| 浮窗 | `src/ui/panel-view.ts` 与 `styles.ts` |
| 业务流程 | `src/application/app-controller.ts` |
| userscript 入口 | `src/main.ts` |

## 扩展成本

新增页面结构通常只需要：

1. 添加一个 extractor。
2. 复用 `QuestionFactory` 输出统一模型。
3. 在 `CompositeExtractor` 注册。
4. 增加 fixture 或纯逻辑测试。

Word、TXT、Markdown、历史记录与 UI 不需要同步修改。

## 兼容边界

- 页面必须已由当前用户正常加载；脚本不会绕过登录、权限、验证码或课程访问控制。
- 跨域图片仍受浏览器 CORS 策略影响，失败时写入占位文本。
- 学习通可能灰度发布新的 DOM；通用适配器只能兜底具有可识别语义标记的页面。
- 多章节提取依赖页面存在可点击章节列表，且每章内容可在超时时间内完成加载。
- Word 功能依赖元数据中固定版本的 `docx` UMD 包。
