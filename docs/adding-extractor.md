# 添加新的页面解析器

学习通页面结构变化时，优先添加一个新的 `QuestionExtractor`，不要在 UI 或导出器中加入页面选择器。

## 1. 建立适配器

在 `src/extractors/` 新建文件：

```ts
import type { Question } from '../domain/question';
import type { ExtractorContext, QuestionExtractor } from './contracts';

export class ExampleExtractor implements QuestionExtractor {
  readonly id = 'example';
  readonly confidence = 95;

  supports(root: Document): boolean {
    return root.querySelector('.example-question') !== null;
  }

  extract(context: ExtractorContext): readonly Question[] {
    // 使用 createQuestion、extractRichContent、parseOptions 等公共能力。
    return [];
  }
}
```

## 2. 置信度约定

- `100+`：页面结构唯一且字段定义明确。
- `80–99`：结构稳定，但与其他页面存在部分重叠。
- `50–79`：兼容适配器。
- `<50`：通用兜底，不应覆盖专用解析器。

## 3. 注册

在 `CompositeExtractor` 默认适配器数组中注册。顺序用于阅读，最终选择以置信度为主。

## 4. 复用公共函数

- `detectQuestionType` / `inferQuestionType`：题型。
- `extractRichContent`：图文混排。
- `parseOptions`：选项。
- `extractCorrectAnswer` / `extractUserAnswer` / `extractAnalysis`：答案与解析。
- `hasExplicitWrongMarker`：页面错误标记。
- `createQuestion`：统一 ID、默认值和错题判定。

不要在适配器中直接生成 TXT、Markdown 或 docx 节点。

## 5. 测试与问题报告

至少提供一份脱敏 DOM fixture，覆盖：

- 题型标签。
- 题干和选项。
- 正确答案、用户答案。
- 解析和错题标记。
- 图片或公式（页面存在时）。

运行：

```bash
npm run typecheck
npm run test
npm run build
node --check dist/chaoxing-work-export.user.js
```
