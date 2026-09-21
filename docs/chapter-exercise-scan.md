# 学生学习页面 / 章节测验 提取问题 — 目录扫描与诊断报告

> **状态更新（2026-09-17 23:30）：本文第二节「章节测验.html 是空壳」及第四节的 B 组待确认项（B1–B4）已被
> [`chapter-quiz-diagnosis.md`](./chapter-quiz-diagnosis.md) 取代** —— 用户已重新采集到「四层嵌套 iframe 合并导出」的
> 有效快照，根因已确诊（题干选择器落空 + `font-cxsecret` 字体反爬）。
> 本文仍作为 **A 组 P0（`chapter-locator` 两处）** 与首次目录扫描结果的证据来源保留。

扫描时间：2026-09-17 22:14 ~ 22:40
项目：`Chaoxing-Work-Export`（学习通题目提取导出 userscript，TypeScript 六层架构，package.json v3.0.0，CHANGELOG 最新 3.0.4）
本次目标：学生学习页面（章节练习 / 章节测验）的题目提取优化

---

## 一、当前工作区扫描结果

> 注意：扫描期间 `html/` 目录发生过一次重组（`1.html`、`其他页面.html`、`ele.html`、`试题网页/` 已被清理），以下为**复核后的当前真实状态**。

### 1.1 `html/` 快照清单

| 文件 | 大小 | 修改时间 | 内容判定 |
| --- | --- | --- | --- |
| `html/做作业.html` | 272.6 KB | 06-26 | **有效题目快照**，93 处 `TiMu`、44 处 `newTiMu`、89 处 `Zy_TItle` |
| `html/章节测验.html` | 152.8 KB | 09-17 22:25 | **空壳，无任何题目内容**（详见第二节） |
| `html/头歌试题页面/1.html` | 31.6 KB | 06-25 | 头歌 EduCoder 平台页，**非学习通，与本次无关** |
| `html/头歌试题页面/image.png` | 241.5 KB | 06-25 | 截图 |

### 1.2 源码模块规模（Top 10）

| 文件 | 大小 | 职责 |
| --- | --- | --- |
| `src/ui/panel-view.ts` | 34.3 KB | Shadow DOM 面板 |
| `src/exporters/word-exporter.ts` | 22.1 KB | Word 生成 |
| `src/ui/styles.ts` | 14.6 KB | 面板样式 |
| `src/application/app-controller.ts` | 11.6 KB | 总控 |
| `src/infrastructure/history-repository.ts` | 9.7 KB | 历史记录 |
| `src/exporters/legacy-bridge.ts` | 9.2 KB | 旧行为对齐 |
| `src/extractors/rich-content.ts` | 8.8 KB | 富文本规范化 |
| `src/infrastructure/frame-bridge.ts` | 5.9 KB | iframe postMessage |
| `src/exporters/export-service.ts` | 5.7 KB | 导出入口 |
| `src/application/extraction-service.ts` | 5.3 KB | 提取编排 |

其他：`tests/` 4 个用例文件、`tools/` 4 个脚本、`dist/` 已有 09-11 构建产物（218.2 KB）、`old-version/` 旧脚本归档（135.9 KB）。

---

## 二、决定性发现：`html/章节测验.html` 是一个空壳，不含任何题目

用脚本逐项核验（剥离 `<script>`/`<style>`/注释后统计）：

| 检测项 | 命中数 | 说明 |
| --- | --- | --- |
| `TiMu` / `newTiMu` | **0 / 0** | 无题目容器 |
| `Zy_TItle` / `.qtContent` | **0 / 0** | 无题干结构 |
| `mark_item` / `questionLi` | **0 / 0** | 无其他受支持结构 |
| `input[type=radio]` / `checkbox` | **0 / 0** | 无选项控件 |
| `正确答案` / `我的答案` / `答案解析` | **0 / 0** | 无答案内容 |
| `单选题` / `多选题` / `判断题` | **0 / 0** | 无题型标签 |
| `章节测验` / `章节练习` / `测验` / `作业` / `考试` | **0** | 页面上没有任何任务文案 |
| `.posCatalog_select` | **0** | 左侧章节目录**完全没有内容** |
| `id="iframe"` | **0** | **章节内容 iframe 元素不存在** |
| 可见中文文本 | **271 字符 / 41 块** | 全部是外壳 UI 文案 |

可见文本全文基本是：`学生学习页面`、`提示`、`预览网页`、`返回课程`、`章节详情`、`目录`、`讨论`、`笔记`、`标注`、`新建话题`、`暂无笔记, 新建一个吧～`、`请输入验证码：`、`看不清`、`举报`、`打开学习通` 等。

关键 DOM 原文（`章节测验.html` L1701-1703）：

```html
<div class="onetoone posCatalog " id="coursetree" style="overflow-y: auto;overflow-x: hidden;">
</div>
```

`#coursetree` 是**空的**；页面里也没有 `<iframe id="iframe">`（只有 `reportFrame`、`popFrameId`、`hidePopFrameId`、`voteIframe` 四个 `src=""` 的弹窗占位 iframe）。文件中出现的 `#iframe`、`.ans-job-icon` 全部只存在于**内联 JS 字符串**里（如 `$("#iframe").contents().find(".ans-job-icon")`），并没有对应的真实 DOM。

### 结论

**这个快照无法复现线上问题。** 学习通学生学习页面的章节目录与章节内容是 AJAX/JS 动态注入的，浏览器「另存为 / 查看源代码」得到的是注入前的静态骨架。脚本在此文件上返回「当前页面未识别到题目」是**正确行为**——页面里确实没有题目。

因此，目前**无法判断**线上章节测验页失败的真实原因，需要重新采集 DOM（方法见第五节）。

---

## 三、受支持结构 vs 快照结构对照

当前 `CompositeExtractor` 按置信度注册 4 个提取器：

| 提取器 | 置信度 | 触发选择器 |
| --- | --- | --- |
| `TiMuExtractor` | 110 | `.TiMu.newTiMu`、`#ZyBottom .TiMu` |
| `MarkItemExtractor` | 100 | `.mark_item` |
| `QuestionLiExtractor` | 90 | `.questionLi` |
| `GenericExerciseExtractor` | 40 | `[data-question-id]`、`.question-item`、`.TiMu` 等 |

### 3.1 `做作业.html` 的真实结构（与提取器对照）

```
#ZyBottom.ans-cc
├── <h3 class="newTestType">一. 单选题（共23题）</h3>      ← 题型分组标题
├── .singleQuesId[data=214729069]
│   └── .TiMu.newTiMu[data="0"]
│       ├── .Zy_TItle.clearfix
│       │   ├── <i class="fl">1</i>                        ← 题号
│       │   └── <div class="clearfix font-cxsecret fontLabel">
│       │         <span class="newZy_TItle">【单选题】</span>题干正文…
│       └── ul.Zy_ulTop > li[role=radio]                   ← 选项
│             └── label.fl.before > span.num_option[data="A"]
│             └── a.fl.after > 选项文本
└── input[type=hidden][name=answer214729069]               ← 作答值
```

### 3.2 对照中发现的两处结构落差（需在拿到真实快照后一并验证）

1. **`TiMuExtractor` 的题干选择器可能落空**
   `src/extractors/timu-extractor.ts:52`：
   ```ts
   const stemElement = title.querySelector('.qtContent, .question-content, .mark_name');
   ```
   而 `做作业.html` 中 `.qtContent` / `.question-content` / `.mark_name` **均为 0 命中**，题干文本直接放在 `.Zy_TItle > .fontLabel` 里。此时 `stemElement === null` → `extractRichContent(null)` 返回 `[]`（`rich-content.ts:128`）→ **题干为空**，但题目仍会被创建（题型、选项能取到），因此不会报「未识别到题目」，而是**静默丢题干**。
   `.fontLabel` 未出现在任何选择器列表中，属于潜在缺口。

2. **`TiMuExtractor` 未处理无 `.aiArea` 的分组**
   `timu-extractor.ts:21` 先找 `#ZyBottom .aiArea`，为空时退化为遍历全部 `.TiMu.newTiMu` 并传 `inheritedType = null`，此时题型只能靠容器内 `.newZy_TItle` 文本判断。`做作业.html` 里 `aiArea` 为 0 命中，走的正是该退化路径——能工作，但分组信息（`<h3 class="newTestType">`）未被利用，题型分组顺序完全依赖单题标签。

> 以上两点是基于**有效快照**推断出的确定性缺口。它们与「章节测验提取失败」可能是同一批问题，也可能无关，需真实快照确认。

---

## 四、P0 缺陷清单

### A 组：已确诊，可直接修（依据来自已清理的 `html/1.html` 学生学习页面快照）

> 该快照（10.4 MB）在扫描期间被清理，但已采集到完整结构证据，结论有效。

学习页面 `#coursetree` 共 **200 个 `.posCatalog_select`**，实为 28 个 `.firstLayer` 分组标题 + 172 个 `id="cur<chapterId>"` 真实章节点；`onclick="getTeacherAjax(courseId,clazzid,chapterId)"` **只挂在 `.posCatalog_name`（span）上，容器 div 上 0 个**。

| # | 位置 | 问题 | 影响 |
| --- | --- | --- | --- |
| A1 | `src/application/chapter-locator.ts:47-60` | `activate()` 命中容器 div，`a, button, [role="button"]` 全无匹配 → 对 div 调 `.click()`。事件目标是 div，**子节点 span 的内联 `onclick` 不会触发** | 章节切换完全不生效，每章空等 14 s 超时 |
| A2 | `src/application/chapter-locator.ts:18-25, 62-81` | 选择器首项 `.posCatalog_select` 命中 200 个即返回，不再尝试 `.posCatalog_name`（172 个）；`titleOf()` 的 `querySelector('[title]')` 还会取到分组标题的 `title` | 28 个分组标题被当成章节，列表 200 项，点「全选」放大失败率 |

> 旧版 `old-version/old-chaoxing-work-export.user.js:2829` 同样是 `chapterLinks[i].click()`，A1 缺陷两者共有。

### B 组：待真实快照才能确诊

| # | 现象 | 待确认 |
| --- | --- | --- |
| B1 | 章节测验页面提取失败，报「当前页面未识别到题目」 | 题目到底渲染在哪：顶层文档 / `#iframe`（`knowledge/cards`）/ iframe 内再嵌套的答题页；实际类名是什么 |
| B2 | 章节测验是否为跨子域 iframe | 若 `contentDocument` 抛错，需走 `FrameBridge` postMessage 通道 |
| B3 | 题干选择器 `.qtContent` 落空（见 3.2） | 线上章节测验题干的实际容器类名 |
| B4 | 选项/答案结构 | `.Zy_ulTop`、`.answerBg`、正确答案与我的答案的实际容器 |

---

## 五、下一步：请重新采集可用的页面快照

需要两份材料：

1. **顶层文档**：`章节测验` 页面（学生学习页面 / 答题页均可）**渲染完成后**的 DOM。
2. **iframe 文档**：若题目在 iframe 里，需要该 iframe 内部的 DOM。

采集方式（浏览器 DevTools Console，需在页面题目已经显示出来之后运行）：

```js
// 把主文档 + 所有同源 iframe 的 DOM 合并导出为一个 .html 文件
(() => {
  const parts = [];
  const seen = new Set();
  const walk = (doc, depth, label) => {
    if (!doc || seen.has(doc) || depth > 6) return;
    seen.add(doc);
    let url = '';
    try { url = doc.location.href; } catch {}
    parts.push(`\n<!-- ===== ${label} | depth=${depth} | url=${url} ===== -->\n`);
    parts.push(doc.documentElement.outerHTML);
    doc.querySelectorAll('iframe').forEach((f, i) => {
      try { walk(f.contentDocument, depth + 1, `${label} > iframe[${i}]`); } catch {}
    });
  };
  walk(document, 0, 'TOP');
  const text = parts.join('');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([text], { type: 'text/html;charset=utf-8' }));
  a.download = 'dom-capture.html';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 8000);
  console.log('已导出，约', (text.length / 1024).toFixed(0), 'KB');
})();
```

注意事项：

- **必须在题目渲染出来之后运行**，否则拿到的还是空壳。
- 若某个 iframe 是跨子域的，`contentDocument` 会抛错并被跳过；此时请单独打开该 iframe 的 URL，在新标签页里再跑一遍上面的片段。
- 采集到的片段内含课程名、账号相关 URL，请自行脱敏后再放入仓库。
- 保存到 `html/` 下即可（该目录已在 `.gitignore` 中）。

---

## 六、校验闭环

```bash
npm run validate   # typecheck + 单测(14 用例) + 构建，改动 src/ 后必跑
npm run lint       # ESLint 0 警告
npm run test:browser  # 可选浏览器冒烟测试
```

协作约束（`AGENTS.md`）：手术式修改、只改必需文件、注释与文案统一中文、`extractors/rich-content.ts` 与 `exporters/legacy-bridge.ts` 需同步修改。
