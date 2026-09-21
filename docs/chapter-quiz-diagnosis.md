# 章节测验 / 做作业页面 提取失败 — 完整诊断报告

诊断时间：2026-09-17 23:30 ~ 次日
材料：`html/章节测验.html`（427.5 KB，09-17 23:24 采集，**四层嵌套 iframe 合并导出**）
项目：`Chaoxing-Work-Export`（学习通题目提取导出 userscript，TypeScript 六层架构）

> 本文回答并取代 `chapter-exercise-scan.md` 第四节的 B 组（B1–B4）待确认项。
> A 组（`chapter-locator` 两处 P0）证据仍有效，见该文第四节。

> **实施状态（2026-09-18）**
> - **RC-1 已完成**：新增 `extractors/question-stem.ts` 分层定位并接入四个提取器（v3.0.5）。
> - **RC-2 已完成**：新增字体解码模块 `extractors/cxsecret-font.ts` + `cxsecret-decoder.ts` + 生成物 `cxsecret-table.ts`（v3.0.6）。
>   实施结论与本文原方案的三点差异：① 表改为**二进制紧凑存放 + 二分查找**（122.5 KB 裸数据 / 163.3 KB base64），而非 JSON；
>   ② 解码**按作用域生效**（元素自身或祖先类名含 `cxsecret`），不做无条件整页替换——作用域外的同码位文字是正常内容；
>   ③ 体积取**全量 20902 字**（零冲突）。验证：与栅格真值 83/83、154/154 一致；三份真实快照端到端通过。
> - RC-3 仍待真机确认；RC-4 仍缺「已提交」状态快照。

---

## 一、四层嵌套链路（本次材料的确切结构）

`html/章节测验.html` 由 4 个文档拼接而成（`<html` 出现 4 次，与原「1.html 三层」相比多一层）。
`s/<title>/` 定位到的层级与行号边界：

| 层 | 行范围 | `<title>` | 该层内的 iframe 出口 |
| --- | --- | --- | --- |
| L1 | 23 – 1187 | 视频 | `<iframe id="iframe" info="card" src="/mooc-ans/knowledge/cards?...&num=0">` |
| L2 | 1192 – 2042 | 章节测验 | `<iframe id="iframe" info="card" src="/mooc-ans/knowledge/cards?...&num=1">`，<br>`<iframe src="/ananas/modules/work/index.html?... jobid="work-33742e88…">` |
| L3 | 2051 – 2282 | （无） | `<iframe id="frame_content" name="frame_content" src="/mooc-ans/api/work?api=1&workId=33742e88…">` |
| L4 | 2287 – 6137 | **做作业** | 无（终端文档，**题目在这里**） |

链路：学生学习页面 → 知识卡片页 → work 模块壳 → **答题页（`/mooc-ans/api/work`）**。

题目全部位于 **L4**：`fontLabel` 5 处、`newZy_TItle` 5 处、`singleQuesId` 5 处 → **共 5 道题**（4 单选 + 1 多选，与 `<h3 class="newTestType">一. 单选题（共4题）</h3>`、`二. 多选题（共1题）` 吻合）。

### L4 单题真实结构

```html
<div class="singleQuesId" id="question405842140" data="405842140">
  <div class="TiMu newTiMu" data="0">
    <div class="Zy_TItle clearfix">
      <i class="fl">1</i>                                    <!-- 题号 -->
      <div class="clearfix font-cxsecret fontLabel">         <!-- 题干正文容器 -->
        <span class="newZy_TItle">【单选题】</span>题干正文…
      </div>
    </div>
    <div class="clearfix">
      <ul class="Zy_ulTop w-top fl">
        <li class="font-cxsecret before-after" onclick="addChoice(this);"
            qid="405842140" qtype="0" role="radio" aria-label="A 毛泽东选择">
          <label class="fl before">
            <span class="num_option choice405842140" name="answer405842140" data="A">A</span>
          </label>
          <a href="javascript:void(0);" class="fl after">毛泽东</a>   <!-- 选项正文 -->
          <div class="clear"></div>
        </li>
        …
      </ul>
      <input type="hidden" name="answer405842140" id="answer405842140" value="">
      <input type="hidden" id="answertype405842140" name="answertype405842140" value="0">
    </div>
    <!--新的答案开始-->
    <!--新的答案结束-->                                    <!-- 本快照为空 → 无答案 -->
  </div>
</div>
```

---

## 二、RC-1（决定性根因）：题干选择器不匹配 → 所有题目被静默丢弃

`src/extractors/timu-extractor.ts:52`：

```ts
const stemElement = title.querySelector('.qtContent, .question-content, .mark_name');
```

这三个类名在 L4 中 **0 命中**（题干文本直接放在 `.Zy_TItle > div.fontLabel`）。于是：

1. `stemElement === null`
2. `extractRichContent(null)` → `[]`
3. `src/extractors/question-factory.ts:28` → `if (isRichContentEmpty(draft.stem)) return null;` → **每道题都返回 null**
4. `src/extractors/composite-extractor.ts:39` → `if (questions.length > 0)` 把空结果提取器丢弃
5. 其余三个提取器 `supports()` 也不成立：`mark_item` / `questionLi` 均 **0 命中** → 无候选 → `return null`
6. 面板提示「当前页面未识别到题目」

**这就是「章节测验提取失败」的直接原因。** 注意与嵌套层级无关——`Zy_TItle`、`Zy_ulTop`、`newZy_TItle`、`i.fl` 现有选择器**全部能命中**，唯独题干这一处漏了。

题型解析其实是好的：`title.querySelector('.newZy_TItle')` 能取到 `【单选题】`；题号 `i.fl` 能取到 `1`；选项 `.Zy_ulTop > li` 也能命中。

> 附带结论：`html/做作业.html`（272.6 KB，同一页面类型）也是同样的 `.fontLabel` 结构，因此**该缺陷影响的不止章节测验，而是整类「做作业/答题」页**。

---

## 三、RC-2（必须解决）：`font-cxsecret` 字体反爬

即使补上 `.fontLabel`，抽出的文字仍是**乱码**，因为题干与选项都带 `font-cxsecret` 类。

### 3.1 机制

L4 的 `<style>` 内联：

```css
@font-face {
  font-family: 'font-cxsecret';
  src: url('data:application/font-ttf;charset=utf-8;base64,…');   /* 20,756 字节 TTF */
}
.font-cxsecret, .font-cxsecret p, .font-cxsecret div, … { font-family: 'font-cxsecret' !important; }
```

`.font-cxsecret` 在本页命中 **35 次**。

提取出的字体经解析（fontTools）：

| 属性 | 值 |
| --- | --- |
| name 表 nameID=1/4/6 | `Source Han Sans CN Normal`（**思源黑体 CN Normal 裁剪子集**） |
| sfntVersion | `\x00\x01\x00\x00`（TrueType / glyf） |
| unitsPerEm | 1000 |
| cmap 条目 | 83（本页用到的全部汉字） |
| 字形数 | 84（含 `.notdef`），**全部为简单字形，无复合字形** |

服务端行为：把真实汉字替换为「备用码位」，subfont 再把该码位的字形轮廓换成**真实汉字的轮廓**。因此 `textContent` 拿到的是替换字符，浏览器却显示正确汉字。

原样读取 vs 真实（本页实例）：

```
中悢共产党悤史上，第一悡明悞提出了“马克悝主义中悢化”…     ← textContent（乱码）
中国共产党历史上，第一个明确提出了“马克思主义中国化”…     ← 人眼所见 / 解码后
```

### 3.2 映射是**每页随机**的，静态对照表不可行

对比两份材料的内联字体（重叠 46 个混淆码位）：

| 混淆码位 | `章节测验.html`（09-17） | `做作业.html`（06-26） |
| --- | --- | --- |
| U+60A2 悢 | 国 | 举 |
| U+609D 悝 | 思 | 亡 |
| U+609E 悞 | 确 | 标 |
| …（共 46 个） | | |

**一致 0 / 不一致 46。**

结论：内置一张「混淆字 → 真字」静态表**行不通**——同一混淆码位在不同页面代表不同汉字。

### 3.3 可行方案：按字形二进制哈希反查（已实测验证）

社区成熟做法（52pojie 逆向分析帖、ScriptCat「超星字体解密」、`TellMeYourWish/chaoxing_solution_of_font_confusion`）的本质：

> 子集字体的 `glyf` 表**逐字节保留了源字体对应字形的原始数据**，因此只要离线建立
> `hash(字形原始字节) → 汉字`，运行时对页面字体做同样计算即可反查。

本报告采用的参考实现（`glyfSearch.py`）：

```python
hashed = (sha1(glyph.data).digest(), md5(glyph.data).digest())   # glyph.data = glyf 原始字节
```

### 3.4 验证结果（两条独立路径交叉验证）

| 页面字体 | 哈希表解码成功 | 与「栅格化字形形状匹配」交叉验证 |
| --- | --- | --- |
| `章节测验.html`（83 字） | 83 / 83，未命中 0 | **83 / 83 一致** |
| `做作业.html`（154 字） | 154 / 154，未命中 0 | **154 / 154 一致** |

交叉验证方法（独立于哈希表，用于排除「拿错参照表」的可能）：把子集字体的每个字形渲染为 32×32 归一化位图，与 `NotoSansSC-VF`（思源黑体同源设计）在 wght = 300/350/400 三档下渲染的全量汉字位图做余弦相似度匹配，三档投票取一致结果。82/83 相似度落在 0.87–0.98，且三档结论完全一致；唯一分歧位 `U+60D0`（候选 `＿ / 一 / －`，字形本就同形）判为「一」。

解码后阈值校验（人工可读性）：

```
【单选题】在中国共产党历史上，第一个明确提出了“马克思主义中国化”的科学命题和重大任务的是（）。
  A 毛泽东    B 李大钊    C 陈独秀    D 周恩来
【单选题】关于独立自主，下列做法中错误的是（）。
  A 坚持独立思考，走自己的路   ⚠ 后续选项见解码全文
```

### 3.5 内置解码表体积实测（全量 20902 汉字，U+4E00–U+9FA5）

| 键位宽 | 裸数据 | base64 | 哈希冲突 |
| --- | --- | --- | --- |
| `sha1[:4]` + 码位 | 122.5 KB | **163.3 KB** | 0 / 20902 |
| `sha1[:5]` + 码位 | 142.9 KB | **190.5 KB** | 0 / 20902 |
| `sha1[:6]` + 码位 | 163.3 KB | **217.7 KB** | 0 / 20902 |
| `sha1[:8]` + 码位 | 204.1 KB | 272.2 KB | 0 / 20902 |

gzip 基本无效（哈希本身随机，压缩率仅约 1–2%），因此体积按 base64 计。

---

## 四、RC-3：四层嵌套链路 —— 现有框架已支持，待真机确认

`src/infrastructure/frame-bridge.ts` 已具备多层能力，**无需大改**：

- `FrameBridge.requestRefresh()`（:96-112）递归遍历 iframe，`depth > 5` 才停止 → 覆盖 4 层
- `FrameAgent`（:119-196）在每个文档内 `window.parent.postMessage(message, this.parentOrigin())` 逐层回传，消息校验含 `protocolVersion`、`isTrustedOrigin`（`chaoxing.com` 及其子域）
- `src/userscript.meta.txt` 的 `@match *://chaoxing.com/*`、`*://*.chaoxing.com/*` 覆盖四层 URL

风险点（需真机确认）：

1. **每层都会加载 `@require docx`**（约 400 KB），4 层 = 4 份，是额外开销。
2. 若某一层 URL 落到 `@match` 之外（如 `*.chaoxing.com` 之外的静态域），该层不会注入，链路断开。
3. L4 的 `api/work` 响应是标准的答题页，`document.referrer` 未必是受信 origin，`parentOrigin()` 会退化为 `'*'`——功能可用但安全性略降。

---

## 五、RC-4：本快照没有答案（非缺陷）

L4 中答案容器为空：

```html
<!--新的答案开始-->
<!--新的答案结束-->
```

解码后全文 `正确答案` / `我的答案` / `答案解析` / `解析` 均 **0 命中**，`Zy_ulTop` 7 处、`num_option` 26 处但 `answerBg` / `answer_p` 0 命中。

说明这是**答题前**的视图。正确答案/错题/解析只在「提交后」或「已完成预览」视图中出现。若要验证答案提取，需**另外采集一份已提交状态的快照**（同一份 `html/做作业.html` 可作为对照，其体内也无答案）。

---

## 六、修复方案

### A. 题干选择器（必做，改动极小，收益最大）

`src/extractors/timu-extractor.ts:52` 补入 `.fontLabel`（并保留原有类名以兼容旧页面）；同时把题干内的题型标记 `.newZy_TItle` 从题干正文中剥离（现由 `typeMeta` 承载）。

修复后即可命中 5/5 题，并解开 `question-factory` 的 null 短路。

### B. 字体解码（必做，新增模块）

需要落地三件事：

1. **构建期**：从思源黑体生成 `hash(字形字节) → 汉字` 表（全量 20902 或常用子集），产出 base64 字符串
2. **运行期**：极简 TTF 解析器，读 `head`（`indexToLocFormat`）/ `maxp`（`numGlyphs`）/ `loca` / `glyf` / `cmap`，按码位取出字形的原始字节
3. **哈希**：SHA-1（Web Crypto 或内置精简实现），取前 N 字节查表

**体积取舍**

| 方案 | 内置体积 | 覆盖 | 取舍 |
| --- | --- | --- | --- |
| 常用字子集 | 约 70 KB | 通用规范汉字约 7000 字，日常题目覆盖 ~99.9% | 生僻字不解码（原样保留 + 调试日志） |
| 全量表（推荐） | 163–218 KB | 20902 字，零遗漏 | 脚本体积增加 |

解码结果建议按「页面文档」维度缓存，`MutationObserver` 触发重扫时不必重复解析字体。

### C. 顺带可修（A 组，证据见 `chapter-exercise-scan.md`）

`src/application/chapter-locator.ts`：

- `activate()`（:47-60）对容器 div 调 `.click()`，但 `onclick` 只挂在 `.posCatalog_name`（span）上 → 章节切换不生效
- 选择器首项 `.posCatalog_select` 把 28 个分组标题也当成章节 → 列表 200 项

---

## 七、复现与验证资产

本次诊断产出的可复用脚本（均在 `.workbuddy/`，已在 `.gitignore` 内）：

| 脚本 | 用途 |
| --- | --- |
| `split_layers.py` | 按 `<html>/<iframe>/<title>` 切分多层嵌套快照，输出层级边界 |
| `extract_font.py` | 从页面抽取内联 base64 字体并落盘 `.ttf` |
| `parse_font.py` | 解析字体 name / cmap / glyf / post，识别反爬机制 |
| `raster_match.py` | 栅格化字形形状匹配（不依赖外部表），得出真实字符 |
| `compare_maps.py` | 比对多页面映射一致性（证明映射随页面随机） |
| `validate_table.py` | 用 HanSansCN 预生成哈希表解码，并与栅格结果交叉验证 |
| `decode_page.py` | 整页解码，产出 `章节测验.decoded.html`（可读全文） |
| `build_table.py` | 生成并实测内置解码表的体积与冲突数 |

校验命令：

```bash
npm run validate      # typecheck + 单测 + 构建，改动 src/ 后必跑
npm run lint          # ESLint 0 警告
```
