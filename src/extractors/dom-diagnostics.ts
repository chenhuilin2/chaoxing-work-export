import type { ExtractionResult } from '../domain/question';
import { collectAccessibleDocuments, firstMatch } from '../utils/dom';
import { CompositeExtractor } from './composite-extractor';
import { getCxSecretDecoder } from './cxsecret-decoder';
import { FALLBACK_OPTION_SELECTORS, parseOptions } from './option-parser';
import { STEM_CONTAINER_SELECTORS } from './question-stem';
import { extractRichContent, richContentToText } from './rich-content';
import { TIMU_OPTION_SELECTORS } from './timu-extractor';

/**
 * 提取诊断：把各层文档里「题目容器与选项容器的真实结构」汇总成一段纯文本报告。
 *
 * 为什么需要它：离线快照是「渲染完成之后」的 DOM，「提取到题干、提取不到选项」这类
 * 现场恰恰复现不出来（同一份快照离线能读全选项）。所以只能把现场的结构带回来，
 * 据此区分三种成因：选择器没命中 / 选项在题目容器之外 / 选项节点本身解析为空。
 *
 * 报告只读取 DOM，不修改页面；内容按块截断，避免剪贴板过大。
 */

export interface DiagnosticsOptions {
  /** 诊断的根文档，默认取全局 document */
  readonly root?: Document;
  /** 脚本版本，写进报告头部（避免每次都要先问「你装的是哪一版」） */
  readonly version?: string;
}

// 报告总长度上限与各块截断长度：只保留判断结构所需的信息
const MAX_REPORT_CHARS = 9_000;
const MAX_QUESTIONS_PER_DOC = 6;
const MAX_OUTLINE_DEPTH = 3;

/** 一个文档里可能承载题目的容器，与各提取器的 supports() 保持一致 */
const QUESTION_CONTAINER_SELECTORS = [
  '.TiMu.newTiMu',
  '.TiMu',
  '.questionLi',
  '.mark_item',
  '.answerBg',
] as const;

/** 题目容器的外层结构：选项若不在题目容器里，最可能落在这一层 */
const QUESTION_SCOPE_SELECTORS = ['.singleQuesId', '.questionLi', '.mark_item'] as const;

/** 需要选项才算完整的题型（此处就地判断，避免反向依赖应用层） */
const CHOICE_TYPES = new Set(['single-choice', 'multiple-choice']);

/** 元素结构速览的剩余节点预算 */
interface OutlineBudget {
  left: number;
}

function safeQuery(root: ParentNode, selector: string): Element | null {
  try {
    return root.querySelector(selector);
  } catch {
    return null;
  }
}

function safeQueryAllCount(root: ParentNode, selector: string): number {
  try {
    return root.querySelectorAll(selector).length;
  } catch {
    return -1;
  }
}

/** 空白折叠 + 截断，便于把结构压进一行 */
function clip(value: string, limit: number): string {
  const text = value.replace(/\s+/gu, ' ').trim();
  return text.length <= limit ? text : `${text.slice(0, limit)}…(+${text.length - limit}字)`;
}

function describeElement(element: Element): string {
  const tag = element.tagName.toLowerCase();
  const id = element.getAttribute('id');
  const className = (element.getAttribute('class') ?? '').trim().replace(/\s+/gu, '.');
  return `${tag}${id ? `#${id}` : ''}${className ? `.${className}` : ''}`;
}

/**
 * 元素文本：走与提取相同的富文本管道（含 cxsecret 还原）。
 * 直接读 textContent 拿到的是字体反爬的混淆码位，报告里会满屏乱码、无法核对。
 */
function readableText(element: Element | null): string {
  if (!element) return '';
  return richContentToText(extractRichContent(element));
}

/** 元素结构速览：标签.类名 > 子元素；深度与节点数都受限，足以看出选项容器在不在 */
function outline(node: Element, depth: number, budget: OutlineBudget): string {
  if (budget.left <= 0) return '…';
  budget.left -= 1;
  const name = describeElement(node);
  const children = Array.from(node.children);
  if (children.length === 0) return name;
  if (depth <= 1) return `${name}(+${children.length} 个子元素)`;
  const parts: string[] = [];
  for (const child of children) {
    if (budget.left <= 0) {
      parts.push('…');
      break;
    }
    parts.push(outline(child, depth - 1, budget));
  }
  return `${name} > [${parts.join(', ')}]`;
}

/** 题目的 qid：答题页放在 .singleQuesId 的 id/属性上，用于跨容器找选项 */
function resolveQuestionId(container: Element): string {
  const scope = container.closest('.singleQuesId') ?? container;
  const raw = [scope.getAttribute('id'), scope.getAttribute('data'), container.getAttribute('data')]
    .filter(Boolean)
    .join(' ');
  return raw.match(/\d{4,}/u)?.[0] ?? '(无)';
}

/** 逐候选选择器统计容器内的选项节点数，命中即说明该候选可用 */
function describeOptionCounts(container: Element): string {
  return [...TIMU_OPTION_SELECTORS, ...FALLBACK_OPTION_SELECTORS]
    .map((selector) => `${selector}=${safeQueryAllCount(container, selector)}`)
    .join(' ');
}

/** 选项节点总数（容器内任一候选命中即为该题选项） */
function optionNodeCount(container: Element): number {
  for (const selector of [...TIMU_OPTION_SELECTORS, ...FALLBACK_OPTION_SELECTORS]) {
    const count = safeQueryAllCount(container, selector);
    if (count > 0) return count;
  }
  return 0;
}

/**
 * 容器内实际解析出的选项条数（走与提取完全相同的管道）。
 *
 * 与「候选命中数」对照即可把成因分成两类：命中数为 0 是选择器没命中；
 * 命中数大于 0、解析数为 0，则是节点找到了但读不出文字（历史真机现场即此类）。
 */
function parsedOptionCount(container: Element): number {
  try {
    return parseOptions(container, TIMU_OPTION_SELECTORS).length;
  } catch {
    return -1;
  }
}

/** 单题的诊断块：容器、qid、题干来源、选项计数与结构速览 */
function describeQuestion(container: Element, index: number): string[] {
  const matched =
    QUESTION_CONTAINER_SELECTORS.find((selector) => {
      try {
        return container.matches(selector);
      } catch {
        return false;
      }
    }) ?? '(未匹配已知容器)';
  const title = safeQuery(container, '.Zy_TItle, .question-title') ?? container;
  const stemElement =
    firstMatch(title, STEM_CONTAINER_SELECTORS) ?? firstMatch(container, STEM_CONTAINER_SELECTORS);
  const stemSelector =
    STEM_CONTAINER_SELECTORS.find((selector) => safeQuery(title, selector) !== null) ??
    STEM_CONTAINER_SELECTORS.find((selector) => safeQuery(container, selector) !== null) ??
    '(未命中，走兜底)';
  const qid = resolveQuestionId(container);
  const options = optionNodeCount(container);
  const parsed = parsedOptionCount(container);

  const head =
    `  ${index + 1}) 容器=${matched} qid=${qid} 题干来源=${stemSelector} ` +
    `题干=${clip(readableText(stemElement), 60) || '(空)'}`;
  const counts = `     容器内选项计数: ${describeOptionCounts(container)} 解析出选项=${parsed} 条`;

  // 节点命中且能解析出选项：该题选项链路正常，无需展开结构
  if (options > 0 && parsed > 0) {
    return [head, counts];
  }

  // 其余情况展开结构。同时打印「原样 textContent」与「提取管道读出的文本」：
  // 原文有文字、管道为空，即说明文字确实在节点里、只是没被提取管道读到
  // （历史上这条差异来自跨 window 的构造器判定，见 rich-content.ts 的 isElementNode）。
  const lines = [head, counts];
  const items = Array.from(container.querySelectorAll('li')).slice(0, 8);
  const raw = items.map((item) => clip(item.textContent ?? '', 24) || '(空)');
  const piped = items.map((item) => clip(readableText(item), 24) || '(空)');
  lines.push(
    `     容器内 li 原文(${container.querySelectorAll('li').length} 个): ${raw.join(' | ') || '(无)'}`,
  );
  lines.push(`     容器内 li 管道: ${piped.join(' | ') || '(无)'}`);
  for (const selector of QUESTION_SCOPE_SELECTORS) {
    const scope = container.closest(selector);
    if (!scope || scope === container) continue;
    lines.push(`     外层 ${selector}: 选项计数 ${describeOptionCounts(scope)}`);
    lines.push(`     外层结构: ${clip(outline(scope, MAX_OUTLINE_DEPTH, { left: 26 }), 700)}`);
  }
  if (qid !== '(无)') {
    const owner = container.ownerDocument;
    const byQid = safeQueryAllCount(owner, `li[qid="${qid}"]`);
    const byData = safeQueryAllCount(owner, `li[data-qid="${qid}"]`);
    lines.push(
      `     按 qid 全文档查找: li[qid]=${byQid} li[data-qid]=${byData}` +
        `（>0 说明选项在题目容器之外）`,
    );
  }
  lines.push(`     容器结构: ${clip(outline(container, MAX_OUTLINE_DEPTH, { left: 26 }), 900)}`);
  return lines;
}

/** 单个文档的诊断块 */
function describeDocument(root: Document, depth: number, extractor: CompositeExtractor): string[] {
  const url = (() => {
    try {
      return root.location?.href ?? '(无 url)';
    } catch {
      return '(跨域)';
    }
  })();
  const decoder = getCxSecretDecoder(root);
  const containers: Element[] = [];
  for (const selector of QUESTION_CONTAINER_SELECTORS) {
    const found = Array.from(root.querySelectorAll(selector));
    for (const element of found) if (!containers.includes(element)) containers.push(element);
  }

  const lines = [
    `[L${depth}] readyState=${root.readyState ?? '未知'} 题目容器=${QUESTION_CONTAINER_SELECTORS.map(
      (selector) => `${selector}×${safeQueryAllCount(root, selector)}`,
    ).join(' ')}`,
    `     cxsecret: 作用域标记=${decoder ? (decoder.scoped ? '有' : '无') : '无字体'} 解码表=${
      decoder?.size ?? 0
    } 条`,
    `     url=${clip(url, 120)}`,
  ];

  const result: ExtractionResult | null = extractor.extract(root);
  if (result) {
    const missing = result.questions.filter(
      (question) => CHOICE_TYPES.has(question.type) && question.options.length === 0,
    ).length;
    lines.push(
      `     链路自检: 提取器=${result.extractor} 题数=${result.questions.length} ` +
        `选项总数=${result.questions.reduce((total, question) => total + question.options.length, 0)} ` +
        `缺选项选择题=${missing}`,
    );
  } else {
    lines.push('     链路自检: 该文档未提取到题目');
  }

  containers.slice(0, MAX_QUESTIONS_PER_DOC).forEach((container, index) => {
    lines.push(...describeQuestion(container, index));
  });
  if (containers.length > MAX_QUESTIONS_PER_DOC) {
    lines.push(`  （其余 ${containers.length - MAX_QUESTIONS_PER_DOC} 个题目容器已省略）`);
  }
  return lines;
}

/**
 * 生成当前页面的提取诊断报告。
 * 报告包含：各层文档（含同源 iframe）的就绪状态、题目容器数量、逐题的选项候选命中数、
 * 容器/外层结构速览、以及提取链路自检结果。
 */
export function buildExtractionDiagnostics(options: DiagnosticsOptions = {}): string {
  const root =
    options.root ?? (typeof document === 'undefined' ? null : (document as Document | null));
  const lines: string[] = [
    `=== 学习通题目导出 · 提取诊断 ${options.version ? `v${options.version} ` : ''}${new Date().toISOString()} ===`,
    '说明：把整段内容复制回传，即可定位「提取到题干、提取不到选项」的成因。',
    '读法：每题的「解析出选项」= 提取管道真正读出的选项条数；候选命中数>0 而解析数=0，',
    '说明节点找到了但读不出文字。「li 原文」是未处理的 textContent、「li 管道」是提取管道读到的',
    '文本，原文有而管道为空即属此类（字体反爬页面的原文会是混淆码位，以管道为准）。',
  ];
  if (!root) {
    lines.push('（当前环境没有 document，无法诊断）');
    return lines.join('\n');
  }

  const extractor = new CompositeExtractor();
  for (const { document: current, depth } of collectAccessibleDocuments(root)) {
    lines.push(...describeDocument(current, depth, extractor));
  }

  const report = lines.join('\n');
  if (report.length <= MAX_REPORT_CHARS) return report;
  return `${report.slice(0, MAX_REPORT_CHARS)}\n…（报告已截断）`;
}
