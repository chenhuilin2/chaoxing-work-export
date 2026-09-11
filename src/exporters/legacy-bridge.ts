// 本文件将主脚本 chaoxing-work-export.user.js 的导出数据结构与公共辅助函数移植为 TypeScript，
// 供 TXT / Markdown / Word 导出直接复用原版实现（直接覆盖，不保留重构版差异）。
import {
  type ChapterExtraction,
  type ExtractionResult,
  type ImagePart,
  type Question,
  type QuestionType,
  type RichContent,
  type RichPart,
  type TextPart,
} from '../domain/question';

// 原版脚本的题型键（中文）
export type LegacyTypeKey = '单选' | '多选' | '填空' | '判断' | '简答';

// 原版脚本的选项结构
export interface LegacyOption {
  letter: string;
  text?: string;
  content: RichContent;
}

// 原版脚本的题目结构
export interface LegacyQuestion {
  stem: string;
  stemContent: RichContent;
  typeMeta?: string;
  options: LegacyOption[];
  correctAnswer: string;
  correctAnswerContent: RichContent;
  myAnswer: string;
  isWrong: boolean;
  images?: string[];
}

// 原版脚本按题型分组的题目集合
export type LegacyResults = Record<LegacyTypeKey, LegacyQuestion[]>;

// 原版脚本的提取结果结构
export interface LegacyResultData {
  title: string;
  results: LegacyResults;
  typeOrder: LegacyTypeKey[];
  wrongCount: number;
  hasMyAnswer: boolean;
  hasCorrectAnswer: boolean;
}

// 原版脚本的分章节数据结构
export interface LegacyChapterData {
  title: string;
  results: LegacyResults;
  typeOrder: LegacyTypeKey[];
  wrongCount: number;
  hasMyAnswer: boolean;
  hasCorrectAnswer: boolean;
}

// QuestionType → 原版中文题型键映射
const TYPE_KEY_MAP: Readonly<Record<QuestionType, LegacyTypeKey>> = {
  'single-choice': '单选',
  'multiple-choice': '多选',
  'fill-blank': '填空',
  'true-false': '判断',
  'short-answer': '简答',
};

// 创建原版结构的空题目集合（五种题型各一个空数组）
export function emptyLegacyResults(): LegacyResults {
  return { '单选': [], '多选': [], '填空': [], '判断': [], '简答': [] };
}

// ==================== 原版富文本辅助函数（逐行移植） ====================

// 判断富文本是否有实质内容（文本或图片）
export function hasRichContent(content: RichContent): boolean {
  return (content || []).some(
    (part) => part.type === 'image' || (part.type === 'text' && part.text.trim().length > 0),
  );
}

// 规范化富文本数组：合并相邻文本节点、去首尾空行、压缩空白
export function normalizeRichContent(parts: readonly RichPart[]): RichPart[] {
  const normalized: RichPart[] = [];
  const pushText = (part: TextPart): void => {
    if (!part.text) return;
    const value = part.text.replace(/\u00a0/g, ' ').replace(/[ \t\r\f]+/g, ' ');
    if (!value) return;
    const last = normalized[normalized.length - 1];
    // 合并相邻文本，保留第一个 part 的格式属性
    if (last && last.type === 'text') {
      normalized[normalized.length - 1] = { ...last, text: last.text + value };
    } else {
      normalized.push({
        type: 'text',
        text: value,
        bold: part.bold,
        italic: part.italic,
        subScript: part.subScript,
        superScript: part.superScript,
      });
    }
  };
  const pushBreak = (): void => {
    const last = normalized[normalized.length - 1];
    if (!last || last.type !== 'break') normalized.push({ type: 'break' });
  };

  for (const part of parts || []) {
    if (!part) continue;
    if (part.type === 'text') {
      pushText(part);
    } else if (part.type === 'image' && part.url) {
      normalized.push(part);
    } else if (part.type === 'break') {
      pushBreak();
    }
  }

  // 修剪首尾的换行与纯空白文本：尾部空白会挡住相邻换行，导致 Word 导出时题目与选项之间出现多余空行
  const isTrimmable = (part: RichPart | undefined): boolean => {
    if (!part) return false;
    return part.type === 'break' || (part.type === 'text' && !part.text.trim());
  };
  while (isTrimmable(normalized[0])) normalized.shift();
  while (isTrimmable(normalized[normalized.length - 1])) normalized.pop();
  return normalized;
}

// 将富文本数组转为纯文本，图片通过回调格式化
export function richContentToText(
  content: RichContent,
  imageFormatter?: (url: string, part: ImagePart) => string,
): string {
  let text = '';
  for (const part of content || []) {
    if (part.type === 'text') text += part.text;
    else if (part.type === 'image') text += imageFormatter ? imageFormatter(part.url, part) : '';
    else if (part.type === 'break') text += '\n';
  }
  return text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

// 适配器：兼容新旧数据格式，统一返回富文本数组
export function questionContent(q: LegacyQuestion): RichContent {
  if (hasRichContent(q.stemContent)) return q.stemContent;
  const parts: RichPart[] = q.stem ? [{ type: 'text', text: q.stem }] : [];
  if (q.images && q.images.length) {
    q.images.forEach((url) =>
      parts.push({ type: 'break' }, { type: 'image', url, alt: '' }),
    );
  }
  return normalizeRichContent(parts);
}

export function optionContent(opt: LegacyOption): RichContent {
  if (hasRichContent(opt.content)) return opt.content;
  return opt.text ? [{ type: 'text', text: opt.text }] : [];
}

export function answerContent(q: LegacyQuestion): RichContent {
  if (hasRichContent(q.correctAnswerContent)) return q.correctAnswerContent;
  return q.correctAnswer ? [{ type: 'text', text: q.correctAnswer }] : [];
}

// Markdown URL 转义，防止特殊字符破坏图片语法
export function escapeMarkdownAlt(text: string): string {
  return (text || '图片').replace(/[[\]\n\r]/g, ' ').trim() || '图片';
}

export function escapeMarkdownUrl(url: string): string {
  // 编码 URL 中的特殊字符，防止破坏 Markdown 图片语法，保留已编码部分
  return (url || '').replace(/[()\\]/g, (ch) => '%' + ch.charCodeAt(0).toString(16).toUpperCase());
}

// 富文本格式化：TXT/MD 分别处理
export function formatRichForText(content: RichContent): string {
  return richContentToText(content, (url) => `\n[图片: ${url}]\n`);
}

export function formatRichForMD(content: RichContent): string {
  return richContentToText(
    content,
    (url, part) => `\n![${escapeMarkdownAlt(part.alt)}](${escapeMarkdownUrl(url)})\n`,
  );
}

// 富文本转纯文本（忽略图片）
export function richTextOnly(content: RichContent): string {
  return richContentToText(content || [], () => '').replace(/\s+/g, ' ').trim();
}

// Fisher-Yates 洗牌算法，同类型题目内部打乱
export function shuffleQuestions(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): LegacyResults {
  const shuffled = emptyLegacyResults();
  for (const qtype of typeOrder) {
    const arr = [...(results[qtype] || [])];
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      // Fisher-Yates 交换（noUncheckedIndexedAccess 下需显式判空）
      const left = arr[i];
      const right = arr[j];
      if (left && right) {
        arr[i] = right;
        arr[j] = left;
      }
    }
    shuffled[qtype] = arr;
  }
  return shuffled;
}

// ==================== 领域模型 → 原版结构转换 ====================

// 将 TS 领域模型的单题转换为原版题目结构
function toLegacyQuestion(question: Question): LegacyQuestion {
  return {
    stem: formatRichForText(question.stem),
    stemContent: question.stem,
    typeMeta: question.typeMeta,
    options: question.options.map((option) => ({
      letter: option.key,
      text: richTextOnly(option.content),
      content: option.content,
    })),
    correctAnswer: richContentToText(question.correctAnswer, () => '')
      .replace(/\s+/g, ' ')
      .trim(),
    correctAnswerContent: question.correctAnswer,
    myAnswer: richTextOnly(question.userAnswer),
    isWrong: question.isWrong,
  };
}

// 将 TS 领域模型结果转换为原版提取数据结构（全部题目按题型合并）
export function toLegacyResult(result: ExtractionResult): LegacyResultData {
  const results = emptyLegacyResults();
  for (const question of result.questions) {
    results[TYPE_KEY_MAP[question.type]].push(toLegacyQuestion(question));
  }
  const typeOrder = result.typeOrder.map((type) => TYPE_KEY_MAP[type]);
  return {
    title: result.title,
    results,
    typeOrder,
    wrongCount: result.statistics.wrong,
    hasMyAnswer: result.statistics.withUserAnswer > 0,
    hasCorrectAnswer: result.statistics.withCorrectAnswer > 0,
  };
}

// 将 TS 章节数据转换为原版分章节数据结构
export function toLegacyChapter(chapter: ChapterExtraction): LegacyChapterData {
  const results = emptyLegacyResults();
  for (const question of chapter.questions) {
    results[TYPE_KEY_MAP[question.type]].push(toLegacyQuestion(question));
  }
  const typeOrder = chapter.typeOrder.map((type) => TYPE_KEY_MAP[type]);
  return {
    title: chapter.title,
    results,
    typeOrder,
    wrongCount: chapter.statistics.wrong,
    hasMyAnswer: chapter.statistics.withUserAnswer > 0,
    hasCorrectAnswer: chapter.statistics.withCorrectAnswer > 0,
  };
}
