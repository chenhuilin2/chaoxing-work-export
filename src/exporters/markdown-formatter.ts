import type { ExportOptions } from '../domain/export-options';
import {
  QUESTION_TYPE_LONG_LABELS,
  groupQuestions,
  type ExtractionResult,
  type Question,
  type RichContent,
} from '../domain/question';
import { richContentToText } from '../extractors/rich-content';
import { escapeMarkdownAlt, escapeMarkdownUrl } from '../utils/text';

export function formatRichForMarkdown(content: RichContent): string {
  return richContentToText(content, (url, alt) => {
    return `![${escapeMarkdownAlt(alt)}](${escapeMarkdownUrl(url)})`;
  });
}

function escapeInline(value: string): string {
  return value.replace(/([\\`*_{}[\]<>])/g, '\\$1');
}

function questionLines(question: Question, number: number): string[] {
  const lines = [`${number}. ${formatRichForMarkdown(question.stem)}`];
  for (const option of question.options) {
    lines.push(`   - **${escapeInline(option.key)}.** ${formatRichForMarkdown(option.content)}`);
  }
  return lines;
}

function appendQuestionGroups(
  lines: string[],
  questions: readonly Question[],
  typeOrder: ExtractionResult['typeOrder'],
  startNumber: number,
): number {
  const groups = groupQuestions(questions);
  let number = startNumber;
  for (const type of typeOrder) {
    const typedQuestions = groups[type];
    if (typedQuestions.length === 0) continue;
    lines.push(`## ${QUESTION_TYPE_LONG_LABELS[type]}（共 ${typedQuestions.length} 题）`, '');
    for (const question of typedQuestions) {
      lines.push(...questionLines(question, number), '');
      number += 1;
    }
  }
  return number;
}

function appendMainBody(lines: string[], result: ExtractionResult): void {
  let number = 1;
  if (result.chapters && result.chapters.length > 0) {
    for (const chapter of result.chapters) {
      lines.push(`# ${escapeInline(chapter.title)}`, '');
      number = appendQuestionGroups(lines, chapter.questions, chapter.typeOrder, number);
    }
    return;
  }
  appendQuestionGroups(lines, result.questions, result.typeOrder, number);
}

function appendAnswerAppendix(
  lines: string[],
  result: ExtractionResult,
  includeAnalysis: boolean,
): void {
  lines.push('# 参考答案', '');
  result.questions.forEach((question, index) => {
    const answer = formatRichForMarkdown(question.correctAnswer) || '暂无答案';
    lines.push(`${index + 1}. **答案：** ${answer}`);
    if (includeAnalysis && question.analysis.length > 0) {
      lines.push(`   - **解析：** ${formatRichForMarkdown(question.analysis)}`);
    }
  });
  lines.push('');
}

function appendWrongAppendix(
  lines: string[],
  result: ExtractionResult,
  includeAnalysis: boolean,
): void {
  const wrong = result.questions
    .map((question, index) => ({ question, number: index + 1 }))
    .filter(({ question }) => question.isWrong);
  if (wrong.length === 0) return;

  lines.push('# 错题汇总', '');
  for (const { question, number } of wrong) {
    lines.push(...questionLines(question, number));
    lines.push(`   - **我的答案：** ${formatRichForMarkdown(question.userAnswer) || '未作答'}`);
    lines.push(
      `   - **正确答案：** ${formatRichForMarkdown(question.correctAnswer) || '暂无答案'}`,
    );
    if (includeAnalysis && question.analysis.length > 0) {
      lines.push(`   - **解析：** ${formatRichForMarkdown(question.analysis)}`);
    }
    lines.push('');
  }
}

export function formatMarkdown(result: ExtractionResult, options: ExportOptions): string {
  const lines = [
    `# ${escapeInline(result.title)}`,
    '',
    `> 共 ${result.statistics.total} 题；错题 ${result.statistics.wrong} 题；提取时间 ${new Date(
      result.extractedAt,
    ).toLocaleString()}`,
    '',
  ];
  appendMainBody(lines, result);
  if (options.withAnswers) appendAnswerAppendix(lines, result, options.includeAnalysis);
  if (options.withWrong) appendWrongAppendix(lines, result, options.includeAnalysis);
  return `${lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
}
