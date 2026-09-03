import type { ExportOptions } from '../domain/export-options';
import {
  QUESTION_TYPE_LONG_LABELS,
  groupQuestions,
  type ExtractionResult,
  type Question,
  type RichContent,
} from '../domain/question';
import { richContentToText } from '../extractors/rich-content';

export function formatRichForText(content: RichContent): string {
  return richContentToText(content, (url) => `[图片: ${url}]`);
}

function questionLines(question: Question, number: number): string[] {
  const lines = [`${number}. ${formatRichForText(question.stem)}`];
  for (const option of question.options) {
    lines.push(`   ${option.key}. ${formatRichForText(option.content)}`);
  }
  return lines;
}

function answerText(question: Question): string {
  return formatRichForText(question.correctAnswer) || '暂无答案';
}

function appendQuestions(
  lines: string[],
  questions: readonly Question[],
  startNumber: number,
): number {
  let number = startNumber;
  for (const question of questions) {
    lines.push(...questionLines(question, number), '');
    number += 1;
  }
  return number;
}

function appendMainBody(lines: string[], result: ExtractionResult): void {
  let number = 1;
  if (result.chapters && result.chapters.length > 0) {
    for (const chapter of result.chapters) {
      lines.push(`【章节】${chapter.title}`, '');
      const groups = groupQuestions(chapter.questions);
      for (const type of chapter.typeOrder) {
        const questions = groups[type];
        if (questions.length === 0) continue;
        lines.push(`${QUESTION_TYPE_LONG_LABELS[type]}（共 ${questions.length} 题）`, '');
        number = appendQuestions(lines, questions, number);
      }
    }
    return;
  }

  const groups = groupQuestions(result.questions);
  for (const type of result.typeOrder) {
    const questions = groups[type];
    if (questions.length === 0) continue;
    lines.push(`${QUESTION_TYPE_LONG_LABELS[type]}（共 ${questions.length} 题）`, '');
    number = appendQuestions(lines, questions, number);
  }
}

function appendAnswerAppendix(
  lines: string[],
  result: ExtractionResult,
  includeAnalysis: boolean,
): void {
  lines.push('', '================ 参考答案 ================', '');
  result.questions.forEach((question, index) => {
    lines.push(`${index + 1}. ${answerText(question)}`);
    if (includeAnalysis && question.analysis.length > 0) {
      lines.push(`   解析：${formatRichForText(question.analysis)}`);
    }
  });
}

function appendWrongAppendix(
  lines: string[],
  result: ExtractionResult,
  includeAnalysis: boolean,
): void {
  const wrongQuestions = result.questions
    .map((question, index) => ({ question, number: index + 1 }))
    .filter(({ question }) => question.isWrong);
  if (wrongQuestions.length === 0) return;

  lines.push('', '================ 错题汇总 ================', '');
  for (const { question, number } of wrongQuestions) {
    lines.push(...questionLines(question, number));
    lines.push(`   我的答案：${formatRichForText(question.userAnswer) || '未作答'}`);
    lines.push(`   正确答案：${answerText(question)}`);
    if (includeAnalysis && question.analysis.length > 0) {
      lines.push(`   解析：${formatRichForText(question.analysis)}`);
    }
    lines.push('');
  }
}

export function formatText(result: ExtractionResult, options: ExportOptions): string {
  const lines = [
    result.title,
    `共 ${result.statistics.total} 题；错题 ${result.statistics.wrong} 题；提取时间 ${new Date(
      result.extractedAt,
    ).toLocaleString()}`,
    '',
  ];

  appendMainBody(lines, result);
  if (options.withAnswers) appendAnswerAppendix(lines, result, options.includeAnalysis);
  if (options.withWrong) appendWrongAppendix(lines, result, options.includeAnalysis);
  return `${lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
}
