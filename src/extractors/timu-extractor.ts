import type { Question, QuestionType } from '../domain/question';
import { parseLeadingNumber, textOf } from '../utils/dom';
import { hasExplicitWrongMarker } from './answer-comparison';
import { extractAnalysis, extractCorrectAnswer, extractUserAnswer } from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { resolveQuestionStem } from './question-stem';
import { detectQuestionType, inferQuestionType } from './question-type';

export class TiMuExtractor implements QuestionExtractor {
  readonly id = 'timu';
  readonly confidence = 110;

  supports(root: Document): boolean {
    return root.querySelector('.TiMu.newTiMu, #ZyBottom .TiMu') !== null;
  }

  extract(context: ExtractorContext): readonly Question[] {
    const questions: Question[] = [];
    const areas = Array.from(context.root.querySelectorAll<HTMLElement>('#ZyBottom .aiArea'));

    if (areas.length === 0) {
      context.root.querySelectorAll<HTMLElement>('.TiMu.newTiMu').forEach((container) => {
        const question = this.extractQuestion(container, null, context);
        if (question) questions.push(question);
      });
      return questions;
    }

    let currentType: QuestionType | null = null;
    for (const area of areas) {
      currentType =
        detectQuestionType(
          textOf(area.querySelector('.newTestType')),
          area.getAttribute('data-question-type'),
        ) ?? currentType;
      area.querySelectorAll<HTMLElement>('.TiMu.newTiMu, .TiMu').forEach((container) => {
        const question = this.extractQuestion(container, currentType, context);
        if (question) questions.push(question);
      });
    }
    return questions;
  }

  private extractQuestion(
    container: HTMLElement,
    inheritedType: QuestionType | null,
    context: ExtractorContext,
  ): Question | null {
    const title = container.querySelector('.Zy_TItle, .question-title') ?? container;
    const type =
      detectQuestionType(
        textOf(title.querySelector('.newZy_TItle')),
        container.getAttribute('typeName'),
        container.getAttribute('data-question-type'),
      ) ??
      inheritedType ??
      inferQuestionType(container);
    if (!type) return null;

    // 题干经分层定位，避免学习通改版换掉题干类名后整题被丢弃
    const stem = resolveQuestionStem(container, title);

    return createQuestion({
      number: parseLeadingNumber(textOf(title.querySelector('i.fl')) || stem.rawText),
      type,
      typeMeta: textOf(title.querySelector('.newZy_TItle')) || undefined,
      stem: stem.content,
      options: parseOptions(container, [
        '.Zy_ulTop.qtDetail > li',
        '.Zy_ulTop > li',
        '.answerBg',
        '.option-list > li',
      ]),
      correctAnswer: extractCorrectAnswer(container),
      userAnswer: extractUserAnswer(container),
      analysis: extractAnalysis(container),
      explicitWrong: hasExplicitWrongMarker(container),
      source: {
        extractor: this.id,
        pageUrl: context.pageUrl,
        selector: '.TiMu.newTiMu',
      },
    });
  }
}
