import type { Question } from '../domain/question';
import { parseLeadingNumber, textOf } from '../utils/dom';
import { hasExplicitWrongMarker } from './answer-comparison';
import { extractAnalysis, extractCorrectAnswer, extractUserAnswer } from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { resolveQuestionStem } from './question-stem';
import { detectQuestionType } from './question-type';

export class MarkItemExtractor implements QuestionExtractor {
  readonly id = 'mark-item';
  readonly confidence = 100;

  supports(root: Document): boolean {
    return root.querySelector('.mark_item .questionLi') !== null;
  }

  extract(context: ExtractorContext): readonly Question[] {
    const questions: Question[] = [];
    context.root.querySelectorAll<HTMLElement>('.mark_item').forEach((section) => {
      const sectionType = detectQuestionType(textOf(section.querySelector('.type_tit')));
      if (!sectionType) return;

      section.querySelectorAll<HTMLElement>('.questionLi').forEach((container) => {
        const type =
          detectQuestionType(
            container.getAttribute('typeName'),
            container.getAttribute('data-question-type'),
            textOf(container.querySelector('.colorShallow')),
          ) ?? sectionType;
        const stem = resolveQuestionStem(container);
        const question = createQuestion({
          number: parseLeadingNumber(stem.rawText),
          type,
          typeMeta: textOf(container.querySelector('.colorShallow')) || undefined,
          stem: stem.content,
          options: parseOptions(container, [
            '.mark_letter > li',
            '.mark_letter li',
            '.answerBg',
          ]),
          correctAnswer: extractCorrectAnswer(container),
          userAnswer: extractUserAnswer(container),
          analysis: extractAnalysis(container),
          explicitWrong: hasExplicitWrongMarker(container),
          source: {
            extractor: this.id,
            pageUrl: context.pageUrl,
            selector: '.mark_item .questionLi',
          },
        });
        if (question) questions.push(question);
      });
    });
    return questions;
  }
}
