import type { Question } from '../domain/question';
import { parseLeadingNumber, textOf } from '../utils/dom';
import { hasExplicitWrongMarker } from './answer-comparison';
import { extractAnalysis, extractCorrectAnswer, extractUserAnswer } from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { detectQuestionType } from './question-type';
import { extractRichContent, stripQuestionPrefix } from './rich-content';

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
        const stemElement = container.querySelector('.qtContent, .mark_name, .question-stem');
        const type =
          detectQuestionType(
            container.getAttribute('typeName'),
            container.getAttribute('data-question-type'),
            textOf(container.querySelector('.colorShallow')),
          ) ?? sectionType;
        const rawStem = extractRichContent(stemElement);
        const question = createQuestion({
          number: parseLeadingNumber(textOf(stemElement)),
          type,
          typeMeta: textOf(container.querySelector('.colorShallow')) || undefined,
          stem: stripQuestionPrefix(rawStem),
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
