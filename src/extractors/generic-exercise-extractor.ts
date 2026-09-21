import type { Question } from '../domain/question';
import { allMatches, parseLeadingNumber, textOf } from '../utils/dom';
import { uniqueBy } from '../utils/array';
import { hasExplicitWrongMarker } from './answer-comparison';
import { extractAnalysis, extractCorrectAnswer, extractUserAnswer } from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { resolveQuestionStem } from './question-stem';
import { detectQuestionType, inferQuestionType } from './question-type';

const CONTAINER_SELECTORS = [
  '[data-question-id]',
  '.question-item',
  '.subject-item',
  '.exercise-question',
  '.TiMu',
  '.questionLi',
] as const;

export class GenericExerciseExtractor implements QuestionExtractor {
  readonly id = 'generic-exercise';
  readonly confidence = 40;

  supports(root: Document): boolean {
    return allMatches(root, CONTAINER_SELECTORS).length > 0;
  }

  extract(context: ExtractorContext): readonly Question[] {
    const containers = uniqueBy(allMatches<Element>(context.root, CONTAINER_SELECTORS), (element) => {
      return element.getAttribute('data-question-id') ?? `${element.tagName}:${textOf(element).slice(0, 80)}`;
    });
    const questions: Question[] = [];

    for (const container of containers) {
      const type =
        detectQuestionType(
          container.getAttribute('typeName'),
          container.getAttribute('data-question-type'),
          textOf(container.querySelector('.question-type, .type_tit, .newTestType')),
        ) ?? inferQuestionType(container);
      if (!type) continue;

      const stem = resolveQuestionStem(container);
      const question = createQuestion({
        number: parseLeadingNumber(stem.rawText),
        type,
        typeMeta: textOf(container.querySelector('.question-type, .colorShallow')) || undefined,
        stem: stem.content,
        options: parseOptions(container, [
          '.answerBg',
          '.option-list > li',
          '.options > li',
          '.mark_letter > li',
          '.Zy_ulTop > li',
          'label.option',
        ]),
        correctAnswer: extractCorrectAnswer(container),
        userAnswer: extractUserAnswer(container),
        analysis: extractAnalysis(container),
        explicitWrong: hasExplicitWrongMarker(container),
        source: {
          extractor: this.id,
          pageUrl: context.pageUrl,
          selector: CONTAINER_SELECTORS.join(', '),
        },
      });
      if (question) questions.push(question);
    }
    return questions;
  }
}
