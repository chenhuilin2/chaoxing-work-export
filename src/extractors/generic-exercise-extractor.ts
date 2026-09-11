import type { Question } from '../domain/question';
import { allMatches, parseLeadingNumber, textOf } from '../utils/dom';
import { uniqueBy } from '../utils/array';
import { hasExplicitWrongMarker } from './answer-comparison';
import { extractAnalysis, extractCorrectAnswer, extractUserAnswer } from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { detectQuestionType, inferQuestionType } from './question-type';
import { extractRichContent, isRichContentEmpty, stripQuestionPrefix } from './rich-content';

const CONTAINER_SELECTORS = [
  '[data-question-id]',
  '.question-item',
  '.subject-item',
  '.exercise-question',
  '.TiMu',
  '.questionLi',
] as const;

const STEM_SELECTORS = [
  '[data-role="stem"]',
  '.question-stem',
  '.questionStem',
  '.subject-title',
  '.subject',
  '.qtContent',
  '.mark_name',
  '.Zy_TItle .qtContent',
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

      let stemElement: Element | null = null;
      for (const selector of STEM_SELECTORS) {
        const candidate = container.querySelector(selector);
        if (!candidate) continue;
        const content = stripQuestionPrefix(extractRichContent(candidate));
        if (!isRichContentEmpty(content)) {
          stemElement = candidate;
          break;
        }
      }
      if (!stemElement) continue;

      const question = createQuestion({
        number: parseLeadingNumber(textOf(stemElement)),
        type,
        typeMeta: textOf(container.querySelector('.question-type, .colorShallow')) || undefined,
        stem: stripQuestionPrefix(extractRichContent(stemElement)),
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
