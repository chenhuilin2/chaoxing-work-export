import type { Question } from '../domain/question';
import { parseLeadingNumber, textOf } from '../utils/dom';
import { hasExplicitWrongMarker } from './answer-comparison';
import { extractAnalysis, extractCorrectAnswer, extractUserAnswer } from './common-answer';
import type { ExtractorContext, QuestionExtractor } from './contracts';
import { parseOptions } from './option-parser';
import { createQuestion } from './question-factory';
import { detectQuestionType, inferQuestionType } from './question-type';
import { extractRichContent, stripQuestionPrefix } from './rich-content';

export class QuestionLiExtractor implements QuestionExtractor {
  readonly id = 'question-li';
  readonly confidence = 90;

  supports(root: Document): boolean {
    return root.querySelector('.questionLi') !== null;
  }

  extract(context: ExtractorContext): readonly Question[] {
    const questions: Question[] = [];
    context.root.querySelectorAll<HTMLElement>('.questionLi').forEach((container) => {
      const type =
        detectQuestionType(
          container.getAttribute('typeName'),
          container.getAttribute('typename'),
          container.getAttribute('data-question-type'),
          textOf(container.querySelector('.type_tit, .question-type, .colorShallow')),
        ) ?? inferQuestionType(container);
      if (!type) return;

      const stemElement = container.querySelector(
        '.mark_name, .qtContent, .questionStem, .question-stem, .subject, [data-role="stem"]',
      );
      const question = createQuestion({
        number: parseLeadingNumber(textOf(stemElement)),
        type,
        typeMeta: textOf(container.querySelector('.colorShallow, .question-type')) || undefined,
        stem: stripQuestionPrefix(extractRichContent(stemElement)),
        options: parseOptions(container, [
          '.answerBg',
          '.mark_letter > li',
          '.option-list > li',
          '.options > li',
        ]),
        correctAnswer: extractCorrectAnswer(container),
        userAnswer: extractUserAnswer(container),
        analysis: extractAnalysis(container),
        explicitWrong: hasExplicitWrongMarker(container),
        source: {
          extractor: this.id,
          pageUrl: context.pageUrl,
          selector: '.questionLi',
        },
      });
      if (question) questions.push(question);
    });
    return questions;
  }
}
