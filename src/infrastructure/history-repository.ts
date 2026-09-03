import type { ExportOptions } from '../domain/export-options';
import type { HistoryEntry } from '../domain/history';
import {
  buildStatistics,
  deriveTypeOrder,
  type ChapterExtraction,
  type ExtractionResult,
  type Question,
  type QuestionOption,
  type QuestionType,
  type RichContent,
} from '../domain/question';
import { textContent } from '../extractors/rich-content';
import { stableHash } from '../utils/hash';
import { SafeStorage } from './safe-storage';

const KEY = 'chaoxing-work-export:history:v3';
const LEGACY_KEY = 'xxt_history';
const MAX_ENTRIES = 10;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isHistoryEntry(value: unknown): value is HistoryEntry {
  if (!isRecord(value) || !isRecord(value.result) || !isRecord(value.options)) return false;
  return (
    typeof value.id === 'string' &&
    typeof value.createdAt === 'string' &&
    typeof value.title === 'string' &&
    typeof value.result.title === 'string' &&
    Array.isArray(value.result.questions) &&
    typeof value.options.format === 'string'
  );
}

function legacyType(value: string): QuestionType | null {
  const normalized = value.replace(/[【】()[\]（）\s]/gu, '');
  const map: Record<string, QuestionType> = {
    单选: 'single-choice',
    单选题: 'single-choice',
    单项选择题: 'single-choice',
    多选: 'multiple-choice',
    多选题: 'multiple-choice',
    多项选择题: 'multiple-choice',
    填空: 'fill-blank',
    填空题: 'fill-blank',
    判断: 'true-false',
    判断题: 'true-false',
    简答: 'short-answer',
    简答题: 'short-answer',
  };
  return map[normalized] ?? null;
}

function legacyRich(value: unknown, fallback?: unknown): RichContent {
  if (Array.isArray(value)) {
    const valid = value.filter((part) => {
      return isRecord(part) && ['text', 'image', 'break'].includes(String(part.type));
    });
    if (valid.length > 0) return valid as unknown as RichContent;
  }
  if (typeof value === 'string') return textContent(value);
  return typeof fallback === 'string' ? textContent(fallback) : [];
}

function legacyOptions(value: unknown): readonly QuestionOption[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((option, index) => {
    if (typeof option === 'string') {
      const match = option.match(/^\s*([A-Z])\s*[.、．:：]?\s*(.*)$/iu);
      return [
        {
          key: match?.[1]?.toUpperCase() ?? String.fromCharCode(65 + index),
          content: textContent(match?.[2] ?? option),
        },
      ];
    }
    if (!isRecord(option)) return [];
    const keySource = option.key ?? option.letter ?? option.label;
    const key = typeof keySource === 'string' ? keySource : String.fromCharCode(65 + index);
    return [
      {
        key: key.replace(/[.、．:：\s]/gu, '').toUpperCase(),
        content: legacyRich(option.content ?? option.richContent, option.text ?? option.value),
      },
    ];
  });
}

function legacyPayload(entry: Record<string, unknown>): Record<string, unknown> {
  return isRecord(entry.data) ? { ...entry, ...entry.data } : entry;
}

function extractLegacyQuestions(source: Record<string, unknown>): Question[] {
  const results = isRecord(source.results)
    ? source.results
    : isRecord(source.questionGroups)
      ? source.questionGroups
      : null;
  if (!results) return [];

  const questions: Question[] = [];
  for (const [legacyLabel, rawQuestions] of Object.entries(results)) {
    const type = legacyType(legacyLabel);
    if (!type || !Array.isArray(rawQuestions)) continue;
    rawQuestions.forEach((rawQuestion, index) => {
      if (!isRecord(rawQuestion)) return;
      const stem = legacyRich(
        rawQuestion.stemContent ?? rawQuestion.questionContent,
        rawQuestion.stem ?? rawQuestion.question,
      );
      const options = legacyOptions(rawQuestion.options);
      const fingerprint = `${legacyLabel}|${index}|${JSON.stringify(stem)}`;
      questions.push({
        id: `legacy-${stableHash(fingerprint)}`,
        number:
          typeof rawQuestion.qnum === 'number'
            ? rawQuestion.qnum
            : typeof rawQuestion.number === 'number'
              ? rawQuestion.number
              : undefined,
        type,
        typeMeta: typeof rawQuestion.typeMeta === 'string' ? rawQuestion.typeMeta : undefined,
        stem,
        options,
        correctAnswer: legacyRich(
          rawQuestion.correctAnswerContent ?? rawQuestion.answerContent,
          rawQuestion.correctAnswer ?? rawQuestion.answer,
        ),
        userAnswer: legacyRich(
          rawQuestion.myAnswerContent ?? rawQuestion.userAnswerContent,
          rawQuestion.myAnswer ?? rawQuestion.userAnswer,
        ),
        analysis: legacyRich(rawQuestion.analysisContent, rawQuestion.analysis),
        isWrong: rawQuestion.isWrong === true,
        source: {
          extractor: 'legacy-history',
          pageUrl: typeof source.sourceUrl === 'string' ? source.sourceUrl : '',
        },
      });
    });
  }
  return questions;
}

function migrateLegacyResult(entry: Record<string, unknown>): ExtractionResult | null {
  const source = legacyPayload(entry);
  const rawChapters = source.chapterDataList;
  const chapters: ChapterExtraction[] = [];

  if (Array.isArray(rawChapters)) {
    rawChapters.forEach((rawChapter, index) => {
      if (!isRecord(rawChapter)) return;
      const chapterSource = legacyPayload(rawChapter);
      const titleSource = chapterSource.title ?? chapterSource.chapterTitle ?? chapterSource.name;
      const title = typeof titleSource === 'string' ? titleSource : `第 ${index + 1} 章`;
      const idSource = chapterSource.id ?? chapterSource.chapterId;
      const id = typeof idSource === 'string' ? idSource : `legacy-chapter-${index + 1}`;
      const questions = extractLegacyQuestions(chapterSource).map((question, questionIndex) => ({
        ...question,
        id: `legacy-${stableHash(`${id}|${question.id}|${questionIndex}`)}`,
        chapterId: id,
        chapterTitle: title,
      }));
      if (questions.length === 0) return;
      chapters.push({
        id,
        title,
        questions,
        typeOrder: deriveTypeOrder(questions),
        statistics: buildStatistics(questions),
        sourceUrl: typeof chapterSource.sourceUrl === 'string' ? chapterSource.sourceUrl : '',
        extractor: 'legacy-history',
      });
    });
  }

  const questions = chapters.length > 0
    ? chapters.flatMap((chapter) => chapter.questions)
    : extractLegacyQuestions(source);
  if (questions.length === 0) return null;
  const titleSource = source.title ?? entry.title;
  const title = typeof titleSource === 'string' ? titleSource : '历史题目';
  return {
    title,
    questions,
    typeOrder: deriveTypeOrder(questions),
    statistics: buildStatistics(questions),
    sourceUrl: typeof source.sourceUrl === 'string' ? source.sourceUrl : '',
    extractor: 'legacy-history',
    extractedAt: typeof source.extractedAt === 'string' ? source.extractedAt : new Date().toISOString(),
    chapters: chapters.length > 0 ? chapters : undefined,
  };
}

function migrateLegacyEntry(value: unknown): HistoryEntry | null {
  if (!isRecord(value)) return null;
  const payload = legacyPayload(value);
  const result = migrateLegacyResult(value);
  if (!result) return null;
  const rawFormat = payload.format ?? (isRecord(payload.options) ? payload.options.format : undefined);
  const format = rawFormat === 'txt' || rawFormat === 'md' ? rawFormat : 'word';
  const title = typeof payload.title === 'string' ? payload.title : result.title;
  const legacyOptions = isRecord(payload.options) ? payload.options : payload;
  const options: ExportOptions = {
    format,
    filename: typeof legacyOptions.filename === 'string' ? legacyOptions.filename : title,
    withAnswers: legacyOptions.withAnswers === true,
    withWrong: legacyOptions.withWrong === true,
    includeAnalysis: legacyOptions.includeAnalysis !== false,
    shuffle: legacyOptions.shuffle === true,
    bankImport: legacyOptions.bankImport === true,
    splitByChapter: legacyOptions.splitByChapter === true,
  };
  const idSource = payload.id;
  const createdAtSource = payload.createdAt ?? payload.date;
  return {
    id: typeof idSource === 'string' ? idSource : `legacy-${Date.now()}-${stableHash(title)}`,
    createdAt: typeof createdAtSource === 'string' ? createdAtSource : new Date().toISOString(),
    title,
    result,
    options,
    textSnapshot: typeof payload.outputText === 'string' ? payload.outputText : undefined,
  };
}

export class HistoryRepository {
  constructor(private readonly storage = new SafeStorage()) {}

  load(): HistoryEntry[] {
    const current = this.storage.read(KEY);
    if (Array.isArray(current)) return current.filter(isHistoryEntry).slice(0, MAX_ENTRIES);

    const legacy = this.storage.read(LEGACY_KEY);
    if (Array.isArray(legacy)) {
      const migrated = legacy
        .map(migrateLegacyEntry)
        .filter((entry): entry is HistoryEntry => entry !== null)
        .slice(0, MAX_ENTRIES);
      return this.save(migrated);
    }
    return [];
  }

  add(entry: HistoryEntry): HistoryEntry[] {
    const entries = [entry, ...this.load().filter((existing) => existing.id !== entry.id)].slice(
      0,
      MAX_ENTRIES,
    );
    return this.save(entries);
  }

  remove(id: string): HistoryEntry[] {
    return this.save(this.load().filter((entry) => entry.id !== id));
  }

  clear(): void {
    this.storage.remove(KEY);
  }

  private save(entries: readonly HistoryEntry[]): HistoryEntry[] {
    let candidates = [...entries].slice(0, MAX_ENTRIES);
    while (candidates.length > 0) {
      if (this.storage.write(KEY, candidates)) return candidates;
      candidates = candidates.slice(0, -1);
    }
    this.storage.remove(KEY);
    return [];
  }
}
