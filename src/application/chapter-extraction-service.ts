import {
  buildStatistics,
  deriveTypeOrder,
  type ChapterExtraction,
  type ExtractionResult,
  type Question,
} from '../domain/question';
import { delay } from '../utils/async';
import { stableHash } from '../utils/hash';
import { ChapterLocator, type ChapterDescriptor } from './chapter-locator';
import { ExtractionService } from './extraction-service';

export interface ChapterProgress {
  readonly completed: number;
  readonly total: number;
  readonly chapter: ChapterDescriptor;
  readonly state: 'loading' | 'success' | 'failed';
  readonly message?: string;
}

export type ChapterProgressListener = (progress: ChapterProgress) => void;

/** Sequentially visits selected chapters and produces a single aggregate snapshot. */
export class ChapterExtractionService {
  constructor(
    private readonly extractionService: ExtractionService,
    private readonly locator = new ChapterLocator(),
  ) {}

  listChapters(): ChapterDescriptor[] {
    return this.locator.list();
  }

  async extractSelected(
    selectedIndexes: readonly number[],
    onProgress: ChapterProgressListener = () => undefined,
  ): Promise<ExtractionResult> {
    const chapters = this.locator.list();
    const selected = [...new Set(selectedIndexes)]
      .map((index) => chapters[index])
      .filter((chapter): chapter is ChapterDescriptor => Boolean(chapter));
    if (selected.length === 0) throw new Error('未选择可提取的章节');

    const originalIndex = this.locator.activeIndex();
    const chapterResults: ChapterExtraction[] = [];
    let previousFingerprint = this.extractionService.fingerprint(this.extractionService.extract());

    try {
      for (let index = 0; index < selected.length; index += 1) {
        const chapter = selected[index] as ChapterDescriptor;
        onProgress({
          completed: index,
          total: selected.length,
          chapter,
          state: 'loading',
        });

        const alreadyActive = this.locator.activeIndex() === chapter.index;
        const startedAt = Date.now();
        if (!alreadyActive && !this.locator.activate(chapter.index)) {
          onProgress({
            completed: index,
            total: selected.length,
            chapter,
            state: 'failed',
            message: '章节入口已变化或不可点击',
          });
          continue;
        }

        await delay(alreadyActive ? 150 : 450);
        const result = alreadyActive
          ? this.extractionService.extract()
          : await this.extractionService.waitForChangedResult({
              afterTimestamp: startedAt,
              previousFingerprint,
              timeoutMs: 14_000,
              intervalMs: 350,
            });

        if (!result) {
          onProgress({
            completed: index + 1,
            total: selected.length,
            chapter,
            state: 'failed',
            message: '未在限定时间内识别到题目',
          });
          continue;
        }

        const questions = result.questions.map((question, questionIndex) => {
          return this.attachChapter(question, chapter, questionIndex);
        });
        const chapterExtraction: ChapterExtraction = {
          id: chapter.id,
          title: chapter.title,
          questions,
          typeOrder: deriveTypeOrder(questions),
          statistics: buildStatistics(questions),
          sourceUrl: result.sourceUrl,
          extractor: result.extractor,
        };
        chapterResults.push(chapterExtraction);
        previousFingerprint = this.extractionService.fingerprint(result);

        onProgress({
          completed: index + 1,
          total: selected.length,
          chapter,
          state: 'success',
        });
      }
    } finally {
      if (originalIndex !== null) this.locator.activate(originalIndex);
    }

    if (chapterResults.length === 0) throw new Error('所选章节均未提取到题目');

    const questions = chapterResults.flatMap((chapter) => chapter.questions);
    return {
      title: this.aggregateTitle(chapterResults),
      questions,
      typeOrder: deriveTypeOrder(questions),
      statistics: buildStatistics(questions),
      sourceUrl: window.location.href,
      extractor: [...new Set(chapterResults.map((chapter) => chapter.extractor))].join('+'),
      extractedAt: new Date().toISOString(),
      chapters: chapterResults,
    };
  }

  private attachChapter(
    question: Question,
    chapter: ChapterDescriptor,
    questionIndex: number,
  ): Question {
    return {
      ...question,
      id: stableHash(`${chapter.id}|${question.id}|${questionIndex}`),
      chapterId: chapter.id,
      chapterTitle: chapter.title,
    };
  }

  private aggregateTitle(chapters: readonly ChapterExtraction[]): string {
    const documentTitle = document.title.trim();
    const suffix = chapters.length === 1 ? chapters[0]?.title : `${chapters.length} 个章节`;
    return [documentTitle || '学习通课程', suffix].filter(Boolean).join(' - ');
  }
}
