// TXT 格式化：以下函数逐行移植自主脚本 chaoxing-work-export.user.js，保持输出完全一致。
//
// 多章节导出时按「章 → 小节」两级分段，且每章后面紧跟该章的答案与错题（见文件末尾的分段实现）。
// 下面这些单章 formatter 本来就以「传入的结果集」为范围从 1 开始编号，按小节传入即自动按小节重排，
// 因此不需要为多章节另写一套编号逻辑。
import {
  type LegacyResults,
  type LegacyTypeKey,
  answerContent,
  formatRichForText,
  optionContent,
  questionContent,
} from './legacy-bridge';
import {
  type ExportSection,
  type SectionGroup,
  READABLE_GAPS,
  blankLines,
  chapterBlocks,
  renderSectionGroups,
} from './export-plan';

// ==================== TXT 格式化（富文本） ====================
export function formatOutput(results: LegacyResults, typeOrder: readonly LegacyTypeKey[]): string {
  const typeLabels: Record<string, string> = {
    单选: '单选题',
    多选: '多选题',
    填空: '填空题',
    判断: '判断题',
    简答: '简答题',
  };
  const typeNumbers = ['一', '二', '三', '四', '五', '六'];

  let output = '';
  let globalNum = 0;
  let sectionIdx = 0;

  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;

    const label = typeLabels[qtype] ?? qtype;
    const num = typeNumbers[sectionIdx] || sectionIdx + 1;
    output += `${num}. ${label}（共${questions.length}题）\n`;

    for (const q of questions) {
      globalNum++;
      output += `${globalNum}. ${formatRichForText(questionContent(q))}\n`;
      if (q.options && q.options.length > 0) {
        for (const opt of q.options) {
          output += `${opt.letter}. ${formatRichForText(optionContent(opt))}\n`;
        }
      }
      output += '\n';
    }
    sectionIdx++;
  }
  return output.trim();
}

export function formatAnswersTXT(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  const typeLabels: Record<string, string> = {
    单选: '单选题',
    多选: '多选题',
    填空: '填空题',
    判断: '判断题',
    简答: '简答题',
  };
  const typeNumbers = ['一', '二', '三', '四', '五', '六'];

  let output = '';
  let globalNum = 0;
  let sectionIdx = 0;
  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;

    const num = typeNumbers[sectionIdx] || sectionIdx + 1;
    output += `${num}、${typeLabels[qtype] ?? qtype}\n\n`;
    sectionIdx++;

    for (const q of questions) {
      globalNum++;
      const answer = formatRichForText(answerContent(q)) || '（未找到答案）';
      if (qtype === '填空' && answer.includes('；')) {
        const parts = answer.split('；').map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
        output += `${globalNum}. \n`;
        parts.forEach((part, i) => {
          output += `(${i + 1}) ${part}\n`;
        });
        output += '\n';
      } else {
        output += `${globalNum}. ${answer}\n\n`;
      }
    }
  }
  return output.trim();
}

export function formatOutputWithAnswers(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = formatOutput(results, typeOrder);
  output += '\n\n\n';
  output += '========================================\n';
  output += '              答案汇总\n';
  output += '========================================\n\n';
  output += formatAnswersTXT(results, typeOrder);
  return output.trim();
}

/**
 * 错题正文（不含「错题汇总」页头）。
 *
 * 题号以传入的结果集为范围从 1 数起，因此按章传入即得该章的错题题号 —— 多章节分段与
 * 单章节页头拼接都复用它，避免两处各写一遍编号。
 */
function wrongQuestionsBodyTXT(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = '';
  let globalNum = 0;
  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;
    if (qtype === '简答') {
      globalNum += questions.length;
      continue;
    }
    for (const q of questions) {
      globalNum++;
      if (!q.isWrong) continue;
      const typeLabel = qtype === '填空' ? '填空题' : '题目';
      output += `${globalNum}. (${typeLabel})${formatRichForText(questionContent(q))}\n`;
      output += `   我的答案: ${q.myAnswer || '无'}\n`;
      output += `   正确答案: ${formatRichForText(answerContent(q)) || '（未找到答案）'}\n\n`;
    }
  }
  return output;
}

export function formatWrongQuestionsTXT(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = '';
  output += '\n\n\n';
  output += '========================================\n';
  output += '              错题汇总\n';
  output += '========================================\n\n';
  output += wrongQuestionsBodyTXT(results, typeOrder);
  return output.replace(/\n+$/, '');
}

// ==================== 多章节分段（Word 的章节分节在 TXT 上的对应实现） ====================
//
// 每章渲染成一个区块：章标题 → 各小节题目 → **该章的答案汇总** → **该章的错题汇总**。
// 答案与错题此前堆在文件末尾，做完一章的题要翻到最后才能对；现在读完一章就能对，
// 且答案编号天然与该小节的题号一一对应。
//
// 单章节（含「按章节拆分」产出的单章文件）时章内即文件末尾，各分支退化成改动前的原函数，
// 输出与改动前逐字一致。

/** 章内的「答案汇总 / 错题汇总」页头 —— 与原来文件末尾的汇总页同一样式 */
function summaryBanner(text: string): string {
  const line = '='.repeat(40);
  return `${blankLines(2)}${line}\n${' '.repeat(14)}${text}\n${line}${blankLines(1)}`;
}

/**
 * 章内的答案 / 错题区：逐小节一个小标题，题号按小节从 1 重排，便于回正文对照。
 * `skipEmpty` 为真时把没有内容的行整条去掉（错题区用，一道错题都没有的小节连标题一起省）。
 */
function chapterSummary(
  group: SectionGroup<ExportSection>,
  render: (section: ExportSection) => string,
  skipEmpty: boolean,
): string {
  const entries = group.entries
    .map((entry) => ({ heading: entry.heading, value: render(entry.value) }))
    .filter((entry) => !skipEmpty || entry.value.trim().length > 0);
  return renderSectionGroups([{ heading: null, entries }], READABLE_GAPS);
}

/** 章标题 → 各小节题目（题号与题型序号在各小节内从 1 / 一 重新数起） */
function chapterBodyText(group: SectionGroup<ExportSection>): string {
  return renderSectionGroups(
    [
      {
        heading: group.heading,
        entries: group.entries.map((entry) => ({
          heading: entry.heading,
          value: formatOutput(entry.value.results, entry.value.typeOrder),
        })),
      },
    ],
    READABLE_GAPS,
  );
}

/** 一章的完整内容：题目 → 答案 → 错题 */
function formatChapterText(
  group: SectionGroup<ExportSection>,
  options: { withAnswers: boolean; withWrong: boolean },
): string {
  let output = chapterBodyText(group);

  if (options.withAnswers) {
    output += summaryBanner('答案汇总');
    output += chapterSummary(
      group,
      (section) => formatAnswersTXT(section.results, section.typeOrder),
      false,
    );
  }
  if (options.withWrong) {
    const wrong = chapterSummary(
      group,
      (section) => wrongQuestionsBodyTXT(section.results, section.typeOrder).replace(/\n+$/, ''),
      true,
    );
    // 一道错题都没有的章连页头一起省掉
    if (wrong) output += summaryBanner('错题汇总') + wrong;
  }
  return output;
}

/**
 * 正文：多章节时插「第 N 章 章名」+「1.1 小节名」两级标题。
 *
 * 只有「单章**且**这一节没有小节标题」时才走原 formatter（扁平目录 / 老数据，输出与改动前逐字一致）；
 * 带小节标题的单节文件（例如只含「1.3 本章测试」的章）仍走章块渲染，标题不能丢。
 */
export function formatSectionsText(sections: readonly ExportSection[]): string {
  const only = sections[0];
  if (sections.length <= 1 && !only?.heading) {
    return only ? formatOutput(only.results, only.typeOrder) : '';
  }
  return chapterBlocks(sections)
    .map((group) => chapterBodyText(group))
    .join(blankLines(READABLE_GAPS.chapter));
}

/**
 * 完整输出：逐章「题目 → 答案 → 错题」。
 *
 * 单章节时答案与错题本来就在题目之后，各分支退化成改动前的原函数，输出逐字不变。
 */
export function formatSectionsOutput(
  sections: readonly ExportSection[],
  options: { withAnswers: boolean; withWrong: boolean },
): string {
  const only = sections[0];
  if (sections.length <= 1 && !only?.heading) {
    if (!only) return '';
    let output = options.withAnswers
      ? formatOutputWithAnswers(only.results, only.typeOrder)
      : formatOutput(only.results, only.typeOrder);
    if (options.withWrong) output += formatWrongQuestionsTXT(only.results, only.typeOrder);
    return output;
  }
  return chapterBlocks(sections)
    .map((group) => formatChapterText(group, options))
    .join(blankLines(READABLE_GAPS.chapter));
}
