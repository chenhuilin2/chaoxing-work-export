// Markdown 格式化：以下函数逐行移植自主脚本 chaoxing-work-export.user.js，保持输出完全一致。
//
// 多章节导出时按「章 → 小节」两级分段，且每章后面紧跟该章的答案与错题（见文件末尾的分段实现）。
// 下面这些单章 formatter 本来就以「传入的结果集」为范围从 1 开始编号，按小节传入即自动按小节重排。
import {
  type LegacyResults,
  type LegacyTypeKey,
  answerContent,
  formatRichForMD,
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

// 标题级别：章占 ##、小节占 ###；题型在「本节有小节标题」时降到 ####，保持层级单调
const MD_LEVEL = {
  chapter: 2,
  section: 3,
  type: 3,
  typeUnderSection: 4,
} as const;

// ==================== Markdown 格式化（富文本） ====================
/**
 * @param typeLevel 题型大标题的级别（默认 3；正文里插了「### 1.1 小节名」时降到 4）
 */
export function formatOutputMD(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
  typeLevel: number = MD_LEVEL.type,
): string {
  const typeLabels: Record<string, string> = {
    单选: '单选题',
    多选: '多选题',
    填空: '填空题',
    判断: '判断题',
    简答: '简答题',
  };
  const typeNumbers = ['一', '二', '三', '四', '五', '六'];
  const prefix = '#'.repeat(typeLevel);

  let output = '';
  let globalNum = 0;
  let sectionIdx = 0;

  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;

    const label = typeLabels[qtype] ?? qtype;
    const num = typeNumbers[sectionIdx] || sectionIdx + 1;
    output += `${prefix} ${num}、${label}（共${questions.length}题）\n\n`;

    for (const q of questions) {
      globalNum++;
      output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
      if (q.options && q.options.length > 0) {
        for (const opt of q.options) {
          output += `- ${opt.letter}. ${formatRichForMD(optionContent(opt))}\n`;
        }
        output += '\n';
      } else {
        output += '\n';
      }
    }
    sectionIdx++;
  }
  return output.trim();
}

export function formatAnswersMD(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = '';
  let globalNum = 0;
  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;
    const typeLabels: Record<string, string> = {
      单选: '单选题',
      多选: '多选题',
      填空: '填空题',
      判断: '判断题',
      简答: '简答题',
    };
    output += `**${typeLabels[qtype] ?? qtype}**\n\n`;
    for (const q of questions) {
      globalNum++;
      const answer = formatRichForMD(answerContent(q)) || '（未找到答案）';
      if (qtype === '填空' && answer.includes('；')) {
        const parts = answer.split('；').map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
        output += `${globalNum}.  \n`;
        parts.forEach((part, i) => {
          output += `    (${i + 1}) ${part}  \n`;
        });
        output += '\n';
      } else {
        output += `${globalNum}. ${answer}  \n`;
      }
    }
    output += '\n';
  }
  return output.trim();
}

export function formatOutputWithAnswersMD(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = formatOutputMD(results, typeOrder);
  output += '\n\n---\n\n';
  output += '## 答案汇总\n\n';
  output += formatAnswersMD(results, typeOrder);
  return output.trim();
}

/**
 * 错题正文（不含「## 错题汇总」页头）。
 *
 * 题号以传入的结果集为范围从 1 数起，多章节分段与单章节页头拼接都复用它。
 */
function wrongQuestionsBodyMD(results: LegacyResults, typeOrder: readonly LegacyTypeKey[]): string {
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
      output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
      output += `- 我的答案: ${q.myAnswer || '无'}\n`;
      output += `- 正确答案: ${formatRichForMD(answerContent(q)) || '（未找到答案）'}\n\n`;
    }
  }
  return output;
}

export function formatWrongQuestionsMD(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = '\n\n\n---\n\n';
  output += '## 错题汇总\n\n';
  output += wrongQuestionsBodyMD(results, typeOrder);
  return output.replace(/\n+$/, '');
}

// ==================== 多章节分段（Word 的章节分节在 Markdown 上的对应实现） ====================
//
// 每章渲染成一个区块：章标题 → 各小节题目 → **该章的答案汇总** → **该章的错题汇总**。
// 答案与错题此前集中在文件末尾，现在跟着章走，读完一章就能对答案。
//
// 单章节（含「按章节拆分」产出的单章文件）时章内即文件末尾，各分支退化成改动前的原函数，
// 输出与改动前逐字一致。

/**
 * 给两级标题带上 Markdown 标记。
 * 只有一级（老数据 / 扁平目录）时章标题仍是 `##`、小节标题为空，输出与改动前一致。
 */
function decorateGroups(groups: readonly SectionGroup<string>[]): SectionGroup<string>[] {
  return groups.map((group) => ({
    heading: group.heading ? `${'#'.repeat(MD_LEVEL.chapter)} ${group.heading}` : null,
    entries: group.entries.map((entry) => ({
      heading: entry.heading ? `${'#'.repeat(MD_LEVEL.section)} ${entry.heading}` : null,
      value: entry.value,
    })),
  }));
}

/**
 * 章内的答案 / 错题页头。
 *
 * 用 `##`（与章标题同级）而不是原来的 `#`：汇总页现在夹在章与章之间，若仍是 `#`，
 * 后面的章标题就会在大纲里挂到「答案汇总」底下；同级时各章才是干净的兄弟节点。
 * `---` 分隔线与单章时的输出保持一致。
 */
function summaryHeading(text: string): string {
  return `${blankLines(1)}---${blankLines(1)}${'#'.repeat(MD_LEVEL.chapter)} ${text}${blankLines(1)}`;
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
  return renderSectionGroups(decorateGroups([{ heading: null, entries }]), READABLE_GAPS);
}

/** 章标题 → 各小节题目（题型头在有小节标题时降到 ####，保持层级单调） */
function chapterBodyMD(group: SectionGroup<ExportSection>): string {
  return renderSectionGroups(
    decorateGroups([
      {
        heading: group.heading,
        entries: group.entries.map((entry) => ({
          heading: entry.heading,
          value: formatOutputMD(
            entry.value.results,
            entry.value.typeOrder,
            entry.heading ? MD_LEVEL.typeUnderSection : MD_LEVEL.type,
          ),
        })),
      },
    ]),
    READABLE_GAPS,
  );
}

/** 一章的完整内容：题目 → 答案 → 错题 */
function formatChapterMD(
  group: SectionGroup<ExportSection>,
  options: { withAnswers: boolean; withWrong: boolean },
): string {
  let output = chapterBodyMD(group);

  if (options.withAnswers) {
    output += summaryHeading('答案汇总');
    output += chapterSummary(
      group,
      (section) => formatAnswersMD(section.results, section.typeOrder),
      false,
    );
  }
  if (options.withWrong) {
    const wrong = chapterSummary(
      group,
      (section) => wrongQuestionsBodyMD(section.results, section.typeOrder).replace(/\n+$/, ''),
      true,
    );
    // 一道错题都没有的章连页头一起省掉
    if (wrong) output += summaryHeading('错题汇总') + wrong;
  }
  return output;
}

/**
 * 正文：多章节时插「## 第 N 章 章名」+「### 1.1 小节名」两级标题。
 *
 * 只有「单章**且**这一节没有小节标题」时才走原 formatter（扁平目录 / 老数据，输出与改动前逐字一致）；
 * 带小节标题的单节文件（例如只含「1.3 本章测试」的章）仍走章块渲染，标题不能丢。
 */
export function formatSectionsMD(sections: readonly ExportSection[]): string {
  const only = sections[0];
  if (sections.length <= 1 && !only?.heading) {
    return only ? formatOutputMD(only.results, only.typeOrder) : '';
  }
  return chapterBlocks(sections)
    .map((group) => chapterBodyMD(group))
    .join(blankLines(READABLE_GAPS.chapter));
}

/**
 * 完整输出：逐章「题目 → 答案 → 错题」。
 *
 * 单章节时答案与错题本来就在题目之后，各分支退化成改动前的原函数，输出逐字不变。
 */
export function formatSectionsOutputMD(
  sections: readonly ExportSection[],
  options: { withAnswers: boolean; withWrong: boolean },
): string {
  const only = sections[0];
  if (sections.length <= 1 && !only?.heading) {
    if (!only) return '';
    let output = options.withAnswers
      ? formatOutputWithAnswersMD(only.results, only.typeOrder)
      : formatOutputMD(only.results, only.typeOrder);
    if (options.withWrong) output += formatWrongQuestionsMD(only.results, only.typeOrder);
    return output;
  }
  return chapterBlocks(sections)
    .map((group) => formatChapterMD(group, options))
    .join(blankLines(READABLE_GAPS.chapter));
}
