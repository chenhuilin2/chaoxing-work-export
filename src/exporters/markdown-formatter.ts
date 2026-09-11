// Markdown 格式化：以下函数逐行移植自主脚本 chaoxing-work-export.user.js，保持输出完全一致。
import {
  type LegacyResults,
  type LegacyTypeKey,
  answerContent,
  formatRichForMD,
  optionContent,
  questionContent,
} from './legacy-bridge';

// ==================== Markdown 格式化（富文本） ====================
export function formatOutputMD(
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

    const label = typeLabels[qtype] ?? qtype;
    const num = typeNumbers[sectionIdx] || sectionIdx + 1;
    output += `### ${num}、${label}（共${questions.length}题）\n\n`;

    for (const q of questions) {
      globalNum++;
      output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
      output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
      if (q.options && q.options.length > 0) {
        for (const opt of q.options) {
          output += `- ${opt.letter}. ${formatRichForMD(optionContent(opt))}\n`;
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
        const parts = answer
          .split('；')
          .map((p) => p.trim().replace(/^\(\d+\)\s*/, ''));
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

export function formatWrongQuestionsMD(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = '\n\n\n---\n\n';
  output += '## 错题汇总\n\n';

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
      output += `**${globalNum}.** ${formatRichForMD(questionContent(q))}\n\n`;
      output += `- 我的答案: ${q.myAnswer || '无'}\n`;
      output += `- 正确答案: ${formatRichForMD(answerContent(q)) || '（未找到答案）'}\n\n`;
    }
  }
  return output.replace(/\n+$/, '');
}
