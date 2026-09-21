// TXT 格式化：以下函数逐行移植自主脚本 chaoxing-work-export.user.js，保持输出完全一致。
import {
  type LegacyResults,
  type LegacyTypeKey,
  answerContent,
  formatRichForText,
  optionContent,
  questionContent,
} from './legacy-bridge';

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

export function formatWrongQuestionsTXT(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
): string {
  let output = '';
  output += '\n\n\n';
  output += '========================================\n';
  output += '              错题汇总\n';
  output += '========================================\n\n';

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
  return output.replace(/\n+$/, '');
}
