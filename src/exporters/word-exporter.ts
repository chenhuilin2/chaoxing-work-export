// Word 文档生成：以下函数逐行移植自主脚本 chaoxing-work-export.user.js，保持输出完全一致。
import type { RichContent } from '../domain/question';
import {
  type LegacyResults,
  type LegacyTypeKey,
  answerContent,
  formatRichForText,
  optionContent,
  questionContent,
} from './legacy-bridge';

// 图片加载失败的全局计数（与主脚本一致，通过 window 传递）
declare global {
  interface Window {
    __xxt_failed_image_count?: number;
  }
}

export const WORD_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

// ==================== Word 文档生成（富文本） ====================
async function fetchImageAsset(
  url: string,
): Promise<{ data: string; type: string; width: number; height: number } | null> {
  if (!url) return null;
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10000);
    const resp = await fetch(url, { mode: 'cors', signal: controller.signal });
    clearTimeout(timeoutId);
    if (!resp.ok) return null;
    const blob = await resp.blob();
    if (!blob.type.startsWith('image/')) return null;
    const data = await new Promise<string | null>((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : null);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
    if (!data) return null;
    const size = await new Promise<{ width: number; height: number }>((resolve) => {
      const img = new Image();
      img.onload = () =>
        resolve({ width: img.naturalWidth || 300, height: img.naturalHeight || 200 });
      img.onerror = () => resolve({ width: 300, height: 200 });
      img.src = data;
    });
    const typeMap: Record<string, string> = {
      'image/png': 'png',
      'image/jpeg': 'jpg',
      'image/jpg': 'jpg',
      'image/gif': 'gif',
      'image/bmp': 'bmp',
    };
    const type = typeMap[blob.type] || 'png';
    return { data, type, width: size.width, height: size.height };
  } catch {
    return null;
  }
}

async function buildImageRun(url: string): Promise<any> {
  if (!url) return null;
  const base64 = await fetchImageAsset(url);
  if (base64) {
    // 按比例限制最大宽高，避免固定尺寸导致变形
    const maxWidth = 420;
    const maxHeight = 260;
    const scale = Math.min(maxWidth / base64.width, maxHeight / base64.height, 1);
    return new docx.ImageRun({
      data: base64,
      transformation: {
        width: Math.round(base64.width * scale),
        height: Math.round(base64.height * scale),
      },
      type: base64.type,
    });
  }
  // 图片加载失败，记录计数
  window.__xxt_failed_image_count = (window.__xxt_failed_image_count || 0) + 1;
  return null;
}

async function buildRichRuns(content: RichContent, prefix = ''): Promise<any[]> {
  const { TextRun } = docx;
  const runs: any[] = [];
  if (prefix) runs.push(new TextRun({ text: prefix, font: 'Microsoft YaHei', size: 22 }));
  for (const part of content || []) {
    if (part.type === 'text') {
      const normalized = part.text.replace(/\n+/g, ' ');
      if (normalized) {
        runs.push(
          new TextRun({
            text: normalized,
            font: 'Microsoft YaHei',
            size: 22,
            bold: part.bold || false,
            italics: part.italic || false,
            subScript: part.subScript || false,
            superScript: part.superScript || false,
          }),
        );
      }
    } else if (part.type === 'image') {
      if (runs.length > 0)
        runs.push(new TextRun({ text: ' ', font: 'Microsoft YaHei', size: 22 }));
      const imgRun = await buildImageRun(part.url);
      if (imgRun) runs.push(imgRun);
      runs.push(new TextRun({ text: ' ', font: 'Microsoft YaHei', size: 22 }));
    } else if (part.type === 'break') {
      runs.push(new TextRun({ text: '\n', break: 1, font: 'Microsoft YaHei', size: 22 }));
    }
  }
  return runs.length ? runs : [new TextRun({ text: prefix, font: 'Microsoft YaHei', size: 22 })];
}

// 将富文本内容构建为带段后间距的 Paragraph 数组（Word 导出用）
async function buildRichParagraphs(
  content: RichContent,
  prefix = '',
  spacing = 0,
): Promise<any[]> {
  const { Paragraph } = docx;
  const runs = await buildRichRuns(content, prefix);
  return [
    new Paragraph({
      children: runs,
      spacing: { after: spacing },
    }),
  ];
}

// 清洗答案文本：去除“正确答案:/我的答案:”等标签，避免组合出现
function cleanAnswerText(text: string): string {
  if (!text) return '';
  return text
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, '\n')
    .replace(/(?:正确答案|参考答案)[:：]?\s*/g, '')
    .replace(/(?:我的答案|你的答案|学生答案)[:：]?\s*/g, '')
    .replace(/^[：:\s]+/, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export async function generateWordBlob(
  results: LegacyResults,
  typeOrder: readonly LegacyTypeKey[],
  title: string,
  withAnswers: boolean,
  withWrong: boolean,
  bankImport: boolean,
): Promise<Blob> {
  const { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, HeadingLevel, PageBreak } =
    docx;

  const FONT = 'Microsoft YaHei';
  const COLOR = {
    title: '1F2937',
    muted: '6B7280',
    border: 'E5E7EB',
    answer: '00A870',
    wrong: 'DC2626',
    analysisBg: 'F7F9FC',
    type: '64748B',
  };

  const typeHeaders: Record<string, string> = {
    单选: '一、单项选择题',
    多选: '二、多项选择题',
    填空: '三、填空题',
    判断: '四、判断题',
    简答: '五、简答题',
  };

  const bankTypeLabels: Record<string, string> = {
    单选: '【单选题】',
    多选: '【多选题】',
    填空: '【填空题】',
    判断: '【判断题】',
    简答: '【简答题】',
  };

  // 题库导入格式：生成学习通智能导入兼容的 Word 文档
  if (bankImport) {
    const children: any[] = [];
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: title || '题库导入',
            font: FONT,
            size: 32,
            bold: true,
            color: COLOR.title,
          }),
        ],
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
      }),
    );

    let qNum = 0;

    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;

      const prefix = bankTypeLabels[qtype];

      for (const q of questions) {
        qNum++;
        // 题干（题号 + 题型标签 + 题干内容），括号归一化
        const stemText = formatRichForText(questionContent(q));
        const stem =
          prefix + stemText.replace(/\(\s{2,}\)/g, '（ ）').replace(/（\s{2,}）/g, '（ ）');
        children.push(
          new Paragraph({
            children: [new TextRun({ text: `${qNum}.${stem}`, font: FONT, size: 22 })],
            spacing: { after: 40 },
          }),
        );
        // 题干中的图片
        const stemContent = questionContent(q);
        for (const part of stemContent) {
          if (part.type === 'image') {
            const imgRun = await buildImageRun(part.url);
            if (imgRun) {
              children.push(
                new Paragraph({
                  children: [imgRun],
                  spacing: { after: 80 },
                }),
              );
            }
          }
        }

        // 选项
        const options = q.options || [];
        for (const opt of options) {
          const runs = await buildRichRuns(optionContent(opt), `${opt.letter}. `);
          children.push(
            new Paragraph({
              children: runs,
              indent: { left: 400, hanging: 200 },
              spacing: { after: 40 },
            }),
          );
        }

        // 答案
        const answer = formatRichForText(answerContent(q)).trim();
        if (answer) {
          // 格式化答案内容
          let formattedAnswerParts: RichContent | null = null;
          if (qtype === '多选') {
            const parts = answer.replace(/\s+/g, '').split('');
            formattedAnswerParts = [{ type: 'text', text: parts.join('，') }];
          } else if (qtype === '判断') {
            if (/^[√✓Tt]|正确|True|TRUE/.test(answer)) {
              formattedAnswerParts = [{ type: 'text', text: '对' }];
            } else {
              formattedAnswerParts = [{ type: 'text', text: '错' }];
            }
          }
          children.push(
            ...(await buildRichParagraphs(
              formattedAnswerParts || [{ type: 'text', text: answer }],
              '答案：',
              120,
            )),
          );
        } else {
          children.push(
            new Paragraph({
              children: [new TextRun({ text: '', font: FONT, size: 22 })],
              spacing: { after: 120 },
            }),
          );
        }
      }
    }

    const doc = new Document({
      styles: {
        paragraphStyles: [
          {
            id: 'Normal',
            name: 'Normal',
            run: { font: FONT, size: 22, color: COLOR.title },
            paragraph: { spacing: { line: 330, lineRule: 'auto' } },
          },
        ],
      },
      sections: [
        {
          properties: {
            page: {
              size: { width: 11906, height: 16838 },
              margin: {
                top: convertMillimetersToTwip(20),
                bottom: convertMillimetersToTwip(20),
                left: convertMillimetersToTwip(25),
                right: convertMillimetersToTwip(25),
              },
            },
          },
          children,
        },
      ],
    });

    return await Packer.toBlob(doc);
  }

  const children: any[] = [];

  // 标题
  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: title || '试卷', font: FONT, size: 32, bold: true, color: '2563EB' }),
      ],
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 120 },
    }),
  );

  // 统计摘要
  let totalQ = 0;
  const typeCount: Record<string, number> = {};
  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;
    typeCount[qtype] = questions.length;
    totalQ += questions.length;
  }
  const summary = Object.entries(typeCount)
    .map(([type, count]) => `${type} ${count} 道`)
    .join(' / ');
  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: `共 ${totalQ} 道题`, font: FONT, size: 20, color: COLOR.muted }),
        ...(summary
          ? [new TextRun({ text: ` · ${summary}`, font: FONT, size: 20, color: COLOR.muted })]
          : []),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 260 },
    }),
  );

  let qNum = 0;

  for (const qtype of typeOrder) {
    const questions = results[qtype];
    if (!questions || questions.length === 0) continue;

    const header = typeHeaders[qtype] || qtype;
    const count = questions.length;

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${header}（本大题共${count}小题）`,
            font: FONT,
            size: 28,
            bold: true,
            color: COLOR.title,
          }),
        ],
        spacing: { after: 120 },
      }),
    );

    for (const q of questions) {
      qNum++;
      const stemContent = questionContent(q);
      const typeMetaRun = q.typeMeta
        ? new TextRun({ text: `${q.typeMeta} `, font: FONT, size: 22, color: COLOR.type })
        : null;

      if (qtype === '单选' || qtype === '多选') {
        // 题目块
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `${qNum}. `,
                font: FONT,
                size: 22,
                bold: true,
                color: COLOR.title,
              }),
              ...(typeMetaRun ? [typeMetaRun] : []),
              ...(await buildRichRuns(stemContent)),
            ],
            spacing: { before: 220, after: 100 },
          }),
        );
        const options = q.options || [];
        if (options.length > 0) {
          // 单选/多选选项统一一列多行排列
          for (const opt of options) {
            children.push(
              new Paragraph({
                children: await buildRichRuns(optionContent(opt), `${opt.letter}. `),
                indent: { left: 620, hanging: 220 },
                spacing: { before: 30, after: 30 },
              }),
            );
          }
          children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
        }
      } else if (qtype === '填空') {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `${qNum}. `,
                font: FONT,
                size: 22,
                bold: true,
                color: COLOR.title,
              }),
              ...(await buildRichRuns(stemContent)),
            ],
            spacing: { before: 220, after: 40 },
          }),
        );
        children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
      } else if (qtype === '判断') {
        const judgeRuns = await buildRichRuns(stemContent);
        judgeRuns.push(new TextRun({ text: '（  ）', font: FONT, size: 22 }));
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `${qNum}. `,
                font: FONT,
                size: 22,
                bold: true,
                color: COLOR.title,
              }),
              ...(typeMetaRun ? [typeMetaRun] : []),
              ...judgeRuns,
            ],
            spacing: { before: 220, after: 120 },
          }),
        );
      } else if (qtype === '简答') {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `${qNum}. `,
                font: FONT,
                size: 22,
                bold: true,
                color: COLOR.title,
              }),
              ...(typeMetaRun ? [typeMetaRun] : []),
              ...(await buildRichRuns(stemContent)),
            ],
            spacing: { before: 220, after: 40 },
          }),
        );
        for (let i = 0; i < 8; i++) {
          children.push(
            new Paragraph({
              children: [new TextRun({ text: '', font: FONT, size: 22 })],
              spacing: { after: 40 },
            }),
          );
        }
        children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
      }
    }
    // 判断题与简答题之间补一行空行
    if (qtype === '判断') {
      children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
    }
  }

  // 答案页
  if (withAnswers) {
    children.push(
      new Paragraph({
        children: [new PageBreak()],
        spacing: { after: 0 },
      }),
    );
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: '正确答案',
            font: FONT,
            size: 32,
            bold: true,
            color: COLOR.title,
          }),
        ],
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 240 },
      }),
    );
    let aNum = 0;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions || questions.length === 0) continue;
      const header = typeHeaders[qtype] || qtype;
      children.push(
        new Paragraph({
          children: [new TextRun({ text: header, font: FONT, size: 28, bold: true, color: COLOR.title })],
          spacing: { after: 120 },
        }),
      );
      for (const q of questions) {
        aNum++;
        // 清洗答案文本，去除“正确答案:/我的答案:”等标签组合
        const answerContentArr = answerContent(q);
        const cleanedContent = answerContentArr.map((part) => {
          if (part.type === 'text') {
            return { ...part, text: cleanAnswerText(part.text) };
          }
          return part;
        });
        const answerRuns = await buildRichRuns(cleanedContent);
        if (answerRuns.length === 0) {
          answerRuns.push(
            new TextRun({ text: '（未找到答案）', font: FONT, size: 22, color: COLOR.muted }),
          );
        } else {
          // 将答案文本改为绿色加粗
          answerRuns.forEach((run) => {
            if (run.font) run.font = FONT;
            run.size = 22;
            run.bold = true;
            run.color = COLOR.answer;
          });
        }
        children.push(
          new Paragraph({
            children: [new TextRun({ text: `${aNum}. `, font: FONT, size: 22, color: COLOR.title }), ...answerRuns],
            spacing: { before: 90, after: 70 },
            indent: { left: 420 },
          }),
        );
        // 简答题答案常有多行，之间空一行便于区分
        if (qtype === '简答') {
          children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
        }
      }
    }
  }

  // 错题汇总（Word 试卷）
  if (withWrong) {
    let hasWrong = false;
    for (const qtype of typeOrder) {
      const questions = results[qtype];
      if (!questions) continue;
      for (const q of questions) {
        if (q.isWrong) {
          hasWrong = true;
          break;
        }
      }
      if (hasWrong) break;
    }

    if (hasWrong) {
      children.push(
        new Paragraph({
          children: [new PageBreak()],
          spacing: { after: 0 },
        }),
      );
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: '错题汇总',
              font: FONT,
              size: 32,
              bold: true,
              color: COLOR.title,
            }),
          ],
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 240 },
        }),
      );

      let globalNum = 0;
      for (const qtype of typeOrder) {
        const questions = results[qtype];
        if (!questions || questions.length === 0) continue;
        if (qtype === '简答') {
          globalNum += questions.length;
          continue;
        }

        const header = typeHeaders[qtype] || qtype;
        let sectionHasWrong = false;
        for (const q of questions) {
          if (q.isWrong) {
            sectionHasWrong = true;
            break;
          }
        }
        if (!sectionHasWrong) {
          globalNum += questions.length;
          continue;
        }

        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: header, font: FONT, size: 28, bold: true, color: COLOR.title }),
            ],
            spacing: { after: 120 },
          }),
        );

        for (const q of questions) {
          globalNum++;
          if (!q.isWrong) continue;
          children.push(
            new Paragraph({
              children: await buildRichRuns(questionContent(q), `${globalNum}. `),
              spacing: { before: 160, after: 40 },
            }),
          );
          const options = q.options || [];
          if (options.length > 0) {
            for (const opt of options) {
              children.push(
                new Paragraph({
                  children: await buildRichRuns(optionContent(opt), `${opt.letter}. `),
                  indent: { left: 620, hanging: 220 },
                  spacing: { before: 30, after: 30 },
                }),
              );
            }
          }
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: `我的答案: ${q.myAnswer || '无'}`,
                  font: FONT,
                  size: 22,
                  color: COLOR.wrong,
                }),
              ],
              spacing: { before: 60, after: 40 },
              indent: { left: 420 },
            }),
          );
          const answerContentArr = answerContent(q);
          const cleanedContent = answerContentArr.map((part) => {
            if (part.type === 'text') {
              return { ...part, text: cleanAnswerText(part.text) };
            }
            return part;
          });
          const correctRuns = await buildRichRuns(cleanedContent);
          if (correctRuns.length === 0) {
            correctRuns.push(
              new TextRun({ text: '（未找到答案）', font: FONT, size: 22, color: COLOR.muted }),
            );
          } else {
            correctRuns.forEach((run) => {
              if (run.font) run.font = FONT;
              run.size = 22;
              run.bold = true;
              run.color = COLOR.answer;
            });
          }
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: '正确答案：',
                  font: FONT,
                  size: 22,
                  bold: true,
                  color: COLOR.answer,
                }),
                ...correctRuns,
              ],
              spacing: { before: 40, after: 120 },
              indent: { left: 420 },
            }),
          );
          // 每道题之间空一行
          children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
        }
        // 每个类别之间空一行
        children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
      }
    }
  }

  const doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: 'Normal',
          name: 'Normal',
          run: { font: FONT, size: 22, color: COLOR.title },
          paragraph: { spacing: { line: 330, lineRule: 'auto' } },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: {
              top: 1134,
              right: 1134,
              bottom: 1134,
              left: 1134,
            },
          },
        },
        children,
      },
    ],
  });

  return await Packer.toBlob(doc);
}
