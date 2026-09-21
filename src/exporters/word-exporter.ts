// Word 文档生成：以下函数逐行移植自主脚本 chaoxing-work-export.user.js，保持输出完全一致。
// 多章节支持：正文按 sections 顺序渲染，sections 多于一个时每章前插章节标题并分页
//（按章节拆分成多个文件时，每份文档只有一个 section，因此不插章节标题）。
import type { RichContent } from '../domain/question';
import { answerContent, formatRichForText, optionContent, questionContent } from './legacy-bridge';
import type { LegacyTypeKey } from './legacy-bridge';
import type { WordSection } from './word-plan';

// 图片加载失败的全局计数（与主脚本一致，通过 window 传递）
declare global {
  interface Window {
    __xxt_failed_image_count?: number;
  }
}

export const WORD_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** Word 渲染开关（文件划分由 word-plan.ts 决定） */
export interface WordRenderOptions {
  readonly withAnswers: boolean;
  readonly withWrong: boolean;
  readonly bankImport: boolean;
}

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
      if (runs.length > 0) runs.push(new TextRun({ text: ' ', font: 'Microsoft YaHei', size: 22 }));
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
async function buildRichParagraphs(content: RichContent, prefix = '', spacing = 0): Promise<any[]> {
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

const FONT = 'Microsoft YaHei';
// 题型大题头字号（半磅）：比章节标题小一号（章节标题 28 = 14pt → 题型 24 = 12pt），与正文 22 拉开层次
const SIZE_TYPE_HEADER = 24;
const COLOR = {
  title: '1F2937',
  muted: '6B7280',
  border: 'E5E7EB',
  answer: '00A870',
  wrong: 'DC2626',
  analysisBg: 'F7F9FC',
  type: '64748B',
};

// 题型名（不含序号）：序号由 buildTypeHeaders 按本章实际出现的题型顺序现场排
const TYPE_NAMES: Record<string, string> = {
  单选: '单项选择题',
  多选: '多项选择题',
  填空: '填空题',
  判断: '判断题',
  简答: '简答题',
};

const CN_NUMERALS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'] as const;

/** 大题头的中文序号（超出表长时退化为阿拉伯数字） */
function ordinalLabel(index: number): string {
  return CN_NUMERALS[index] ?? String(index + 1);
}

/**
 * 本章的「题型 → 大题头」映射。
 *
 * 序号按本章**实际有题目**的题型从「一」重新排：此前用全局固定的题型序号，
 * 某章没有单选题时多选题就会顶着「二、多项选择题」的跳号。
 */
function buildTypeHeaders(section: WordSection): Map<LegacyTypeKey, string> {
  const headers = new Map<LegacyTypeKey, string>();
  let ordinal = 0;
  for (const qtype of section.typeOrder) {
    const questions = section.results[qtype];
    if (!questions || questions.length === 0) continue;
    headers.set(qtype, `${ordinalLabel(ordinal)}、${TYPE_NAMES[qtype] ?? qtype}`);
    ordinal += 1;
  }
  return headers;
}

const BANK_TYPE_LABELS: Record<string, string> = {
  单选: '【单选题】',
  多选: '【多选题】',
  填空: '【填空题】',
  判断: '【判断题】',
  简答: '【简答题】',
};

/** 题库导入格式的页面设置：A4 纵向，页边距比试卷略宽 */
function bankImportSections(children: any[]): any[] {
  const { convertMillimetersToTwip } = docx;
  return [
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
  ];
}

/** 试卷正文的章节大标题：多章节单文件时用来分节（除首章外先分页） */
function chapterHeadingParagraph(text: string, pageBreak: boolean): any[] {
  const { Paragraph, TextRun, AlignmentType, HeadingLevel, PageBreak } = docx;
  const parts: any[] = [];
  if (pageBreak) parts.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));
  parts.push(
    new Paragraph({
      children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: COLOR.title })],
      heading: HeadingLevel.HEADING_2,
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 200 },
    }),
  );
  return parts;
}

/** 答案页 / 错题页里的章节小标题（不强制分页，与上级大标题区分） */
function chapterLabelParagraph(text: string): any {
  const { Paragraph, TextRun } = docx;
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: COLOR.title })],
    spacing: { before: 200, after: 120 },
  });
}

/** 题库导入格式：生成学习通智能导入兼容的 Word 文档 */
async function buildBankImportBlob(title: string, sections: readonly WordSection[]): Promise<Blob> {
  const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = docx;
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

  for (const section of sections) {
    for (const qtype of section.typeOrder) {
      const questions = section.results[qtype];
      if (!questions || questions.length === 0) continue;

      const prefix = BANK_TYPE_LABELS[qtype];

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
    sections: bankImportSections(children),
  });

  return await Packer.toBlob(doc);
}

/**
 * 试卷正文：标题 + 统计摘要 + 各章节题目（按题型分组）。
 *
 * title 是文档大标题（多章节时是「课程 - N 个章节」这类聚合标题），
 * 与 section.title（章节名，多章节时作为正文里的分节标题）不是一回事。
 */
async function buildPaperBody(title: string, sections: readonly WordSection[]): Promise<any[]> {
  const { Paragraph, TextRun, AlignmentType, HeadingLevel } = docx;
  const children: any[] = [];
  const multiSection = sections.length > 1;

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

  // 统计摘要（多章节时跨章节汇总）
  let totalQ = 0;
  const typeCount: Record<string, number> = {};
  for (const section of sections) {
    for (const qtype of section.typeOrder) {
      const questions = section.results[qtype];
      if (!questions || questions.length === 0) continue;
      typeCount[qtype] = (typeCount[qtype] ?? 0) + questions.length;
      totalQ += questions.length;
    }
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

  for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
    const section = sections[sectionIndex];
    if (!section) continue;
    // 多章节单文件：每章前插章节大标题
    if (multiSection) children.push(...chapterHeadingParagraph(section.title, sectionIndex > 0));

    // 题号与题型序号都按章重新开始：每章从第 1 题、从「一、」数起
    const typeHeaders = buildTypeHeaders(section);
    let qNum = 0;

    for (const qtype of section.typeOrder) {
      const questions = section.results[qtype];
      if (!questions || questions.length === 0) continue;

      const header = typeHeaders.get(qtype) ?? TYPE_NAMES[qtype] ?? qtype;
      const count = questions.length;

      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${header}（本大题共${count}小题）`,
              font: FONT,
              size: SIZE_TYPE_HEADER,
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
  }

  return children;
}

/** 答案页：分页后按章节、按题型列出正确答案 */
async function appendAnswerPage(children: any[], sections: readonly WordSection[]): Promise<void> {
  const { Paragraph, TextRun, AlignmentType, HeadingLevel, PageBreak } = docx;
  const multiSection = sections.length > 1;

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

  for (const section of sections) {
    if (multiSection) children.push(chapterLabelParagraph(section.title));
    // 答案编号与题目页一致：按章重新从 1 数起，题型序号同样按本章重排
    const typeHeaders = buildTypeHeaders(section);
    let aNum = 0;
    for (const qtype of section.typeOrder) {
      const questions = section.results[qtype];
      if (!questions || questions.length === 0) continue;
      const header = typeHeaders.get(qtype) ?? TYPE_NAMES[qtype] ?? qtype;
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: header,
              font: FONT,
              size: SIZE_TYPE_HEADER,
              bold: true,
              color: COLOR.title,
            }),
          ],
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
            children: [
              new TextRun({ text: `${aNum}. `, font: FONT, size: 22, color: COLOR.title }),
              ...answerRuns,
            ],
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
}

/** 错题汇总页：分页后列出答错的题目、我的答案与正确答案 */
async function appendWrongPage(children: any[], sections: readonly WordSection[]): Promise<void> {
  const { Paragraph, TextRun, AlignmentType, HeadingLevel, PageBreak } = docx;
  const multiSection = sections.length > 1;

  const hasAnyWrong = sections.some((section) =>
    section.typeOrder.some((qtype) =>
      (section.results[qtype] ?? []).some((question) => question.isWrong),
    ),
  );
  if (!hasAnyWrong) return;

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

  for (const section of sections) {
    if (multiSection) {
      // 该章一道错题都没有时，连章节标题一起省掉（与「无错题的题型不出现」一致）
      const sectionHasWrong = section.typeOrder.some((qtype) =>
        (section.results[qtype] ?? []).some((question) => question.isWrong),
      );
      const before = children.length;
      if (sectionHasWrong) children.push(chapterLabelParagraph(section.title));
      // 章节标题先占位，若无错题则回滚（保持既有「空section不出标题」的写法简单）
      if (!sectionHasWrong && children.length !== before) children.length = before;
    }

    // 错题题号要能对上题目页，因此同样按章重新从 1 数起
    const typeHeaders = buildTypeHeaders(section);
    let globalNum = 0;
    for (const qtype of section.typeOrder) {
      const questions = section.results[qtype];
      if (!questions || questions.length === 0) continue;
      if (qtype === '简答') {
        globalNum += questions.length;
        continue;
      }

      const header = typeHeaders.get(qtype) ?? TYPE_NAMES[qtype] ?? qtype;
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
            new TextRun({
              text: header,
              font: FONT,
              size: SIZE_TYPE_HEADER,
              bold: true,
              color: COLOR.title,
            }),
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

export async function generateWordBlob(
  title: string,
  sections: readonly WordSection[],
  options: WordRenderOptions,
): Promise<Blob> {
  const { Document, Packer } = docx;

  if (options.bankImport) {
    return await buildBankImportBlob(title, sections);
  }

  const children = await buildPaperBody(title, sections);
  if (options.withAnswers) await appendAnswerPage(children, sections);
  if (options.withWrong) await appendWrongPage(children, sections);

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
