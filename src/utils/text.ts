export function normalizeWhitespace(value: string): string {
  return value
    .replace(/\u00a0/g, ' ')
    .replace(/[\t\r\f ]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function normalizeInlineWhitespace(value: string): string {
  return value.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

export function sanitizeFilename(value: string, fallback = '学习通题目'): string {
  const sanitized = normalizeInlineWhitespace(value)
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/[. ]+$/g, '')
    .slice(0, 120);
  return sanitized || fallback;
}

export function normalizeAnswer(value: string): string {
  const normalized = normalizeInlineWhitespace(value)
    .replace(/^(?:正确答案|参考答案|答案|我的答案|你的答案|学生答案)\s*[:：]?\s*/u, '')
    .replace(/[，、;；\s]+/g, '')
    .toUpperCase();

  if (/^(?:正确|对|TRUE|T|√|✓)$/u.test(normalized)) return 'TRUE';
  if (/^(?:错误|错|FALSE|F|×|✕|X)$/u.test(normalized)) return 'FALSE';
  return normalized;
}

export function escapeMarkdownAlt(value: string): string {
  return (value || '图片').replace(/[\[\]\r\n]/g, ' ').trim() || '图片';
}

export function escapeMarkdownUrl(value: string): string {
  return value.replace(/[()\\]/g, (character) => {
    return `%${character.charCodeAt(0).toString(16).toUpperCase()}`;
  });
}

export function humanizeError(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return typeof error === 'string' ? error : '发生未知错误';
}
