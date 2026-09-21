export function textOf(element: Element | null | undefined): string {
  return (element?.textContent ?? '').replace(/\u00a0/g, ' ').trim();
}

export function firstMatch<E extends Element>(
  root: ParentNode,
  selectors: readonly string[],
): E | null {
  for (const selector of selectors) {
    const element = root.querySelector<E>(selector);
    if (element) return element;
  }
  return null;
}

export function allMatches<E extends Element>(
  root: ParentNode,
  selectors: readonly string[],
): E[] {
  const output: E[] = [];
  const seen = new Set<E>();
  for (const selector of selectors) {
    root.querySelectorAll<E>(selector).forEach((element) => {
      if (!seen.has(element)) {
        seen.add(element);
        output.push(element);
      }
    });
  }
  return output;
}

export function parseLeadingNumber(value: string): number | undefined {
  const match = value.match(/^\s*(\d+)\s*[.、．]/u);
  if (!match?.[1]) return undefined;
  const parsed = Number.parseInt(match[1], 10);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export interface AccessibleDocument {
  readonly document: Document;
  /** 相对根文档的嵌套层级，根文档为 0 */
  readonly depth: number;
}

/**
 * 深度优先收集可同步访问的文档，含同源 iframe。
 * 学习通把知识卡片、目录、答题页放在多层 iframe 里，定位 DOM 前必须先跨层收集。
 * 跨域 frame 会抛错并被跳过，交由 FrameBridge 通过 postMessage 处理。
 */
export function collectAccessibleDocuments(
  root: Document = document,
  maxDepth = 5,
): AccessibleDocument[] {
  const output: AccessibleDocument[] = [];
  const visited = new Set<Document>();

  const visit = (current: Document, depth: number): void => {
    if (visited.has(current) || depth > maxDepth) return;
    visited.add(current);
    output.push({ document: current, depth });
    current.querySelectorAll<HTMLIFrameElement>('iframe').forEach((frame) => {
      try {
        if (frame.contentDocument) visit(frame.contentDocument, depth + 1);
      } catch {
        // 跨域 frame 无法同步访问，交由 FrameBridge 处理
      }
    });
  };

  visit(root, 0);
  return output;
}

export function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false;
  return Boolean(target.closest('input, textarea, select, [contenteditable="true"]'));
}
