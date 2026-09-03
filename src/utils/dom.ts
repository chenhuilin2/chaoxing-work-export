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

export function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false;
  return Boolean(target.closest('input, textarea, select, [contenteditable="true"]'));
}
