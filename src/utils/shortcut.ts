import type { Shortcut } from '../domain/settings';

export function matchesShortcut(event: KeyboardEvent, shortcut: Shortcut): boolean {
  return (
    event.ctrlKey === shortcut.ctrl &&
    event.shiftKey === shortcut.shift &&
    event.altKey === shortcut.alt &&
    event.key.toLowerCase() === shortcut.key.toLowerCase()
  );
}

export function formatShortcut(shortcut: Shortcut): string {
  return [
    shortcut.ctrl ? 'Ctrl' : '',
    shortcut.shift ? 'Shift' : '',
    shortcut.alt ? 'Alt' : '',
    shortcut.key.length === 1 ? shortcut.key.toUpperCase() : shortcut.key,
  ]
    .filter(Boolean)
    .join('+');
}

export function parseShortcut(value: string, fallback: Shortcut): Shortcut {
  const parts = value
    .split('+')
    .map((part) => part.trim())
    .filter(Boolean);
  const lower = parts.map((part) => part.toLowerCase());
  const key = parts.find((part) => !['ctrl', 'control', 'shift', 'alt', 'option'].includes(part.toLowerCase()));
  if (!key) return { ...fallback };
  return {
    ctrl: lower.includes('ctrl') || lower.includes('control'),
    shift: lower.includes('shift'),
    alt: lower.includes('alt') || lower.includes('option'),
    key: key.toLowerCase(),
  };
}
