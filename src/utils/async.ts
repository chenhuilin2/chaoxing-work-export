export function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

export async function poll<T>(
  producer: () => T | Promise<T>,
  accept: (value: T) => boolean,
  options: { readonly timeoutMs: number; readonly intervalMs: number },
): Promise<T | null> {
  const deadline = Date.now() + options.timeoutMs;
  let lastValue: T | null = null;
  while (Date.now() < deadline) {
    lastValue = await producer();
    if (accept(lastValue)) return lastValue;
    await delay(options.intervalMs);
  }
  return lastValue !== null && accept(lastValue) ? lastValue : null;
}

export function debounce<TArguments extends readonly unknown[]>(
  callback: (...arguments_: TArguments) => void,
  waitMs: number,
): (...arguments_: TArguments) => void {
  let timer: number | null = null;
  return (...arguments_: TArguments) => {
    if (timer !== null) window.clearTimeout(timer);
    timer = window.setTimeout(() => {
      timer = null;
      callback(...arguments_);
    }, waitMs);
  };
}
