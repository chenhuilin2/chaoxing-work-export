export function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
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
