export function shuffleCopy<T>(
  values: readonly T[],
  random: () => number = Math.random,
): T[] {
  const output = [...values];
  for (let index = output.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(random() * (index + 1));
    const temporary = output[index];
    output[index] = output[swapIndex] as T;
    output[swapIndex] = temporary as T;
  }
  return output;
}

export function uniqueBy<T>(values: readonly T[], keyOf: (value: T) => string): T[] {
  const seen = new Set<string>();
  const output: T[] = [];
  for (const value of values) {
    const key = keyOf(value);
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(value);
  }
  return output;
}
