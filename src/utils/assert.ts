export function assertDefined<T>(value: T | null | undefined, message: string): T {
  if (value === null || value === undefined) throw new Error(message);
  return value;
}

export function queryRequired<E extends Element>(
  root: ParentNode,
  selector: string,
): E {
  return assertDefined(root.querySelector<E>(selector), `Missing required element: ${selector}`);
}
