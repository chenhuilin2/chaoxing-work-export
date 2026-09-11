export class SafeStorage {
  constructor(private readonly storage: Storage = window.localStorage) {}

  read(key: string): unknown {
    try {
      const value = this.storage.getItem(key);
      return value === null ? null : (JSON.parse(value) as unknown);
    } catch {
      return null;
    }
  }

  write(key: string, value: unknown): boolean {
    try {
      this.storage.setItem(key, JSON.stringify(value));
      return true;
    } catch {
      return false;
    }
  }

  remove(key: string): void {
    try {
      this.storage.removeItem(key);
    } catch {
      // Storage can be unavailable in privacy modes. Failing silently keeps extraction usable.
    }
  }
}
