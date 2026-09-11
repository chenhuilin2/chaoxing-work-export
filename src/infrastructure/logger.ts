export class Logger {
  constructor(private readonly prefix = '[Chaoxing Work Export]') {}

  debug(message: string, details?: unknown): void {
    try {
      if (window.localStorage.getItem('chaoxing-work-export:debug') !== '1') return;
    } catch {
      return;
    }
    console.debug(this.prefix, message, details ?? '');
  }

  warn(message: string, details?: unknown): void {
    console.warn(this.prefix, message, details ?? '');
  }

  error(message: string, details?: unknown): void {
    console.error(this.prefix, message, details ?? '');
  }
}
