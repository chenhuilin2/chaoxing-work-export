import { rm } from 'node:fs/promises';

await Promise.all([
  rm(new URL('../.tmp', import.meta.url), { recursive: true, force: true }),
  rm(new URL('../dist', import.meta.url), { recursive: true, force: true }),
]);
