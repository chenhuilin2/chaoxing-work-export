import { mkdir, readdir, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const temp = path.join(root, '.tmp', 'test');
await rm(temp, { recursive: true, force: true });
await mkdir(temp, { recursive: true });
await writeFile(path.join(temp, 'package.json'), '{"type":"commonjs"}\n', 'utf8');

const localTsc = path.join(root, 'node_modules', '.bin', process.platform === 'win32' ? 'tsc.cmd' : 'tsc');
let compile = spawnSync(localTsc, ['-p', 'tsconfig.test-build.json'], { cwd: root, stdio: 'inherit' });
if (compile.error?.code === 'ENOENT') {
  compile = spawnSync('tsc', ['-p', 'tsconfig.test-build.json'], { cwd: root, stdio: 'inherit' });
}
if (compile.status !== 0) process.exit(compile.status ?? 1);

const testDirectory = path.join(root, 'tests');
const tests = (await readdir(testDirectory))
  .filter((name) => name.endsWith('.test.cjs'))
  .sort()
  .map((name) => path.join('tests', name));
if (tests.length === 0) throw new Error('No tests found in tests/*.test.cjs');

const test = spawnSync(process.execPath, ['--test', ...tests], {
  cwd: root,
  stdio: 'inherit',
});
process.exit(test.status ?? 1);
