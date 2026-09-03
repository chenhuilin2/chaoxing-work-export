import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
const tempDir = path.join(root, '.tmp');
const distDir = path.join(root, 'dist');

await rm(tempDir, { recursive: true, force: true });
await mkdir(tempDir, { recursive: true });
await mkdir(distDir, { recursive: true });

const localTsc = path.join(root, 'node_modules', '.bin', process.platform === 'win32' ? 'tsc.cmd' : 'tsc');
let result = spawnSync(localTsc, ['-p', 'tsconfig.build.json'], { cwd: root, stdio: 'inherit' });
if (result.error?.code === 'ENOENT') {
  result = spawnSync('tsc', ['-p', 'tsconfig.build.json'], { cwd: root, stdio: 'inherit' });
}
if (result.status !== 0) {
  process.exit(result.status ?? 1);
}

const [metadataTemplate, compiledAmdBundle] = await Promise.all([
  readFile(path.join(root, 'src', 'userscript.meta.txt'), 'utf8'),
  readFile(path.join(tempDir, 'chaoxing-work-export.amd.js'), 'utf8'),
]);
const metadata = metadataTemplate.replaceAll('__VERSION__', packageJson.version).trimEnd();
const amdBundle = compiledAmdBundle.replaceAll('__APP_VERSION__', packageJson.version);

const runtime = String.raw`(function () {
  'use strict';
  const modules = new Map();

  function normalizePath(value) {
    const output = [];
    for (const segment of value.split('/')) {
      if (!segment || segment === '.') continue;
      if (segment === '..') output.pop();
      else output.push(segment);
    }
    return output.join('/');
  }

  function resolveRequest(parent, request) {
    if (!request.startsWith('.')) return request;
    const base = parent.includes('/') ? parent.slice(0, parent.lastIndexOf('/') + 1) : '';
    return normalizePath(base + request);
  }

  function define(name, dependencies, factory) {
    modules.set(name, {
      dependencies,
      factory,
      exports: {},
      initialized: false,
      initializing: false,
    });
  }

  function load(name) {
    const record = modules.get(name);
    if (!record) throw new Error('[Chaoxing Work Export] Missing module: ' + name);
    if (record.initialized) return record.exports;
    if (record.initializing) return record.exports;

    record.initializing = true;
    const module = { exports: record.exports };
    const localRequire = (request) => load(resolveRequest(name, request));
    const args = record.dependencies.map((dependency) => {
      if (dependency === 'require') return localRequire;
      if (dependency === 'exports') return record.exports;
      if (dependency === 'module') return module;
      return load(resolveRequest(name, dependency));
    });
    const returned = record.factory.apply(undefined, args);
    record.exports = returned !== undefined ? returned : module.exports;
    record.initialized = true;
    record.initializing = false;
    return record.exports;
  }

  ${amdBundle}

  load('main');
})();`;

const output = `${metadata}\n\n${runtime}\n`;
await writeFile(path.join(distDir, 'chaoxing-work-export.user.js'), output, 'utf8');
await writeFile(path.join(distDir, 'chaoxing-work-export.meta.js'), `${metadata}\n`, 'utf8');

console.log(`Built dist/chaoxing-work-export.user.js (${Buffer.byteLength(output)} bytes)`);
