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

const tscBin = path.join(root, 'node_modules', 'typescript', 'bin', 'tsc');
let result = spawnSync(process.execPath, [tscBin, '-p', 'tsconfig.build.json'], { cwd: root, stdio: 'inherit' });
if (result.error?.code === 'ENOENT') {
  result = spawnSync('tsc', ['-p', 'tsconfig.build.json'], { cwd: root, stdio: 'inherit' });
}
if (result.status !== 0) {
  process.exit(result.status ?? 1);
}

/**
 * 解析发布版本号。
 *
 * 项目规定「新版本记录一律写入 CHANGELOG.md」，因此油猴 `@version` 以 CHANGELOG 顶部
 * 条目的版本为准；`package.json` 只作兜底。此前 @version 固定取 package.json，而它长期
 * 停留在 3.0.0、CHANGELOG 却已到 3.0.x，导致每次构建出来的脚本版本号都是 v3.0.0，
 * 无法判断手上装的是哪一版。
 */
async function resolveVersion() {
  try {
    const changelog = await readFile(path.join(root, 'CHANGELOG.md'), 'utf8');
    const match = changelog.match(/^##[ \t]+(\d+\.\d+\.\d+)[ \t]*$/m);
    if (!match?.[1]) return { version: packageJson.version, source: 'package.json（CHANGELOG 无版本条目）' };
    const version = match[1];
    if (version !== packageJson.version) {
      console.warn(
        `[build] 版本号不一致：CHANGELOG=${version}，package.json=${packageJson.version}；` +
          `@version 以 CHANGELOG 为准，建议同步 package.json。`,
      );
    }
    return { version, source: 'CHANGELOG.md' };
  } catch {
    return { version: packageJson.version, source: 'package.json' };
  }
}

const [metadataTemplate, compiledAmdBundle, resolved] = await Promise.all([
  readFile(path.join(root, 'src', 'userscript.meta.txt'), 'utf8'),
  readFile(path.join(tempDir, 'chaoxing-work-export.amd.js'), 'utf8'),
  resolveVersion(),
]);
const metadata = metadataTemplate.replaceAll('__VERSION__', resolved.version).trimEnd();
const amdBundle = compiledAmdBundle.replaceAll('__APP_VERSION__', resolved.version);

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

console.log(
  `Built dist/chaoxing-work-export.user.js (${Buffer.byteLength(output)} bytes, @version ${resolved.version} ← ${resolved.source})`,
);
