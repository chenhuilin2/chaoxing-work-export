import { access, mkdir, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const temp = path.join(root, '.tmp', 'browser-smoke');
const bundle = path.join(root, 'dist', 'chaoxing-work-export.user.js');
await access(bundle);
await mkdir(temp, { recursive: true });

const fixture = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>浏览器冒烟测试</title>
  <script>
    window.__copied = '';
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: async (text) => { window.__copied = text; } }
    });
  </script>
</head>
<body>
  <section class="mark_item">
    <h2 class="type_tit">单选题</h2>
    <article class="questionLi" typeName="单选题">
      <h3 class="mark_name">1.【单选题】HTTP 默认使用哪个端口？</h3>
      <ul class="mark_letter">
        <li>A. 80</li>
        <li>B. 443</li>
      </ul>
      <div class="mark_answer">
        <div class="mark_key">
          <div class="colorDeep"><span class="stuAnswerContent">B</span></div>
          <div class="colorGreen"><span class="stuAnswerContent">A</span></div>
        </div>
        <div class="analysis">HTTP 的默认端口是 80。</div>
      </div>
    </article>
  </section>
  <script src="../../dist/chaoxing-work-export.user.js"></script>
  <script>
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const waitFor = async (check, timeout = 2500) => {
      const deadline = Date.now() + timeout;
      while (Date.now() < deadline) {
        if (check()) return true;
        await sleep(30);
      }
      return false;
    };
    (async () => {
      try {
        const hostReady = await waitFor(() => document.getElementById('chaoxing-work-export-root'));
        if (!hostReady) throw new Error('UI host was not mounted');
        const shadow = document.getElementById('chaoxing-work-export-root').shadowRoot;
        shadow.querySelector('.cwe-launcher').click();
        shadow.querySelector('[data-action="extract"]').click();
        const extracted = await waitFor(() => shadow.querySelector('[data-stat="total"]').textContent === '1');
        if (!extracted) throw new Error('Question was not extracted');
        if (shadow.querySelector('[data-stat="wrong"]').textContent !== '1') throw new Error('Wrong-answer detection failed');
        if (shadow.querySelector('[data-stat="analysis"]').textContent !== '1') throw new Error('Analysis extraction failed');

        const answers = shadow.querySelector('[data-option="withAnswers"]');
        answers.checked = true;
        answers.dispatchEvent(new Event('change', { bubbles: true }));
        const analysis = shadow.querySelector('[data-option="includeAnalysis"]');
        analysis.checked = true;
        analysis.dispatchEvent(new Event('change', { bubbles: true }));
        shadow.querySelector('[data-action="copy"]').click();
        const copied = await waitFor(() => window.__copied.includes('HTTP 默认使用哪个端口'));
        if (!copied || !window.__copied.includes('HTTP 的默认端口是 80')) throw new Error('Copy formatter failed');
        document.body.dataset.smoke = 'CWE_SMOKE_PASS';
      } catch (error) {
        document.body.dataset.smoke = 'CWE_SMOKE_FAIL';
        document.body.dataset.error = String(error && error.message ? error.message : error);
      }
    })();
  </script>
</body>
</html>`;

const fixturePath = path.join(temp, 'fixture.html');
await writeFile(fixturePath, fixture, 'utf8');

const candidates = [
  process.env.CHROME_BIN,
  'chromium',
  'chromium-browser',
  'google-chrome',
  'google-chrome-stable',
].filter(Boolean);
let executable = null;
for (const candidate of candidates) {
  const probe = spawnSync(candidate, ['--version'], { encoding: 'utf8' });
  if (probe.status === 0) {
    executable = candidate;
    break;
  }
}
if (!executable) throw new Error('Chromium/Chrome is required for the browser smoke test');

const run = spawnSync(
  executable,
  [
    '--headless=new',
    '--no-sandbox',
    '--disable-gpu',
    '--disable-background-networking',
    '--disable-default-apps',
    '--disable-extensions',
    '--virtual-time-budget=3500',
    '--dump-dom',
    `file://${fixturePath}`,
  ],
  { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024, timeout: 20_000, killSignal: 'SIGKILL' },
);
if (run.error) throw run.error;
if (run.status !== 0) {
  process.stderr.write(run.stderr || run.stdout);
  process.exit(run.status ?? 1);
}
if (!run.stdout.includes('data-smoke="CWE_SMOKE_PASS"')) {
  const error = run.stdout.match(/data-error="([^"]*)"/)?.[1] ?? 'unknown browser failure';
  throw new Error(`Browser smoke test failed: ${error}`);
}
console.log(`Browser smoke test passed with ${executable}`);
