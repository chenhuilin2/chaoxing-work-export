// 设置读取测试：「打开窗口自动提取」的默认值变更与一次性升级。
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { DEFAULT_SETTINGS } = require('../.tmp/test/src/domain/settings.js');
const { SettingsRepository } = require('../.tmp/test/src/infrastructure/settings-repository.js');

const KEY = 'chaoxing-work-export:settings:v3';
const MARKER = 'chaoxing-work-export:auto-extract-default:v1';

/** 用内存对象顶替 localStorage */
function createStorage(initial = {}) {
  const data = new Map(Object.entries(initial));
  return {
    read: (key) => (data.has(key) ? JSON.parse(JSON.stringify(data.get(key))) : null),
    write: (key, value) => {
      data.set(key, value);
      return true;
    },
  };
}

test('「打开窗口自动提取」默认打开', () => {
  assert.equal(DEFAULT_SETTINGS.autoExtractOnLoad, true);
});

test('存量配置里沿用的旧默认值（关闭）在首次加载时升级为打开', () => {
  const storage = createStorage({ [KEY]: { ...DEFAULT_SETTINGS, autoExtractOnLoad: false } });
  const repository = new SettingsRepository(storage);

  assert.equal(repository.load().autoExtractOnLoad, true);
  assert.equal(storage.read(KEY).autoExtractOnLoad, true, '升级结果应写回存储');
  assert.equal(storage.read(MARKER), true, '应记下一次性的升级标记');
});

test('用户之后主动关掉不会被再次改回', () => {
  const storage = createStorage({ [KEY]: { ...DEFAULT_SETTINGS, autoExtractOnLoad: false } });
  const repository = new SettingsRepository(storage);

  repository.load();
  repository.save({ ...DEFAULT_SETTINGS, autoExtractOnLoad: false });
  assert.equal(repository.load().autoExtractOnLoad, false);
});
