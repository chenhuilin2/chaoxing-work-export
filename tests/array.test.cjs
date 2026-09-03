const test = require('node:test');
const assert = require('node:assert/strict');
const { shuffleCopy, uniqueBy } = require('../.tmp/test/src/utils/array.js');

test('shuffleCopy is immutable and deterministic with an injected random source', () => {
  const source = [1, 2, 3, 4];
  const sequence = [0.1, 0.7, 0.3];
  let index = 0;
  const shuffled = shuffleCopy(source, () => sequence[index++]);
  assert.deepEqual(source, [1, 2, 3, 4]);
  assert.deepEqual([...shuffled].sort(), source);
  assert.notDeepEqual(shuffled, source);
});

test('uniqueBy keeps the first value for each key', () => {
  const output = uniqueBy(
    [
      { id: 'a', value: 1 },
      { id: 'a', value: 2 },
      { id: 'b', value: 3 },
    ],
    (item) => item.id,
  );
  assert.deepEqual(output, [
    { id: 'a', value: 1 },
    { id: 'b', value: 3 },
  ]);
});
