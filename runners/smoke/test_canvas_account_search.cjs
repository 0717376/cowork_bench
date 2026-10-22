// Node >=22.13. Real TS router, isolated pool seam; real SQL/MCP checked separately.
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {stripTypeScriptTypes} = require('node:module');
const source = path.resolve(__dirname, '../../local_servers/mcp-canvas-lms/src/pg-canvas-router.ts');
const code = stripTypeScriptTypes(fs.readFileSync(source, 'utf8')
  .replace(/^import pg from 'pg';\r?\n/m, '').replace('export class ', 'class '));
const Router = new Function('pg', code+'\nreturn PgCanvasRouter;')({
  Pool: class { constructor() { throw Error('Offline tests must not connect'); } }
});

function router(rows) {
  const instance = Object.create(Router.prototype);
  const queries = [];
  instance.pool = {query: async (sql, values) => { queries.push({sql, values}); return {rows}; }};
  return {instance, queries};
}

test('account restriction and ordinary listing stay intact without search', async () => {
  for (const params of [{}, {search_term: ''}]) {
    const {instance, queries} = router([{id: 1, account_id: 4}]);
    const result = await instance.get('/accounts/4/courses', {params});
    assert.equal(result.status, 200);
    assert.deepEqual(result.data, [{id: 1, account_id: 4}]);
    assert.equal(queries.length, 1);
    assert.equal(queries[0].sql, 'SELECT * FROM canvas.courses WHERE account_id = $1');
    assert.deepEqual(queries[0].values, ['4']);
  }
});

test('search adds bound case-insensitive name OR code under the same account', async () => {
  const {instance, queries} = router([{id: 2, course_code: 'SYN-2014J'}]);
  const result = await instance.get('/accounts/4/courses', {params: {search_term: '2014j'}});
  assert.equal(result.status, 200);
  assert.equal(result.data[0].id, 2);
  assert.match(queries[0].sql, /WHERE account_id = \$1 AND \(name ILIKE \$2 OR course_code ILIKE \$2\)/);
  assert.deepEqual(queries[0].values, ['4', '%2014j%']);
});

test('search text cannot become SQL or replace account selection', async () => {
  const search = "' OR account_id=9 --";
  const {instance, queries} = router([]);
  const result = await instance.get('/accounts/4/courses', {params: {search_term: search}});
  assert.equal(result.status, 200);
  assert.deepEqual(result.data, []);
  assert.ok(!queries[0].sql.includes(search));
  assert.deepEqual(queries[0].values, ['4', '%'+search+'%']);
});

test('genuine empty filtered result remains successful and empty', async () => {
  const {instance, queries} = router([]);
  const result = await instance.get('/accounts/99/courses', {params: {search_term: 'missing'}});
  assert.equal(result.status, 200);
  assert.deepEqual(result.data, []);
  assert.deepEqual(queries[0].values, ['99', '%missing%']);
});
