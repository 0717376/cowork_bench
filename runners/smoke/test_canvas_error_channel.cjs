// Node >=22.13. Real TS router; pool failure is deliberately injected offline.
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
const calls = [
  ['get', '/courses/1', undefined],
  ['post', '/accounts/4/courses', {name: 'synthetic'}],
  ['put', '/courses/1', {name: 'synthetic'}],
  ['delete', '/courses/1', undefined],
];

for (const [method, url, data] of calls) {
  test(method+' rejects the original DB failure once, including request() dispatch', async () => {
    for (const dispatched of [false, true]) {
      const instance = Object.create(Router.prototype);
      const original = Object.assign(new Error('SYNTHETIC_DB_FAILURE'), {code: '42501'});
      let count = 0;
      instance.pool = {query: async () => { count++; throw original; }};
      const diagnostics = [];
      const previous = console.error;
      console.error = (...args) => diagnostics.push(args.join(' '));
      try {
        await assert.rejects(dispatched ? instance.request({method, url, data}) : instance[method](url, data),
          error => error === original && error.code === '42501');
      } finally { console.error = previous; }
      assert.equal(count, 1, 'mutation/read must not retry implicitly');
      assert.equal(diagnostics.length, 1);
      assert.ok(diagnostics[0].includes(method.toUpperCase()));
      assert.ok(diagnostics[0].includes('SYNTHETIC_DB_FAILURE'));
    }
  });
}

test('normal empty read and successful CRUD preserve their response contract', async () => {
  const instance = Object.create(Router.prototype);
  const queries = [];
  instance.pool = {query: async (sql) => {
    queries.push(sql);
    return {rows: sql.startsWith('SELECT') || sql.startsWith('DELETE') ? [] : [{id: 1, name: 'synthetic'}]};
  }};
  const empty = await instance.get('/accounts/4/courses');
  assert.equal(empty.status, 200);
  assert.deepEqual(empty.data, []);
  const created = await instance.post('/accounts/4/courses', {name: 'synthetic'});
  assert.equal(created.status, 201);
  assert.equal(created.data.id, 1);
  const updated = await instance.put('/courses/1', {name: 'synthetic'});
  assert.equal(updated.status, 200);
  assert.equal(updated.data.id, 1);
  const removed = await instance.delete('/courses/1');
  assert.equal(removed.status, 200);
  assert.deepEqual(removed.data, {});
  assert.equal(queries.length, 4);
});
