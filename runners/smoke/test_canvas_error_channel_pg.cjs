// Controlled real DB failure, not an inference from an ordinary empty reply.
const {test} = require('node:test');
const assert = require('node:assert/strict');
const {enabled, database, mcp, data} = require('./mock_pg_harness.cjs');

test('Canvas DB failure reaches the MCP error channel without retry or mutation', {skip: !enabled}, async t => {
  const {db, config} = await database('mcp-canvas-lms');
  let server;
  try {
    await db.query(`DROP SCHEMA IF EXISTS canvas CASCADE; CREATE SCHEMA canvas;
      CREATE TABLE canvas.courses(id serial PRIMARY KEY, account_id int, name text NOT NULL, course_code text);
      INSERT INTO canvas.courses(account_id,name,course_code) VALUES(4,'Original course','SYN-1');`);
    server = await mcp('mcp-canvas-lms', config);
    await t.test('a real empty result is a successful empty list', async () => {
      assert.deepEqual(data(await server.call('canvas_list_account_courses', {account_id: 99})), []);
    });
    await t.test('normal create/read/update/delete actually persists its effects', async () => {
      const created = data(await server.call('canvas_create_course', {account_id: 4, name: 'Created course', course_code: 'SYN-2'}));
      assert.equal(created.name, 'Created course');
      assert.equal((await db.query('SELECT name FROM canvas.courses WHERE id=$1', [created.id])).rows[0].name, created.name);
      assert.equal(data(await server.call('canvas_get_course', {course_id: created.id})).name, created.name);
      const updated = data(await server.call('canvas_update_course', {course_id: created.id, name: 'Updated course'}));
      assert.equal(updated.name, 'Updated course');
      assert.equal((await db.query('SELECT name FROM canvas.courses WHERE id=$1', [created.id])).rows[0].name, 'Updated course');
      const deleted = await server.call('canvas_delete_course', {course_id: created.id});
      assert.notEqual(deleted.isError, true);
      assert.equal((await db.query('SELECT * FROM canvas.courses WHERE id=$1', [created.id])).rows.length, 0);
    });
    const before = (await db.query('SELECT * FROM canvas.courses ORDER BY id')).rows;
    // Rename the table only in this disposable schema: every CRUD query now
    // receives PostgreSQL 42P01. Retain the original rows to prove no mutation.
    await db.query('ALTER TABLE canvas.courses RENAME TO hidden_courses');
    for (const [method, name, args] of [
      ['GET', 'canvas_get_course', {course_id: 1}],
      ['POST', 'canvas_create_course', {account_id: 4, name: 'Must not exist'}],
      ['PUT', 'canvas_update_course', {course_id: 1, name: 'Must not persist'}],
      ['DELETE', 'canvas_delete_course', {course_id: 1}],
    ]) {
      await t.test(method+' returns MCP isError and preserves the DB cause', async () => {
        const result = await server.call(name, args);
        assert.equal(result.isError, true, JSON.stringify(result));
        assert.match(result.content.map(c => c.text || '').join(''), /relation .*canvas\.courses.* does not exist/i);
        const diagnostic = new RegExp('\\[PgCanvasRouter\\] '+method+' .* error:', 'g');
        assert.equal((server.stderr().match(diagnostic) || []).length, 1, 'no implicit retry in PG mode');
      });
    }
    await t.test('failed calls leave retained rows unchanged and recovery works', async () => {
      assert.deepEqual((await db.query('SELECT * FROM canvas.hidden_courses ORDER BY id')).rows, before);
      await db.query('ALTER TABLE canvas.hidden_courses RENAME TO courses');
      assert.equal(data(await server.call('canvas_get_course', {course_id: 1})).name, 'Original course');
    });
  } finally {
    if (server) await server.close();
    try { await db.query('DROP SCHEMA IF EXISTS canvas CASCADE'); }
    finally { await db.end(); }
  }
});
