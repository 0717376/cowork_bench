// Real compiled server -> MCP stdio -> client -> router -> dedicated PostgreSQL.
const {test} = require('node:test');
const assert = require('node:assert/strict');
const {enabled, database, mcp, data} = require('./mock_pg_harness.cjs');

test('Canvas account course search over real MCP/PostgreSQL', {skip: !enabled}, async t => {
  const {db, config} = await database('mcp-canvas-lms');
  let server;
  try {
    await db.query(`DROP SCHEMA IF EXISTS canvas CASCADE; CREATE SCHEMA canvas;
      CREATE TABLE canvas.courses(id int PRIMARY KEY, account_id int, name text, course_code text);
      INSERT INTO canvas.courses VALUES
      (1,4,'Autumn 2013','AAA-2013J'), (2,4,'Autumn 2014','AAA-2014J'),
      (3,4,'Advanced Algebra','ZZZ-2015'), (4,9,'Autumn 2014','AAA-2014J');`);
    server = await mcp('mcp-canvas-lms', config);
    const tools = await server.request('tools/list', {});
    const tool = tools.tools.find(x => x.name === 'canvas_list_account_courses');
    assert.equal(tool.inputSchema.properties.search_term.type, 'string');
    for (const [label, args, expected] of [
      ['unfiltered account', {account_id: 4}, [1,2,3]],
      ['case insensitive code', {account_id: 4, search_term: '2014j'}, [2]],
      ['case insensitive name', {account_id: 4, search_term: 'alGeBrA'}, [3]],
      ['genuine empty search', {account_id: 4, search_term: 'missing'}, []],
      ['account scope is retained', {account_id: 9, search_term: '2014'}, [4]],
      ['wrong account is not repaired', {account_id: 1, search_term: '2014'}, []],
      ['query remains parameterized', {account_id: 4, search_term: "' OR 1=1 --"}, []],
      ['empty search preserves listing', {account_id: 4, search_term: ''}, [1,2,3]],
    ]) {
      await t.test(label, async () => {
        const rows = data(await server.call(tool.name, args));
        assert.deepEqual(rows.map(r => r.id).sort((a,b) => a-b), expected);
      });
    }
  } finally {
    if (server) await server.close();
    try { await db.query('DROP SCHEMA IF EXISTS canvas CASCADE'); }
    finally { await db.end(); }
  }
});
