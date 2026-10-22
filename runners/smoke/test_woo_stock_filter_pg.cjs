// Real compiled Woo server, camelCase MCP args, actual SQL before pagination.
const {test} = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {enabled, servers, database, mcp, data} = require('./mock_pg_harness.cjs');

test('Woo stock filter over real MCP/PostgreSQL', {skip: !enabled}, async t => {
  const {db, config, requireServer} = await database('woocommerce-mcp');
  let server;
  try {
    await db.query(`DROP SCHEMA IF EXISTS wc CASCADE; CREATE SCHEMA wc;
      CREATE TABLE wc.products(id int PRIMARY KEY, name text, slug text, stock_status text, stock_quantity int, categories jsonb);
      CREATE TABLE wc.orders(id int PRIMARY KEY);
      CREATE TABLE wc.refunds(id int PRIMARY KEY, order_id int, amount numeric, reason text);
      INSERT INTO wc.orders VALUES (1);
      INSERT INTO wc.products VALUES
      (1,'Gadget Blue','blue','instock',4,'[{"id":7}]'),
      (2,'Gadget A','a','outofstock',0,'[{"id":7}]'),
      (3,'Gadget B','b','outofstock',0,'[{"id":7}]'),
      (4,'Gadget Backorder','backorder','onbackorder',0,'[{"id":7}]'),
      (5,'Other item','other','outofstock',0,'[{"id":8}]');`);
    server = await mcp('woocommerce-mcp', config);
    const tools = await server.request('tools/list', {});
    const tool = tools.tools.find(x => x.name === 'woo_products_list');
    assert.deepEqual(tool.inputSchema.properties.stockStatus.enum, ['instock','outofstock','onbackorder']);
    for (const [label, args, expected] of [
      ['outofstock only', {stockStatus: 'outofstock', perPage: 100}, [2,3,5]],
      ['instock only', {stockStatus: 'instock', perPage: 100}, [1]],
      ['backorder only', {stockStatus: 'onbackorder', perPage: 100}, [4]],
      ['without stock filter', {perPage: 100}, [1,2,3,4,5]],
      ['combined category/search first page', {stockStatus: 'outofstock', category: '7', search: 'Gadget', perPage: 1, page: 1}, [3]],
      ['combined category/search second page', {stockStatus: 'outofstock', category: '7', search: 'Gadget', perPage: 1, page: 2}, [2]],
      ['genuine empty filtered result', {stockStatus: 'instock', search: 'absent', perPage: 100}, []],
    ]) {
      await t.test(label, async () => {
        const rows = data(await server.call(tool.name, args));
        assert.deepEqual(rows.map(r => r.id).sort((a,b) => a-b), expected);
      });
    }
    await t.test('invalid filter is an MCP error, not an unfiltered success', async () => {
      const response = await server.call(tool.name, {stockStatus: "outofstock' OR 1=1 --"});
      assert.equal(response.isError, true, JSON.stringify(response));
      assert.match(response.content.map(c => c.text || '').join(''), /stock.?status/i);
      assert.equal((await db.query('SELECT count(*)::int AS n FROM wc.products')).rows[0].n, 5);
    });
    await t.test('a stock filter cannot add a missing column to unrelated orders', async () => {
      const {PgRestRouter} = await import(path.join(servers, 'woocommerce-mcp/dist/services/pg-rest-router.js'));
      const {Pool} = requireServer('pg');
      const router = Object.create(PgRestRouter.prototype);
      router.pool = new Pool(config);
      try {
        const result = await router.get('/orders', {params: {stock_status: 'outofstock'}});
        assert.deepEqual(result.data, [{id: 1, refunds: []}]);
      } finally { await router.pool.end(); }
    });
  } finally {
    if (server) await server.close();
    try { await db.query('DROP SCHEMA IF EXISTS wc CASCADE'); }
    finally { await db.end(); }
  }
});
