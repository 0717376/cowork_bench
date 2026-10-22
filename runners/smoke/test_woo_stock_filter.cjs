'use strict';

// Offline tests for the stock_status filter in PgRestRouter
// (local_servers/woocommerce-mcp/src/services/pg-rest-router.ts).
//
// Seams — this file deliberately does NOT use PostgreSQL:
//  - The real router TS file is loaded by replacing its single pg import with
//    a forbidden Pool stub and dropping the `export` keyword, then stripping
//    types via node:module stripTypeScriptTypes ('transform' mode only if
//    'strip' cannot parse the file).
//  - The router instance is created with Object.create(PgRestRouter.prototype)
//    and a synthetic pool, so the real constructor (needs pg.Pool) never runs.
//  - The pool stub records (text, values) and returns canned rows: assertions
//    verify the SQL text and bound values the router *would* send, not real
//    SQL semantics. Real SQL and MCP are checked by test_woo_stock_filter_pg.cjs.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { stripTypeScriptTypes } = require('node:module');

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const ROUTER_TS = path.join(REPO_ROOT, 'local_servers', 'woocommerce-mcp', 'src', 'services', 'pg-rest-router.ts');

const PG_IMPORT = "import pg from 'pg';\nconst { Pool } = pg;";

let loaded = null;

function loadRouter() {
  if (loaded) return loaded;
  const src = fs.readFileSync(ROUTER_TS, 'utf8');
  assert.ok(src.includes(PG_IMPORT), 'pg import block not found verbatim — update this loader');
  assert.ok(src.includes('export class PgRestRouter'), 'export class PgRestRouter not found');
  // Pool becomes an always-invalid stub: constructing the router offline is a bug.
  let js = src
    .replace(PG_IMPORT, 'const Pool = undefined;')
    .replace('export class PgRestRouter', 'class PgRestRouter');
  let mode = 'strip';
  try {
    js = stripTypeScriptTypes(js);
  } catch {
    mode = 'transform';
    js = stripTypeScriptTypes(js, { mode: 'transform' });
  }
  const PgRestRouter = new Function(`${js};\nreturn PgRestRouter;`)();
  loaded = { PgRestRouter, mode };
  console.log(`[woo_stock_filter] loaded pg-rest-router.ts, stripTypeScriptTypes mode=${mode}`);
  return loaded;
}

function makePool({ rows = [], error = null } = {}) {
  const calls = [];
  return {
    calls,
    query: async (text, values) => {
      calls.push({ text, values });
      if (error) throw error;
      return { rows };
    },
  };
}

function makeRouter(pool) {
  const { PgRestRouter } = loadRouter();
  const router = Object.create(PgRestRouter.prototype);
  router.pool = pool;
  return router;
}

test('each stockStatus enum value becomes a bound stock_status equality', async () => {
  for (const value of ['instock', 'outofstock', 'onbackorder']) {
    const pool = makePool({ rows: [{ id: 1, stock_status: value }] });
    const res = await makeRouter(pool).get('products', { params: { stock_status: value } });
    assert.equal(res.status, 200);
    assert.deepEqual(res.data, [{ id: 1, stock_status: value }]);
    assert.equal(pool.calls.length, 1);
    assert.equal(
      pool.calls[0].text,
      'SELECT * FROM wc.products WHERE stock_status = $1 ORDER BY id DESC LIMIT 10 OFFSET 0'
    );
    assert.deepEqual(pool.calls[0].values, [value]);
  }
});

test('unset, null and empty stock_status add no WHERE clause', async () => {
  const cases = [{}, { stock_status: undefined }, { stock_status: null }, { stock_status: '' }];
  for (const params of cases) {
    const pool = makePool({ rows: [] });
    const res = await makeRouter(pool).get('products', { params });
    assert.equal(res.status, 200);
    assert.deepEqual(res.data, []);
    assert.equal(pool.calls.length, 1);
    assert.equal(pool.calls[0].text, 'SELECT * FROM wc.products ORDER BY id DESC LIMIT 10 OFFSET 0');
    assert.deepEqual(pool.calls[0].values, []);
  }
});

test('invalid stock_status values throw descriptively before any SQL', async () => {
  const bad = [
    'In Stock', 'INSTOCK', 'instock ', 'any', "instock' OR 1=1 --",
    'instock; DROP TABLE wc.products; --', 42, true,
    ['instock'], ['instock', 'outofstock'], { stock_status: 'instock' },
  ];
  for (const value of bad) {
    const pool = makePool({ rows: [] });
    await assert.rejects(
      makeRouter(pool).get('products', { params: { stock_status: value } }),
      (err) => err instanceof Error &&
        /stock_status/.test(err.message) &&
        /instock, outofstock, onbackorder/.test(err.message)
    );
    assert.equal(pool.calls.length, 0, `no SQL must run for ${JSON.stringify(value)}`);
  }
  // Invalid stock_status must also abort a combined request before any query.
  const pool = makePool({ rows: [] });
  await assert.rejects(
    makeRouter(pool).get('products', { params: { stock_status: 'bogus', category: '3', search: 'x' } })
  );
  assert.equal(pool.calls.length, 0);
});

test('stock_status on unrelated resources is ignored, no nonexistent column', async () => {
  // wc.orders has no stock_status column: the param must be dropped entirely.
  const poolOrders = makePool({ rows: [] });
  const resOrders = await makeRouter(poolOrders).get('orders', { params: { stock_status: 'instock' } });
  assert.equal(resOrders.status, 200);
  assert.equal(poolOrders.calls.length, 1);
  assert.ok(poolOrders.calls[0].text.startsWith('SELECT * FROM wc.orders'));
  assert.ok(!poolOrders.calls[0].text.includes('stock_status'));

  // Filter is scoped to wc.products, so even a bogus value must not throw for orders.
  const poolBogus = makePool({ rows: [] });
  await makeRouter(poolBogus).get('orders', { params: { stock_status: 'nonsense' } });
  assert.ok(!poolBogus.calls[0].text.includes('stock_status'));

  // Nested variations map to wc.product_variations, not wc.products.
  const poolVars = makePool({ rows: [] });
  const resVars = await makeRouter(poolVars).get('products/7/variations', { params: { stock_status: 'instock' } });
  assert.equal(resVars.status, 200);
  assert.ok(poolVars.calls[0].text.includes('FROM wc.product_variations'));
  assert.ok(poolVars.calls[0].text.includes('product_id = $1'));
  assert.ok(!poolVars.calls[0].text.includes('stock_status'));
});

test('stock_status combines with category, search and pagination; WHERE before LIMIT', async () => {
  const pool = makePool({ rows: [{ id: 9 }] });
  const res = await makeRouter(pool).get('products', {
    params: { stock_status: 'outofstock', category: '12', search: 'футболка', page: 2, per_page: 5 },
  });
  assert.equal(res.status, 200);
  assert.deepEqual(res.data, [{ id: 9 }]);
  assert.equal(pool.calls.length, 1);
  const { text, values } = pool.calls[0];
  assert.equal(
    text,
    'SELECT * FROM wc.products WHERE (name ILIKE $1 OR slug ILIKE $1) AND categories @> $2::jsonb ' +
    'AND stock_status = $3 ORDER BY id DESC LIMIT 5 OFFSET 5'
  );
  assert.deepEqual(values, ['%футболка%', '[{"id":12}]', 'outofstock']);
  const iWhere = text.indexOf(' WHERE ');
  const iLimit = text.indexOf(' LIMIT ');
  assert.ok(iWhere !== -1 && iLimit > iWhere, 'WHERE must precede LIMIT');
  const iStock = text.indexOf('stock_status = $3');
  assert.ok(iStock > iWhere && iStock < iLimit, 'stock_status equality must live inside WHERE');
});

test('legitimate empty result is a 200 with empty data', async () => {
  const pool = makePool({ rows: [] });
  const res = await makeRouter(pool).get('products', { params: { stock_status: 'onbackorder' } });
  assert.equal(res.status, 200);
  assert.deepEqual(res.data, []);
  assert.deepEqual(pool.calls[0].values, ['onbackorder']);
});

test('pool failure rejects instead of fake success', async () => {
  const pool = makePool({ rows: [], error: new Error('connection refused') });
  await assert.rejects(
    makeRouter(pool).get('products', { params: { stock_status: 'instock' } }),
    /connection refused/
  );
  assert.equal(pool.calls.length, 1);
});
