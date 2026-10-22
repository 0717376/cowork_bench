// Shared optional integration support. Only an empty, dedicated DB is allowed.
const assert = require('node:assert/strict');
const {createRequire} = require('node:module');
const {spawn} = require('node:child_process');
const {createInterface} = require('node:readline');
const {mkdtempSync, rmSync} = require('node:fs');
const {tmpdir} = require('node:os');
const path = require('node:path');
const ROOT = path.resolve(__dirname, '../..');
const enabled = Boolean(process.env.BENCH_MOCK_TEST_PG_HOST);
const servers = process.env.BENCH_MOCK_TEST_SERVER_ROOT || path.join(ROOT, 'local_servers');

async function database(server) {
  const name = process.env.BENCH_MOCK_TEST_PG_DATABASE || 'cowork_mock_contract_test';
  assert.equal(name, 'cowork_mock_contract_test', 'Refusing any retained benchmark database');
  assert.ok(enabled, 'Set BENCH_MOCK_TEST_PG_HOST for this optional integration');
  const requireServer = createRequire(path.join(servers, server, 'package.json'));
  const {Client} = requireServer('pg');
  const config = {
    host: process.env.BENCH_MOCK_TEST_PG_HOST,
    port: Number(process.env.BENCH_MOCK_TEST_PG_PORT || 5432),
    database: name,
    user: process.env.BENCH_MOCK_TEST_PG_USER || 'postgres',
    password: process.env.BENCH_MOCK_TEST_PG_PASSWORD || '',
  };
  const db = new Client(config);
  await db.connect();
  // Serialize suites that share this dedicated DB, even across Node processes.
  await db.query("SELECT pg_advisory_lock(hashtext('cowork_mock_contract_test'))");
  return {db, config, requireServer};
}

async function mcp(server, config) {
  const folder = mkdtempSync(path.join(tmpdir(), 'cowork-mock-contract-'));
  const entry = path.join(servers, server, server === 'woocommerce-mcp' ? 'dist/index.js' : 'build/index.js');
  const child = spawn(process.execPath, [entry], {
    cwd: folder,
    env: {
      PATH: process.env.PATH,
      HOME: folder,
      PG_HOST: config.host, PG_PORT: String(config.port), PG_DATABASE: config.database,
      PG_USER: config.user, PG_PASSWORD: config.password,
      CANVAS_USE_PG: '1', CANVAS_API_TOKEN: 'synthetic-test-token', CANVAS_DOMAIN: 'example.invalid',
      WORDPRESS_USE_PG: '1',
    },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  let stderr = '';
  let nextId = 0;
  let exited = false;
  let closed = false;
  const pending = new Map();
  child.stderr.on('data', data => { stderr = (stderr+data).slice(-24000); });
  const rejectAll = error => {
    for (const p of pending.values()) { clearTimeout(p.timer); p.reject(error); }
    pending.clear();
  };
  child.on('error', rejectAll);
  const exit = new Promise(resolve => child.on('exit', (code, signal) => {
    exited = true;
    rejectAll(new Error('MCP exited '+code+'/'+signal+'\n'+stderr));
    resolve();
  }));
  const lines = createInterface({input: child.stdout});
  lines.on('line', line => {
    let message;
    try { message = JSON.parse(line); }
    catch { rejectAll(new Error('Non-JSON MCP stdout: '+line)); return; }
    if (!pending.has(message.id)) return;
    const p = pending.get(message.id);
    pending.delete(message.id);
    clearTimeout(p.timer);
    if (message.error) p.reject(new Error(JSON.stringify(message.error)));
    else p.resolve(message.result);
  });
  async function request(method, params) {
    assert.ok(!exited, 'MCP process already exited\n'+stderr);
    const id = ++nextId;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { pending.delete(id); reject(new Error('MCP timeout: '+method+'\n'+stderr)); }, 15000);
      pending.set(id, {resolve, reject, timer});
      child.stdin.write(JSON.stringify({jsonrpc: '2.0', id, method, params})+'\n', error => {
        if (error) { clearTimeout(timer); pending.delete(id); reject(error); }
      });
    });
  }
  async function close() {
    if (closed) return;
    closed = true;
    lines.close();
    child.stdin.end();
    if (!exited) child.kill('SIGTERM');
    const killTimer = setTimeout(() => { if (!exited) child.kill('SIGKILL'); }, 2000);
    await exit;
    clearTimeout(killTimer);
    rmSync(folder, {recursive: true, force: true});
  }
  try {
    await request('initialize', {protocolVersion: '2024-11-05', capabilities: {}, clientInfo: {name: 'synthetic-contract-test', version: '1.0'}});
    child.stdin.write(JSON.stringify({jsonrpc: '2.0', method: 'notifications/initialized'})+'\n');
  } catch (error) { await close(); throw error; }
  return {request, close, stderr: () => stderr,
    call: (name, args) => request('tools/call', {name, arguments: args})};
}

function data(result) {
  assert.notEqual(result.isError, true, JSON.stringify(result));
  return JSON.parse(result.content.filter(c => c.type === 'text').map(c => c.text).join(''));
}

module.exports = {enabled, servers, database, mcp, data};
