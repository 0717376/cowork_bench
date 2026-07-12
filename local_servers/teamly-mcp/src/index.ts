#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerTools } from './tools.js';
import { pool, query } from './db.js';

function createServer(): McpServer {
  const server = new McpServer(
    { name: 'teamly-mcp', version: '0.1.0' },
    { capabilities: { logging: {} } },
  );
  registerTools(server);
  return server;
}

async function main() {
  try {
    const r = await query<{ n: string }>('SELECT COUNT(*)::text AS n FROM teamly.spaces');
    console.error(`[teamly-mcp] подключение к PG ok, пространств: ${r.rows[0].n}`);
  } catch (e) {
    console.error('[teamly-mcp] ошибка подключения к PG:', e);
    process.exit(1);
  }

  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('[teamly-mcp] STDIO сервер запущен');
}

process.on('SIGINT', async () => { await pool.end().catch(() => {}); process.exit(0); });
process.on('SIGTERM', async () => { await pool.end().catch(() => {}); process.exit(0); });

main().catch((e) => {
  console.error('[teamly-mcp] критическая ошибка:', e);
  process.exit(1);
});
