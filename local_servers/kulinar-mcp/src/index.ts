#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { Command } from 'commander';
import { loadRecipes, getAllCategories } from './data/recipes.js';
import { registerGetAllRecipesTool } from './tools/getAllRecipes.js';
import { registerGetRecipeByIdTool } from './tools/getRecipeById.js';
import { registerGetRecipesByCategoryTool } from './tools/getRecipesByCategory.js';
import { registerRecommendMealsTool } from './tools/recommendMeals.js';
import { registerWhatToEatTool } from './tools/whatToEat.js';
import { Recipe } from './types/index.js';

let recipes: Recipe[] = [];
let categories: string[] = [];

const program = new Command()
  .option('--transport <stdio>', 'transport type', 'stdio')
  .parse(process.argv);

const opts = program.opts<{ transport: string }>();
if (opts.transport !== 'stdio') {
  console.error(`Поддерживается только --transport stdio (получено: ${opts.transport})`);
  process.exit(1);
}

function createServerInstance(): McpServer {
  const server = new McpServer(
    { name: 'kulinar-mcp', version: '0.1.0' },
    { capabilities: { logging: {} } }
  );
  registerGetAllRecipesTool(server, recipes);
  registerGetRecipesByCategoryTool(server, recipes, categories);
  registerRecommendMealsTool(server, recipes);
  registerWhatToEatTool(server, recipes);
  registerGetRecipeByIdTool(server, recipes);
  return server;
}

async function main() {
  recipes = await loadRecipes();
  categories = getAllCategories(recipes);
  console.error(`[kulinar-mcp] загружено ${recipes.length} рецептов, категорий: ${categories.length}`);

  const server = createServerInstance();
  const transport = new StdioServerTransport();
  try {
    await server.connect(transport);
    console.error('[kulinar-mcp] STDIO сервер запущен');
  } catch (e) {
    console.error('[kulinar-mcp] ошибка запуска:', e);
    process.exit(1);
  }
}

process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));

main().catch((e) => {
  console.error('[kulinar-mcp] критическая ошибка:', e);
  process.exit(1);
});
