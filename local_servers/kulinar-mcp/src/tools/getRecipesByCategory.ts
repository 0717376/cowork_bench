import { z } from 'zod';
import { Recipe } from '../types/index.js';
import { simplifyRecipe } from '../utils/recipeUtils.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerGetRecipesByCategoryTool(
  server: McpServer,
  recipes: Recipe[],
  categories: string[]
) {
  server.tool(
    'mcp_kulinar_getRecipesByCategory',
    `Возвращает рецепты выбранной категории. Доступные категории: ${categories.join(', ')}.`,
    {
      category: z
        .enum(categories as [string, ...string[]])
        .describe('Название категории русской кухни, например: салат, суп, горячее, десерт.'),
    },
    async ({ category }: { category: string }) => {
      const filtered = recipes.filter((r) => r.category === category);
      return {
        content: [{ type: 'text', text: JSON.stringify(filtered.map(simplifyRecipe), null, 2) }],
      };
    }
  );
}
