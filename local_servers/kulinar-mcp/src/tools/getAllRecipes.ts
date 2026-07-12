import { z } from 'zod';
import { Recipe } from '../types/index.js';
import { simplifyRecipeNameOnly } from '../utils/recipeUtils.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerGetAllRecipesTool(server: McpServer, recipes: Recipe[]) {
  server.tool(
    'mcp_kulinar_getAllRecipes',
    'Возвращает список всех рецептов русской кухни (имя, краткое описание, категория). Без параметров.',
    { no_param: z.string().optional().describe('параметров не требуется') },
    async () => {
      const simplified = recipes.map(simplifyRecipeNameOnly);
      return { content: [{ type: 'text', text: JSON.stringify(simplified, null, 2) }] };
    }
  );
}
