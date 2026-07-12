import { z } from 'zod';
import { Recipe } from '../types/index.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerGetRecipeByIdTool(server: McpServer, recipes: Recipe[]) {
  server.tool(
    'mcp_kulinar_getRecipeById',
    'Возвращает полный рецепт по id или названию блюда (с ингредиентами, шагами, временем).',
    {
      query: z.string().describe('id рецепта или название блюда (поддерживается нечёткое совпадение).'),
    },
    async ({ query }: { query: string }) => {
      let found = recipes.find((r) => r.id === query);
      if (!found) found = recipes.find((r) => r.name === query);
      if (!found) {
        const q = query.toLowerCase();
        found = recipes.find((r) => r.name.toLowerCase().includes(q));
      }
      if (!found) {
        const q = query.toLowerCase();
        const possible = recipes
          .filter((r) => r.name.toLowerCase().includes(q) || r.description.toLowerCase().includes(q))
          .slice(0, 5);
        if (possible.length === 0) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(
                  { error: 'Рецепт не найден', query, suggestion: 'Проверьте написание или попробуйте ключевое слово.' },
                  null,
                  2
                ),
              },
            ],
          };
        }
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  message: 'Точного совпадения нет, возможные варианты:',
                  query,
                  possibleMatches: possible.map((r) => ({
                    id: r.id,
                    name: r.name,
                    description: r.description,
                    category: r.category,
                  })),
                },
                null,
                2
              ),
            },
          ],
        };
      }
      return { content: [{ type: 'text', text: JSON.stringify(found, null, 2) }] };
    }
  );
}
