import { z } from 'zod';
import { Recipe, DishRecommendation } from '../types/index.js';
import { simplifyRecipe } from '../utils/recipeUtils.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerWhatToEatTool(server: McpServer, recipes: Recipe[]) {
  server.tool(
    'mcp_kulinar_whatToEat',
    'Не знаете что приготовить — подскажет набор блюд на застолье по числу гостей (мясное/рыбное горячее + салаты/гарниры).',
    {
      peopleCount: z.number().int().min(1).max(10).describe('Число гостей, 1-10.'),
    },
    async ({ peopleCount }: { peopleCount: number }) => {
      const sideCount = Math.floor((peopleCount + 1) / 2);
      const mainCount = Math.ceil((peopleCount + 1) / 2);

      let mains = recipes.filter((r) => r.category === 'горячее');
      let sides = recipes.filter((r) =>
        ['салат', 'закуска', 'гарнир'].includes(r.category)
      );

      const out: Recipe[] = [];
      const pickRandom = (pool: Recipe[]): Recipe | null => {
        if (pool.length === 0) return null;
        const i = Math.floor(Math.random() * pool.length);
        const r = pool[i];
        pool.splice(i, 1);
        return r;
      };

      const meatTypes = ['говядин', 'свинин', 'курин', 'куриц', 'индейк', 'рыб'];
      for (let i = meatTypes.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [meatTypes[i], meatTypes[j]] = [meatTypes[j], meatTypes[i]];
      }

      const pickedMains: Recipe[] = [];
      for (const m of meatTypes) {
        if (pickedMains.length >= mainCount) break;
        const cand = mains.filter((d) =>
          d.ingredients?.some((ing) => (ing.name || '').toLowerCase().includes(m))
        );
        const r = pickRandom(cand);
        if (r) {
          pickedMains.push(r);
          mains = mains.filter((d) => d.id !== r.id);
        }
      }
      while (pickedMains.length < mainCount) {
        const r = pickRandom(mains);
        if (!r) break;
        pickedMains.push(r);
      }

      const pickedSides: Recipe[] = [];
      while (pickedSides.length < sideCount) {
        const r = pickRandom(sides);
        if (!r) break;
        pickedSides.push(r);
      }

      out.push(...pickedMains, ...pickedSides);

      const rec: DishRecommendation = {
        peopleCount,
        meatDishCount: pickedMains.length,
        vegetableDishCount: pickedSides.length,
        dishes: out.map(simplifyRecipe),
        message: `Подборка для ${peopleCount} гостей: ${pickedMains.length} горячих и ${pickedSides.length} салатов/гарниров.`,
      };

      return { content: [{ type: 'text', text: JSON.stringify(rec, null, 2) }] };
    }
  );
}
