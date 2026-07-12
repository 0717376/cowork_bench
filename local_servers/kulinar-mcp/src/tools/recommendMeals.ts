import { z } from 'zod';
import { Recipe, MealPlan, SimpleRecipe, DayPlan } from '../types/index.js';
import { simplifyRecipe, processRecipeIngredients, categorizeIngredients } from '../utils/recipeUtils.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerRecommendMealsTool(server: McpServer, recipes: Recipe[]) {
  server.tool(
    'mcp_kulinar_recommendMeals',
    'Подбирает недельное меню русской кухни с учётом аллергий, нелюбимых ингредиентов и количества едоков. Возвращает план на будни/выходные и список покупок.',
    {
      allergies: z
        .array(z.string())
        .optional()
        .describe('Список аллергенов, например ["орехи", "молоко"].'),
      avoidItems: z
        .array(z.string())
        .optional()
        .describe('Нелюбимые ингредиенты, например ["чеснок", "лук"].'),
      peopleCount: z.number().int().min(1).max(10).describe('Число едоков, 1-10.'),
    },
    async ({
      allergies = [],
      avoidItems = [],
      peopleCount,
    }: {
      allergies?: string[];
      avoidItems?: string[];
      peopleCount: number;
    }) => {
      const filtered = recipes.filter((r) => {
        const bad = r.ingredients?.some((ing) => {
          const n = ing.name?.toLowerCase() || '';
          return (
            allergies.some((a) => n.includes(a.toLowerCase())) ||
            avoidItems.some((a) => n.includes(a.toLowerCase()))
          );
        });
        return !bad;
      });

      const byCat: Record<string, Recipe[]> = {};
      const target = ['салат', 'закуска', 'суп', 'горячее', 'гарнир', 'выпечка', 'десерт', 'напиток'];
      filtered.forEach((r) => {
        if (target.includes(r.category)) {
          (byCat[r.category] ||= []).push(r);
        }
      });

      const plan: MealPlan = {
        weekdays: [],
        weekend: [],
        groceryList: {
          ingredients: [],
          shoppingPlan: { fresh: [], pantry: [], spices: [], others: [] },
        },
      };
      const selected: Recipe[] = [];

      const pick = (cat: string): Recipe | null => {
        const pool = byCat[cat];
        if (!pool || pool.length === 0) return null;
        const idx = Math.floor(Math.random() * pool.length);
        const r = pool[idx];
        pool.splice(idx, 1);
        selected.push(r);
        return r;
      };

      const days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'];
      for (let i = 0; i < 7; i++) {
        const day: DayPlan = { day: days[i], breakfast: [], lunch: [], dinner: [] };
        const br = pick('выпечка') || pick('гарнир');
        if (br) day.breakfast.push(simplifyRecipe(br));
        const lunchCount = Math.max(2, Math.ceil(peopleCount / 4));
        for (let j = 0; j < lunchCount; j++) {
          const cats = ['суп', 'горячее', 'салат'];
          const r = pick(cats[j % cats.length]);
          if (r) day.lunch.push(simplifyRecipe(r));
        }
        const dinnerCount = i >= 5 ? 3 : 2;
        for (let j = 0; j < dinnerCount; j++) {
          const cats = ['горячее', 'гарнир', 'десерт'];
          const r = pick(cats[j % cats.length]);
          if (r) day.dinner.push(simplifyRecipe(r));
        }
        (i < 5 ? plan.weekdays : plan.weekend).push(day);
      }

      const ingMap = new Map<
        string,
        { totalQuantity: number | null; unit: string | null; recipeCount: number; recipes: string[] }
      >();
      selected.forEach((r) => processRecipeIngredients(r, ingMap));
      for (const [name, info] of ingMap.entries()) {
        plan.groceryList.ingredients.push({
          name,
          totalQuantity: info.totalQuantity,
          unit: info.unit,
          recipeCount: info.recipeCount,
          recipes: info.recipes,
        });
      }
      plan.groceryList.ingredients.sort((a, b) => b.recipeCount - a.recipeCount);
      categorizeIngredients(plan.groceryList.ingredients, plan.groceryList.shoppingPlan);

      return { content: [{ type: 'text', text: JSON.stringify(plan, null, 2) }] };
    }
  );
}
