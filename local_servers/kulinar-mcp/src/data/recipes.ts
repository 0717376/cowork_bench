import { Recipe } from '../types/index.js';
import localRecipes from './all_recipes.json' with { type: 'json' };

export async function loadRecipes(): Promise<Recipe[]> {
  return localRecipes as Recipe[];
}

export function getAllCategories(recipes: Recipe[]): string[] {
  const set = new Set<string>();
  recipes.forEach((r) => r.category && set.add(r.category));
  return Array.from(set);
}
