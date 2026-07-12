import { Recipe, SimpleRecipe, NameOnlyRecipe, Ingredient } from '../types/index.js';

export function simplifyRecipe(recipe: Recipe): SimpleRecipe {
  return {
    id: recipe.id,
    name: recipe.name,
    description: recipe.description,
    ingredients: recipe.ingredients.map((i: Ingredient) => ({
      name: i.name,
      text_quantity: i.text_quantity,
    })),
  };
}

export function simplifyRecipeNameOnly(recipe: Recipe): NameOnlyRecipe {
  return {
    name: recipe.name,
    description: recipe.description,
    category: recipe.category,
  };
}

export function processRecipeIngredients(
  recipe: Recipe,
  ingredientMap: Map<
    string,
    { totalQuantity: number | null; unit: string | null; recipeCount: number; recipes: string[] }
  >
) {
  recipe.ingredients?.forEach((ing: Ingredient) => {
    const key = ing.name.toLowerCase();
    if (!ingredientMap.has(key)) {
      ingredientMap.set(key, {
        totalQuantity: ing.quantity,
        unit: ing.unit,
        recipeCount: 1,
        recipes: [recipe.name],
      });
    } else {
      const ex = ingredientMap.get(key)!;
      if (ex.unit && ing.unit && ex.unit === ing.unit && ex.totalQuantity !== null && ing.quantity !== null) {
        ex.totalQuantity += ing.quantity;
      } else {
        ex.totalQuantity = null;
        ex.unit = null;
      }
      ex.recipeCount += 1;
      if (!ex.recipes.includes(recipe.name)) ex.recipes.push(recipe.name);
    }
  });
}

export function categorizeIngredients(
  ingredients: Array<{ name: string; totalQuantity: number | null; unit: string | null; recipeCount: number; recipes: string[] }>,
  shoppingPlan: { fresh: string[]; pantry: string[]; spices: string[]; others: string[] }
) {
  const spice = ['соль', 'сахар', 'перец', 'лавровый', 'укроп', 'петрушка', 'тимьян', 'кориандр', 'паприка', 'уксус', 'горчиц', 'хрен', 'чеснок', 'лук', 'специ', 'приправ'];
  const fresh = ['мясо', 'говядин', 'свинин', 'курин', 'куриц', 'рыб', 'судак', 'сёмг', 'семг', 'окунь', 'яйц', 'молок', 'сметан', 'творог', 'сыр', 'масл', 'капуст', 'морков', 'картофел', 'свёкл', 'свекл', 'огурц', 'помидор', 'томат', 'зелен', 'грибы', 'шампиньон', 'опят', 'сельдь', 'икр'];
  const pantry = ['мука', 'крупа', 'греч', 'рис', 'перлов', 'манк', 'макарон', 'лапш', 'хлеб', 'сухарь', 'дрожж', 'сода', 'разрыхл', 'крахмал', 'сахар', 'мёд', 'мед', 'желатин', 'консерв', 'тушёнк', 'тушенк'];

  ingredients.forEach((ing) => {
    const n = ing.name.toLowerCase();
    if (spice.some((k) => n.includes(k))) shoppingPlan.spices.push(ing.name);
    else if (fresh.some((k) => n.includes(k))) shoppingPlan.fresh.push(ing.name);
    else if (pantry.some((k) => n.includes(k))) shoppingPlan.pantry.push(ing.name);
    else shoppingPlan.others.push(ing.name);
  });
}
