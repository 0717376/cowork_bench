Мы запускаем сервис наборов для готовки (meal kit) вместе с нашим интернет-магазином InSales. Нужно подобрать к товарам магазина рецепты, чтобы собрать готовые наборы. Сначала проверьте в нашем интернет-магазине InSales все доступные товары: их названия, цены, остатки на складе и категории.

Затем подберите рецепты из базы рецептов kulinar. Возьмите рецепты из разных категорий (например, салаты, супы, горячее, гарниры, десерты) — нам нужно разнообразие для линейки наборов. Используйте реальные блюда из базы kulinar (инструменты: mcp_kulinar_getAllRecipes, mcp_kulinar_getRecipesByCategory, mcp_kulinar_getRecipeById, mcp_kulinar_recommendMeals, mcp_kulinar_whatToEat) — не выдумывайте названия рецептов.

С помощью terminal создайте и запустите в рабочей директории Python-скрипт meal_kit_designer.py, который читает store_products.json и recipes.json (оба файла создайте сами), по возможности сопоставляет товары с ингредиентами рецептов, рассчитывает ориентировочную стоимость наборов и сохраняет результат в meal_kit_plans.json.

Создайте файл Excel с именем Meal_Kit_Report.xlsx с четырьмя листами. Первый лист Product_Catalog должен содержать колонки Product_Name, Price (округлить до 2 знаков), Stock и Category, отсортированные по Product_Name. Второй лист Recipe_Collection должен содержать колонки Recipe_Name, Category, Difficulty и Ingredient_Count, отсортированные по Recipe_Name (названия рецептов и категории пишите на русском — так, как они приходят из базы kulinar; Difficulty — это число от 1 до 4). Третий лист Kit_Proposals должен содержать колонки Kit_Name, Recipe_Name, Estimated_Cost (округлить до 2 знаков), Margin_Pct (округлить до 1 знака) и Recommended_Price (округлить до 2 знаков), не менее 5 предложений наборов. Названия наборов (Kit_Name) придумайте сами. Четвёртый лист Summary должен содержать колонки Metric и Value со строками Total_Products, Total_Recipes, Proposed_Kits, Avg_Kit_Cost (округлить до 2 знаков) и Avg_Margin_Pct (округлить до 1 знака).

Рекомендованную цену рассчитывайте по формуле из pricing_guidelines.txt: Recommended_Price = Estimated_Cost / (1 - Margin_Pct/100).

Создайте документ Word с именем Meal_Kit_Proposal.docx с заголовком «Meal Kit Service Launch Proposal» и разделами «Product Analysis», «Recipe Selection», «Kit Design» и «Financial Projections» (заголовки разделов можно оформить на русском: «Анализ товаров», «Подбор рецептов», «Дизайн наборов», «Финансовые прогнозы»).

Отправьте письмо на адрес product-team@company.com с темой «Meal Kit Service Proposal Ready», кратко описав топ-3 предложенных набора с их ориентировочной стоимостью и маржой.

Важно: при вызове инструментов Word и Excel передавайте АБСОЛЮТНЫЙ путь к файлам Meal_Kit_Report.xlsx и Meal_Kit_Proposal.docx внутри рабочей директории (MCP-серверы word и excel не сохраняют текущую рабочую директорию).
