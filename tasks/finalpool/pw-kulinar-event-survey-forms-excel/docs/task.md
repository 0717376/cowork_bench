Помогите проанализировать меню для корпоративного мероприятия по итогам опроса поваров. Наша команда — служба питания «Аналитический институт» — готовит подборку блюд к мероприятию и хочет оценить их по нормативам нутриентов.

Шаг 1. Внешние нормативы. Откройте страницу http://localhost:30320 (используйте playwright) и извлеките таблицу нормативов по нутриентам. В таблице есть колонки Nutrient, Daily_Recommended_mg, Upper_Limit_mg и Priority. Сохраните эти данные в JSON-файл в рабочей директории, чтобы затем использовать их в анализе (например, Protein имеет Daily_Recommended_mg = 50000).

Шаг 2. Рецепты. Подберите в базе рецептов (kulinar) не менее 5 реальных блюд, подходящих для корпоративного обеда. Используйте инструменты kulinar (mcp_kulinar_getAllRecipes, mcp_kulinar_getRecipesByCategory, mcp_kulinar_getRecipeById, mcp_kulinar_recommendMeals, mcp_kulinar_whatToEat). Берите реальные названия блюд из базы — не выдумывайте их. Для каждого блюда оцените калорийность (Calories) и содержание белка (Protein_g) и определите, соответствует ли блюдо рекомендациям по питанию (Meets_Guidelines: «Да» или «Нет»).

Шаг 3. Скрипт обработки. С помощью terminal создайте и запустите Python-скрипт cook_survey_processor.py в рабочей директории. Скрипт читает собранные данные из созданных вами JSON-файлов, выполняет анализ и выводит файл cook_survey_results.json. В cook_survey_results.json должны присутствовать извлечённые со страницы нормативы по нутриентам (Nutrient, Daily_Recommended_mg, Upper_Limit_mg, Priority).

Шаг 4. Отчёт Excel. Создайте в рабочей директории Excel-файл Event_Survey_Report.xlsx с тремя листами. Передавайте абсолютный путь к файлу при работе с ним.

- Лист Data_Analysis. Колонки строго: Recipe, Category, Calories, Protein_g, Meets_Guidelines. По одной строке на каждое выбранное блюдо (не менее 5 строк). Recipe — реальное название блюда из kulinar (русское название допустимо). Строки отсортируйте по алфавиту по колонке Recipe.
- Лист Metrics. Колонки Metric и Value (не менее 4 строк). Включите строки: Total_Recipes (число блюд в Data_Analysis), Avg_Calories (среднее по колонке Calories, округлённое), Recipes_Meeting_Guidelines (число блюд со значением Meets_Guidelines = «Да»), Avg_Protein (среднее по колонке Protein_g). Значения должны быть согласованы с листом Data_Analysis.
- Лист Recommendations. Колонки Priority и Action (не менее 2 строк) — приоритетные действия по итогам анализа.

Шаг 5. Форма обратной связи. Создайте в forms форму с заголовком ровно «Cook Survey Feedback» (используйте create_form), чтобы собрать отзывы команды о выводах. Добавьте в форму хотя бы один вопрос (add_text_question или add_multiple_choice_question). При необходимости получить идентификатор формы используйте list_forms или get_form.
