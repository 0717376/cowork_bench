# Руководство по задаче: Tech Video Marketing Opportunities

## Обзор
Найти товары интернет-магазина InSales, которые соответствуют актуальным технологическим темам из топовых видео канала Fireship на YouTube.

## Шаг 1. Получить топ-10 видео Fireship (2024+)
Отфильтровать видео, опубликованные начиная с 2024-01-01, отсортировать по view_count по убыванию, взять топ-10.

## Шаг 2. Классифицировать темы
Правила классификации Main_Topic:
- "AI": заголовок содержит DeepSeek, AI, OpenAI, GPT, Grok, Claude или vibe
- "Linux": заголовок содержит Linux
- "Windows": заголовок содержит Windows
- "JavaScript/Web": заголовок содержит JavaScript, CSS, TypeScript, React, Node или Deno
- "Python": заголовок содержит Python
- "Security": заголовок содержит security, hack, Hackers или encrypted
- "Tech/General": всё остальное

## Шаг 3. Сопоставить товары
Ключевые слова для поиска товаров (в этом порядке): laptop, usb, hub, adapter, tv, monitor, tablet, watch, headphone, camera
Совпадение: ключевое слово входит как подстрока (без учёта регистра) в название или описание товара.
Match_Keyword: первое подошедшее ключевое слово в порядке списка выше.

## Результат: Marketing_Opportunity_Report.xlsx

Лист 1: Video_Topics
Столбцы: Rank (1-10), Title, View_Count, Publish_Date (YYYY-MM-DD), Main_Topic
Сортировка: Rank по возрастанию

Лист 2: Product_Matches
Столбцы: Video_Title, Product_ID, Product_Name (полное), Product_Price, Match_Keyword
Сортировка: Rank видео по возрастанию, затем Product_ID по возрастанию

Лист 3: Summary (3 строки)
Строка 1: Total_Videos_Analyzed | 10
Строка 2: Total_Product_Matches | [количество]
Строка 3: Most_Common_Topic | [название темы]

## Письмо
- Кому: marketing@company.com
- Тема: "Tech Video Marketing Opportunities"
- Тело: описать топ-3 совпадения «видео — товар» (заголовок видео + название товара)
