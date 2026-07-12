Мне нужно оптимизировать цены на наши товары на основе данных о конкурентах. По адресу http://localhost:30305 находится дашборд с ценами конкурентов на товары из наших категорий. Зайдите на эту страницу и извлеките все данные о ценах.

Затем выгрузите наш текущий каталог товаров из магазина InSales, чтобы получить названия товаров, цены, остатки на складе и показатели продаж.

С помощью терминала создайте и запустите Python-скрипт price_optimizer.py в рабочей директории. Скрипт должен прочитать competitor_prices.json и our_products.json (оба файла вы создаёте сами), сопоставить товары по названию, рассчитать разницу в ценах и записать price_recommendations.json с рекомендациями по оптимизации.

Создайте Excel-файл Price_Optimization_Report.xlsx с тремя листами. Первый лист Price_Comparison должен содержать столбцы Product_Name, Our_Price, Competitor_Price, Price_Difference (наша цена минус цена конкурента, округлить до 2 знаков), Difference_Pct (округлить Price_Difference делённое на Competitor_Price умноженное на 100 до 1 знака) и Recommendation ("Reduce price", если мы более чем на 15 процентов выше конкурента, "Maintain", если в пределах 15 процентов, "Consider increase", если более чем на 15 процентов ниже конкурента). Отсортируйте по Product_Name в алфавитном порядке.

Второй лист Category_Summary должен агрегировать данные по категориям товаров из магазина со столбцами Category, Product_Count, Avg_Our_Price (округлить до 2 знаков) и Avg_Stock.

Третий лист Executive_Summary должен содержать столбцы Metric и Value со значениями Total_Products_Compared, Products_Overpriced (более чем на 15 процентов выше), Products_Competitive (в пределах 15 процентов), Products_Underpriced (более чем на 15 процентов ниже), Avg_Price_Gap (среднее по всем Difference_Pct, округлить до 1 знака).

Также создайте Word-документ Pricing_Strategy.docx с заголовком "Pricing Strategy Recommendations" и разделами Market Position Analysis, Product-Level Recommendations и Implementation Timeline, минимум по 2 предложения в каждом.

Примечание: при работе с инструментами Excel и Word передавайте абсолютный путь к файлам .xlsx и .docx внутри рабочей директории агента (MCP-серверы excel и word не сохраняют текущую рабочую директорию).
