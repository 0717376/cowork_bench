Мне нужен анализ категорий товаров из нашей витрины данных продаж в ClickHouse. Посмотри таблицу products (схема SALES_DW.PUBLIC.PRODUCTS) и сгруппируй товары по категории (CATEGORY), чтобы понять ценообразование и прибыльность.

Создай файл Excel с именем Sales_Product_Categories.xlsx с двумя листами. Лист "Product Categories" должен содержать столбцы Category, Product_Count, Avg_Price, Avg_Cost и Avg_Margin (цена минус себестоимость), все округлённые до 2 знаков после запятой. Отсортируй по Avg_Margin по убыванию.

Лист "Summary" должен содержать Total_Categories, Total_Products, Most_Profitable_Category и Overall_Avg_Margin как средневзвешенное значение по количеству товаров, округлённое до 2 знаков после запятой.

Также зафиксируй сводку по категориям в Google-таблице с названием "Product Category Report" для команды мерчандайзинга.

Важно: при работе с Excel-инструментом указывай АБСОЛЮТНЫЙ путь к файлу Sales_Product_Categories.xlsx внутри рабочей директории, поскольку относительные пути могут не сработать.
