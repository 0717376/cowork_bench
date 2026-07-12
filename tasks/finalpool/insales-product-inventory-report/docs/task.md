Мне нужен полный отчёт по складским остаткам нашего магазина InSales. В рабочей директории лежит файл Inventory_Guidelines.pdf с нашими правилами учёта остатков. Выгрузи все товары вместе с информацией о ценах.

Создай Excel-файл с именем WC_Inventory_Report.xlsx с двумя листами. На листе "Inventory Report" должны быть колонки: Product (название товара), Type, Regular_Price, Sale_Price (0, если товар не на распродаже), Stock_Qty и Stock_Status. Отсортируй строки по Regular_Price по убыванию.

На листе "Summary" должны быть метрики: Total_Products, Avg_Price (округлённое до 2 знаков), On_Sale_Count и Out_Of_Stock.

Важно: при сохранении Excel-файла указывай АБСОЛЮТНЫЙ путь к WC_Inventory_Report.xlsx внутри рабочей директории, так как Excel MCP не сохраняет текущий рабочий каталог.
