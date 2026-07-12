Мне нужно подготовить комплексный отчёт по секторному анализу Московской биржи (MOEX), объединяющий финансовые данные с выводами из академических исследований. Соберите финансовые данные по акциям из разных секторов, обращая внимание на их секторную классификацию, рыночную динамику и ключевые финансовые показатели. Используйте следующие тикеры MOEX: SBER.ME, GAZP.ME, LKOH.ME, TCSG.ME, MGNT.ME, MTSS.ME. Цены указаны в рублях (RUB).

Также найдите научные статьи, связанные с темами «sector rotation», «industry analysis» и «market cycles», чтобы учесть академические взгляды на закономерности динамики секторов.

С помощью терминала создайте и запустите в рабочей директории Python-скрипт с именем sector_analyst.py, который читает financial_data.json и research_findings.json (сначала создайте оба файла), анализирует показатели динамики по секторам, сопоставляет академические выводы с реальными рыночными данными и выводит результат в sector_analysis.json.

Создайте Excel-файл с именем Sector_Analysis_Report.xlsx с тремя листами. Первый лист Sector_Performance должен содержать столбцы Sector, Stock_Count, Avg_Price (округлить до 2 знаков), Total_Market_Value (округлить до 2 знаков) и Volatility_Score (округлить до 2 знаков), отсортированные по Sector. Total_Market_Value должен равняться Avg_Price, умноженному на Stock_Count. Второй лист Research_Mapping должен содержать столбцы Paper_Title, Key_Finding, Applicable_Sector, Validation_Status («Confirmed», «Partial» или «Inconclusive»), отсортированные по Paper_Title. Третий лист Investment_Thesis должен содержать столбцы Sector, Outlook («Bullish», «Neutral» или «Bearish»), Supporting_Evidence и Risk_Factor.

Создайте документ Word с именем Sector_Research_Brief.docx с заголовком «Cross-Disciplinary Sector Analysis» и разделами «Financial Performance Review», «Academic Research Insights», «Theory vs Practice Comparison» и «Investment Implications» с конкретными данными и ссылками на статьи.

Важно: при работе с инструментами Excel и Word передавайте АБСОЛЮТНЫЙ путь к файлам .xlsx и .docx внутри рабочей директории (Excel/Word MCP теряет текущую директорию).
