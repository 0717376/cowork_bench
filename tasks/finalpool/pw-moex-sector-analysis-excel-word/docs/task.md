Помоги мне с отраслевым анализом акций нашего портфеля на Московской бирже (yf sector analysis). Внешние эталонные (benchmark) данные по отраслям доступны на странице http://localhost:30315 — открой её и извлеки нужные показатели (колонки Sector, Benchmark_PE, Benchmark_Yield_Pct, Market_Outlook).

Затем получи актуальные рыночные данные через финансовый MCP (moex-finance) для акций нашего портфеля. Портфель состоит из шести тикеров MOEX (валюта — RUB):

- SBER.ME
- GAZP.ME
- LKOH.ME
- TCSG.ME
- MGNT.ME
- MTSS.ME

По каждому тикеру возьми текущую цену (Current_Price), отрасль (Sector) и название компании (Name) из данных moex-finance. Целевую цену (Target_Price) рассчитай как справедливую оценку на основе эталонного отраслевого мультипликатора Benchmark_PE с указанной странице (сопоставляя тикер с его отраслью). Апсайд (Upside) посчитай как отклонение целевой цены от текущей в процентах: Upside = (Target_Price - Current_Price) / Current_Price * 100.

Через терминал создай и запусти Python-скрипт с именем yf_sector_processor.py в рабочей директории. Скрипт должен прочитать собранные тобой данные из JSON-файлов (которые ты сам создашь), выполнить анализ и сохранить результат в yf_sector_results.json.

Создай Excel-файл Sector_Analysis_Report.xlsx с тремя листами:

1. Лист Data_Analysis — основная сравнительная таблица. Обязательные колонки (заголовки именно на английском): Symbol, Name, Sector, Current_Price, Target_Price, Upside. Одна строка на каждый тикер портфеля (минимум 5 строк). Отсортируй строки по алфавиту по колонке Symbol.

2. Лист Metrics — сводка ключевых метрик. Ровно две колонки: Metric и Value. Включи итоговые показатели, в том числе: Total_Stocks (число акций), Avg_Upside (средний Upside по всем строкам Data_Analysis), Best_Opportunity (Symbol с максимальным Upside). Минимум 3 строки.

3. Лист Recommendations — список приоритетных действий по результатам анализа апсайда. Колонки: Priority, Action, Symbol. В колонке Symbol указывай тикеры, присутствующие на листе Data_Analysis. Минимум 2 строки.

Также создай документ Word с именем Sector_Analysis_Analysis.docx, содержащий три раздела: краткое резюме (executive summary), ключевые выводы (key findings) и рекомендации (recommendations).

Важно: MCP-инструменты excel и word не сохраняют рабочую директорию между вызовами. Передавай им АБСОЛЮТНЫЕ пути к файлам Sector_Analysis_Report.xlsx и Sector_Analysis_Analysis.docx внутри рабочей директории.
