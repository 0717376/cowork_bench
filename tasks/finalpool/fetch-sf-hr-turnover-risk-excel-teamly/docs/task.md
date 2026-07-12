Мне нужно оценить риск текучести персонала по всей организации. По адресу http://localhost:30302/api/turnover_benchmarks.json доступен API-эндпоинт, который предоставляет отраслевые ориентиры (benchmarks) по текучести и пороги риска в разрезе подразделений от Института HR-аналитики. Пожалуйста, получи эти данные.

Затем выгрузи данные о наших сотрудниках из корпоративного хранилища данных ClickHouse, чтобы получить статистику на уровне подразделений: средние оценки удовлетворённости работой, уровни зарплат, опыт и численность.

Через terminal напиши и запусти Python-скрипт под названием risk_scorer.py в рабочей директории. Скрипт должен прочитать файл combined_data.json (который ты создаёшь, объединив оба источника данных), рассчитать оценки риска по каждому подразделению и записать результат в risk_assessment.json. Оценка риска должна сравнивать наши уровни удовлетворённости с отраслевыми порогами.

Создай Excel-файл под названием Turnover_Risk_Assessment.xlsx с тремя листами. Записывай его по абсолютному пути в рабочей директории. Первый лист Risk_Overview должен содержать столбцы Department, Employee_Count, Avg_Salary, Avg_Satisfaction, Industry_Turnover_Rate, Risk_Threshold и Risk_Level ("High", если Avg_Satisfaction ниже порога риска; "Medium", если в пределах 0.5 выше порога; иначе "Low"). Отсортируй по Department в алфавитном порядке.

Второй лист Risk_Summary должен содержать два столбца Metric и Value со строками: Total_Departments, High_Risk_Count, Medium_Risk_Count, Low_Risk_Count, Highest_Risk_Department (подразделение с наименьшей удовлетворённостью относительно порога), Total_At_Risk_Employees (сумма сотрудников в подразделениях с риском High и Medium).

Третий лист Detailed_Metrics должен содержать столбцы Department, Avg_Experience, Avg_Performance, Satisfaction_Gap (Avg_Satisfaction минус Risk_Threshold, округлённое до 2 знаков) и Estimated_Turnover_Cost (Employee_Count умножить на Avg_Salary, умножить на Industry_Turnover_Rate, делить на 100, округлить до ближайшего целого).

Также создай страницу базы знаний в Teamly под названием "Turnover Risk Dashboard" с кратким изложением выводов, включая ключевые метрики и рекомендуемые действия для подразделений высокого риска.
