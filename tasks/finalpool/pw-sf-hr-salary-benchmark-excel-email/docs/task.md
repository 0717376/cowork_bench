Мне нужно сравнить зарплаты нашей компании с актуальными отраслевыми данными. По адресу http://localhost:30301 опубликован отчёт о зарплатных бенчмарках за 2026 год от CompAnalytics Research, который нужно проверить. Зайди на эту страницу и выгрузи все эталонные значения зарплат в разрезе отделов.

Затем обратись к нашему внутреннему хранилищу данных HR в ClickHouse и получи фактические средние зарплаты по отделам, а также количество сотрудников, средний стаж (в годах) и средний рейтинг производительности.

Прежде чем формировать итоговый отчёт, с помощью терминала напиши и запусти Python-скрипт salary_processor.py в рабочем каталоге. Скрипт должен читать JSON-файл benchmark_raw.json (создай его на основе данных с веб-страницы), сопоставлять их с файлом internal_salaries.json (создай его на основе данных из хранилища) и выводить файл salary_comparison.json с объединённым анализом. При работе с терминалом и скриптом указывай АБСОЛЮТНЫЙ путь к рабочему каталогу для всех файлов (salary_processor.py, benchmark_raw.json, internal_salaries.json, salary_comparison.json).

Создай Excel-файл Salary_Benchmark_Report.xlsx с тремя листами. Первый лист Compensation_Comparison должен содержать столбцы Department, Employee_Count, Our_Avg_Salary, Industry_Benchmark, Difference (наша зарплата минус отраслевая, округлить до 2 знаков), Difference_Pct (Difference, делённое на Industry_Benchmark, умноженное на 100, округлить до 1 знака) и Status (написать "Above", если Difference >= 0, иначе "Below"). Отсортируй по алфавиту по столбцу Department.

Второй лист Department_Details должен содержать столбцы Department, Avg_Experience, Avg_Performance, Our_Avg_Salary и столбец Salary_Per_Year_Exp, вычисляемый как Our_Avg_Salary, делённое на Avg_Experience, округлённое до 2 знаков. Отсортируй по алфавиту.

Третий лист Executive_Summary должен содержать два столбца Metric и Value со строками: Total_Departments, Departments_Above_Benchmark, Departments_Below_Benchmark, Highest_Gap_Department (отдел с наибольшим положительным Difference), Lowest_Gap_Department (отдел с наибольшим отрицательным Difference), Average_Difference (среднее всех значений Difference, округлённое до 2 знаков), Overall_Status ("Competitive", если отделов выше бенчмарка больше, чем ниже, иначе "Needs Attention").

Также отправь письмо на адрес hr-director@company.com с темой "2026 Salary Benchmark Analysis Complete" и телом, кратко суммирующим ключевые выводы: сколько отделов выше и сколько ниже бенчмарка, у какого отдела наибольший разрыв и каков общий статус.

При создании Excel-файла указывай АБСОЛЮТНЫЙ путь к Salary_Benchmark_Report.xlsx в рабочем каталоге.
