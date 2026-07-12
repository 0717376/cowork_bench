Мне нужно выявить студентов, находящихся под угрозой академической неуспеваемости, и составить план вмешательства (intervention plan). Извлеките данные об оценках студентов и сдаче работ из системы управления обучением (canvas). Проанализируйте баллы за тесты, процент выполнения заданий и итоговые оценки по курсам.

Через терминал создайте и запустите Python-скрипт `risk_identifier.py` в рабочей директории. Скрипт должен прочитать `student_performance.json` (сначала создайте его), пометить студентов со средним баллом ниже 50 как "Critical", от 50 до 65 как "At Risk", и выше 65 как "On Track", рассчитать статистику рисков по каждому курсу и записать результат в `risk_assessment.json`.

Создайте Excel-файл `Student_Risk_Assessment.xlsx` с четырьмя листами. Первый лист `Risk_Overview` должен содержать столбцы `Course_Name`, `Total_Students`, `Critical_Count`, `At_Risk_Count`, `On_Track_Count` и `Risk_Rate_Pct` (округление до 1 знака после запятой, процент студентов категорий Critical + At Risk), отсортированные по `Risk_Rate_Pct` по убыванию. Второй лист `Critical_Students` должен содержать столбцы `Student_ID`, `Course_Name`, `Avg_Score` (округление до 1 знака), `Assignments_Submitted` и `Late_Submissions` для всех студентов категории Critical. Третий лист `Intervention_Plan` должен содержать столбцы `Course_Name`, `Risk_Level`, `Recommended_Action`, `Responsible_Party` и `Deadline`. Четвёртый лист `Summary` должен содержать столбцы `Metric` и `Value` со значениями `Total_Students_Assessed`, `Critical_Students`, `At_Risk_Students`, `On_Track_Students`, `Overall_Risk_Rate_Pct` (округление до 1 знака) и `Highest_Risk_Course`.

Создайте Word-документ `Intervention_Report.docx` с заголовком "Student Risk Intervention Report" и разделами "Risk Assessment Methodology", "Course-Level Analysis", "Critical Cases" и "Recommended Interventions" с конкретными данными.

Запланируйте событие в календаре "Academic Intervention Planning Meeting" на 13 марта 2026 года с 15:00 до 16:30 UTC с описанием, в котором перечислены курсы с наибольшими показателями риска (risk rate).

Отправьте письмо на адрес academic-affairs@university.edu с темой "Urgent: Student Risk Assessment Results", выделив курсы с показателем риска выше 40%.

Важно: при работе с инструментами Excel и Word указывайте АБСОЛЮТНЫЙ путь к файлам `Student_Risk_Assessment.xlsx` и `Intervention_Report.docx` внутри рабочей директории, так как эти инструменты не сохраняют текущую рабочую директорию.
