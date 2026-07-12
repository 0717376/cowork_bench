Ваша научно-исследовательская лаборатория (НИИ прикладных исследований) готовит к подаче в журнал три рукописи по результатам трёх завершённых исследований. Вам нужно подготовить пакет документов для отправки: оформить рукописи на основе имеющихся данных и заполнить контрольный список (checklist) готовности к подаче. Все рабочие файлы находятся в рабочем каталоге агента; исходные данные лежат в файле `research_data.xlsx`, а шаблоны рисунков — в `figure_templates.pptx`.

ВАЖНО: инструменты для работы с Word (.docx) и PowerPoint (.pptx) теряют текущий рабочий каталог. Поэтому при каждом вызове передавайте ПОЛНЫЙ АБСОЛЮТНЫЙ путь к файлу внутри рабочего каталога агента (например, `/.../agent_workspace/manuscript_paper1.docx`).

Фаза 1. Изучите исходные данные. Откройте `research_data.xlsx` в рабочем каталоге. В нём три листа: `Paper1_Results`, `Paper2_Analysis`, `Paper3_Experiments`. Каждый лист содержит результаты соответствующего исследования (метрики, средние значения, p-значения, размеры выборок и т.д.). Внимательно изучите числовые значения — они должны быть отражены в тексте рукописей. Сохраняйте названия листов и заголовки столбцов в исходном (английском) виде.

Фаза 2. Подготовьте три рукописи. Создайте в рабочем каталоге три документа Word: `manuscript_paper1.docx`, `manuscript_paper2.docx`, `manuscript_paper3.docx`. Каждая рукопись должна содержать на русском языке полноценные разделы: заголовок, аннотацию (Аннотация / Abstract объёмом не менее 150 слов), Введение, Методы, Результаты, Обсуждение, Заключение и Список литературы. Текст разделов «Результаты» и «Методы» должен опираться на конкретные числовые значения из соответствующего листа `research_data.xlsx`. В частности:
  - в `manuscript_paper1.docx` укажите размер выборки (Sample Size = 156) и первичный исход (Primary Outcome, Value = 0.847; p = 0.0023);
  - в `manuscript_paper2.docx` отразите сравнение групп Baseline Group и Treatment Group (средние 45.2 и 52.8) и размеры выборок;
  - в `manuscript_paper3.docx` укажите результаты экспериментов (например, успешность Experiment 1 = 87.5%, число попыток).
Аннотация и разделы пишутся связной русской прозой; не оставляйте документы пустыми или с одним абзацем-заглушкой.

Фаза 3. Заполните контрольный список подачи. Создайте в рабочем каталоге файл `submission_checklist.xlsx` с листом `Submission Checklist`. На листе разместите таблицу со строкой-заголовком, содержащей ровно столбцы: `Item`, `Status`, `Notes`, `Reviewer` (названия столбцов и значения статусов оставляйте на английском). Заполните пункты проверки готовности рукописей к подаче. Используйте следующие пункты (значения столбца `Status` — литералы `Complete` или `In Progress`; в столбце `Notes` допускается русская прозаическая заметка):

  - `Manuscript formatting` — `Complete` — Reviewer `Editor`
  - `Title page included` — `Complete` — Reviewer `Editor`
  - `Abstract (150-200 words)` — `Complete` — Reviewer `Editor`
  - `All figures included` — `Complete` — Reviewer `Editor`
  - `All tables included` — `Complete` — Reviewer `Editor`
  - `References formatted` — `Complete` — Reviewer `Editor`
  - `Author contributions` — `Complete` — Reviewer `Editor`
  - `Conflict of interest` — `Complete` — Reviewer `Editor`
  - `Funding sources` — `Complete` — Reviewer `Editor`
  - `Supplementary materials` — `Complete` — Reviewer `Editor`
  - `Copyright permissions` — `In Progress` — Reviewer `Legal`
  - `Institutional approval` — `Complete` — Reviewer `IRB`
  - `Data availability` — `Complete` — Reviewer `Data Manager`
  - `Ethical compliance` — `Complete` — Reviewer `Ethics Committee`
  - `Plagiarism check` — `Complete` — Reviewer `QA`

Сохраняйте именно такой порядок пунктов и значения столбцов `Item`, `Status`, `Reviewer` на английском; заметки в столбце `Notes` можно писать по-русски.

По завершении подготовки всех файлов задача считается выполненной. Завершите работу, ответив без вызова инструментов.
