"""Canonical English->Russian relabel map for FULL russification of the Canvas LMS data.

Canvas is NOT a fork (service identity stays "canvas"); this is data-layer russification
only. The seed `canvas` schema (served by mcp-canvas-lms with CANVAS_USE_PG=1) holds
English OULAD-derived realia: person names, UK regions, education bands, faculty roles,
course/department titles, assignment/quiz descriptions, module/page/discussion/announcement
text, tutor comments. We russify all human-readable DATA VALUES and freeze identifiers,
applying the SAME map to seed + eval literals + frozen groundtruth so they stay in sync.

SINGLE SOURCE OF TRUTH consumed by:
  - scripts/canvas_gen_seed_sql.py    (db/zzz_canvas_after_init.sql UPDATEs over canvas.*)
  - scripts/canvas_patch_groundtruth.py (frozen xlsx/json/docx cells + eval .py literals)

All distinct values below were extracted DIRECTLY from the live cowork_pg canvas schema
(no guessing). See scripts/canvas_fork_analysis.md for the KEEP/RUSSIFY split per column.

================================ KEEP (NOT mapped) =============================
  - users.id / login_id / email (@openuniversity.ac.uk) / avatar_url / locale / time_zone / created_at
  - courses.course_code (AAA-2013J...) / account_id / workflow_state / term_id / dates / total_students
  - assignments.name (TMA/CMA/Final Exam + number = codes) / grading_type / submission_types / html_url / FKs / due_at / points
  - quizzes.title (CMA NNNNN codes) ; quiz_questions.question_name (QNN codes)
  - submissions.* except submission_comments (body is 100% NULL); grade/workflow_state/url/flags KEEP
  - all canvas.* table/column identifiers, jsonb keys, comment author_name "Tutor" -> "Преподаватель" (display)
===============================================================================
"""

# ----------------------------------------------------------------------------
# Person-name atoms (gender-tagged). 80 distinct first names, 50 distinct last.
# ----------------------------------------------------------------------------
FIRST_F = {
    "Abigail": "Аделина", "Alice": "Алиса", "Amelia": "Амалия", "Amy": "Алёна",
    "Charlotte": "Карина", "Chloe": "Кристина", "Clara": "Клара", "Daisy": "Дарья",
    "Eleanor": "Элеонора", "Ella": "Элла", "Emily": "Эмилия", "Emma": "Эмма",
    "Evelyn": "Эвелина", "Fiona": "Фаина", "Freya": "Фрида", "Grace": "Галина",
    "Hannah": "Анна", "Harper": "Арина", "Holly": "Олеся", "Isabella": "Изабелла",
    "Isla": "Ирина", "Jessica": "Жанна", "Katie": "Екатерина", "Laura": "Лариса",
    "Lily": "Лилия", "Lucy": "Любовь", "Martha": "Марфа", "Mia": "Марина",
    "Molly": "Милана", "Natasha": "Наталья", "Nora": "Нина", "Olivia": "Ольга",
    "Rachel": "Раиса", "Rebecca": "Регина", "Ruby": "Римма", "Sarah": "Серафима",
    "Sophia": "София", "Sophie": "Софья", "Victoria": "Виктория", "Zoe": "Зоя",
}
FIRST_M = {
    "Aiden": "Артём", "Alex": "Александр", "Andrew": "Андрей", "Ben": "Богдан",
    "Caleb": "Кирилл", "Charlie": "Степан", "Connor": "Константин", "Daniel": "Даниил",
    "David": "Давид", "Dylan": "Денис", "Edward": "Эдуард", "Ethan": "Егор",
    "Evan": "Иван", "Finn": "Фёдор", "George": "Георгий", "Harry": "Григорий",
    "Henry": "Геннадий", "Isaac": "Игорь", "Jack": "Яков", "James": "Дмитрий",
    "Joseph": "Иосиф", "Leo": "Лев", "Liam": "Леонид", "Logan": "Роман",
    "Lucas": "Лука", "Mason": "Матвей", "Matthew": "Марк", "Michael": "Михаил",
    "Nathan": "Назар", "Noah": "Никита", "Oliver": "Олег", "Oscar": "Остап",
    "Owen": "Антон", "Patrick": "Пётр", "Robert": "Ростислав", "Ryan": "Руслан",
    "Samuel": "Савелий", "Sean": "Святослав", "Thomas": "Тимофей", "William": "Владимир",
}
FIRST = {**FIRST_F, **FIRST_M}

# Last names stored MASCULINE; feminine = base + "а" (all end -ов/-ев/-ёв/-ин -> regular).
LAST = {
    "Adams": "Соколов", "Allen": "Морозов", "Anderson": "Андреев", "Bailey": "Беляев",
    "Baker": "Баранов", "Bell": "Белов", "Brown": "Зайцев", "Campbell": "Кошелев",
    "Carter": "Карпов", "Clark": "Климов", "Collins": "Колесников", "Cook": "Гончаров",
    "Davis": "Давыдов", "Edwards": "Егоров", "Evans": "Иванов", "Green": "Зеленин",
    "Hall": "Гаврилов", "Harris": "Харитонов", "Jackson": "Якушев", "Johnson": "Юдин",
    "Jones": "Жуков", "King": "Королёв", "Lee": "Лебедев", "Lewis": "Львов",
    "Martin": "Мартынов", "Miller": "Мельников", "Mitchell": "Михеев", "Moore": "Муравьёв",
    "Morgan": "Моргунов", "Morris": "Маслов", "Murphy": "Мурашов", "Nelson": "Нестеров",
    "Parker": "Панкратов", "Phillips": "Филиппов", "Reed": "Рудин", "Rivera": "Рябов",
    "Roberts": "Романов", "Rogers": "Рогов", "Scott": "Скворцов", "Smith": "Смирнов",
    "Stewart": "Стариков", "Taylor": "Портнов", "Thomas": "Фомин", "Turner": "Токарев",
    "Walker": "Волков", "White": "Седов", "Williams": "Васильев", "Wilson": "Власов",
    "Wright": "Плотников", "Young": "Молодцов",
}

# ----------------------------------------------------------------------------
# Realia value sets
# ----------------------------------------------------------------------------
REGIONS = {
    "London Region": "Московский регион", "Ireland": "Краснодарский край",
    "North Region": "Северный регион", "Wales": "Республика Татарстан",
    "South East Region": "Юго-Восточный регион", "East Midlands Region": "Поволжский регион",
    "Scotland": "Уральский регион", "North Western Region": "Северо-Западный регион",
    "South West Region": "Юго-Западный регион", "West Midlands Region": "Центральный регион",
    "Yorkshire Region": "Сибирский регион", "South Region": "Южный регион",
    "East Anglian Region": "Восточный регион",
}
EDU_LEVELS = {
    "HE Qualification": "Высшее образование",
    "Post Graduate Qualification": "Послевузовское образование",
    "No Formal quals": "Без формального образования",
    "Lower Than A Level": "Неполное среднее",
    "A Level or Equivalent": "Среднее общее",
}
ROLES = {  # full bio strings (faculty)
    "Faculty member, Professor": "Преподаватель, профессор",
    "Faculty member, Teaching Assistant": "Преподаватель, ассистент",
    "Faculty member, Senior Lecturer": "Преподаватель, старший преподаватель",
    "Faculty member, Associate Professor": "Преподаватель, доцент",
    "Faculty member, Lecturer": "Преподаватель, лектор",
}
DEPARTMENTS = {
    "Business": "Бизнес", "Computer Science": "Информатика",
    "Digital Arts": "Цифровое искусство", "Life Sciences": "Естественные науки",
    "Engineering": "Инженерия", "Political Science": "Политология",
    "Social Sciences": "Социальные науки",
}
COURSE_SUBJECTS = {
    "Applied Analytics & Algorithms": "Прикладная аналитика и алгоритмы",
    "Biochemistry & Bioinformatics": "Биохимия и биоинформатика",
    "Creative Computing & Culture": "Креативные вычисления и культура",
    "Data-Driven Design": "Проектирование на основе данных",
    "Environmental Economics & Ethics": "Экологическая экономика и этика",
    "Foundations of Finance": "Основы финансов",
    "Global Governance & Geopolitics": "Глобальное управление и геополитика",
}
SEASONS = {"Fall": "Осень", "Spring": "Весна"}

# Assignment / quiz descriptions (templates; weight number preserved)
ASSIGNMENT_TYPES = {
    "Tutor Marked Assessment": "Оценка наставником",
    "Computer Marked Assessment": "Компьютерная проверка",
    "Final Examination": "Итоговый экзамен",
}
ASSIGNMENT_GROUPS = {
    "Final Exam": "Итоговый экзамен",
    "Assignments (TMA)": "Задания (TMA)",
    "Quizzes (CMA)": "Тесты (CMA)",
}

# Pages / discussion / announcement finite phrase + body sets
PAGE_TITLES = {
    "Academic Integrity": "Академическая честность",
    "Additional Resources": "Дополнительные материалы",
    "Assessment Guide": "Руководство по оцениванию",
    "Getting Started": "Начало работы",
    "Syllabus": "Программа курса",
}
DISCUSSION_PHRASES = {
    "Assignment Questions": "Вопросы по заданиям",
    "Exam Preparation": "Подготовка к экзамену",
    "General Discussion": "Общее обсуждение",
    "Group Project Coordination": "Координация группового проекта",
    "Introduce Yourself": "Представьтесь",
    "Study Tips & Resources": "Советы и материалы для учёбы",
}
ANNOUNCEMENT_PHRASES = {
    "Course Evaluation": "Оценка курса",
    "Exam Preparation Tips": "Советы по подготовке к экзамену",
    "Field Trip Notice - Qufu Visit March 12-15": "Объявление о выездном занятии — поездка в Суздаль 12–15 марта",
    "Grade Update": "Обновление оценок",
    "Office Hours Change": "Изменение часов консультаций",
    "Reminder: CMA Deadline": "Напоминание: дедлайн CMA",
    "Study Group Sessions": "Занятия учебных групп",
    "TMA 1 Released": "Опубликовано задание TMA 1",
    "Welcome to the Course": "Добро пожаловать на курс",
}
DISCUSSION_MSGS = {
    "A space for general course-related discussion.":
        "Пространство для общего обсуждения по курсу.",
    "Discuss exam prep strategies and share practice materials.":
        "Обсуждайте стратегии подготовки к экзамену и делитесь учебными материалами.",
    "Post any questions about the current assignments here.":
        "Задавайте здесь любые вопросы по текущим заданиям.",
    "Share a bit about yourself and your goals for this module.":
        "Расскажите немного о себе и своих целях в этом модуле.",
    "Share any helpful study tips or resources you've found.":
        "Делитесь полезными советами и материалами для учёбы, которые нашли.",
    "Use this thread to coordinate with your study group.":
        "Используйте эту ветку для координации с вашей учебной группой.",
}
ANNOUNCEMENT_MSGS = {
    "A friendly reminder that the computer-marked assessment is due this week.":
        "Напоминаем, что компьютерная проверка должна быть сдана на этой неделе.",
    "Grades for the recent assessment have been posted. Please check your feedback.":
        "Оценки за недавнюю работу опубликованы. Пожалуйста, проверьте обратную связь.",
    "Office hours this week will be moved to Thursday 2-4pm.":
        "Часы консультаций на этой неделе перенесены на четверг с 14:00 до 16:00.",
    "Please complete the end-of-module evaluation survey by the end of the week.":
        "Пожалуйста, пройдите итоговый опрос по модулю до конца недели.",
    "The first Tutor Marked Assessment has been published. Please check the due date carefully.":
        "Опубликовано первое задание с оценкой наставником. Внимательно проверьте срок сдачи.",
    "Train G235 departs Beijing South 17:30 on March 12. Cost: CNY 1106 round trip.":
        "Поезд №036 отправляется с Курского вокзала в 17:30 12 марта. Стоимость: 4200 ₽ туда-обратно.",
    "Weekly study group sessions are now available. Sign up through the collaboration tool.":
        "Открыта запись на еженедельные занятия учебных групп. Запишитесь через инструмент совместной работы.",
    "Welcome everyone! Please review the syllabus and introduce yourself in the discussion forum.":
        "Добро пожаловать! Пожалуйста, изучите программу курса и представьтесь на форуме обсуждений.",
    "With the exam approaching, here are some key topics to review...":
        "Приближается экзамен — вот несколько ключевых тем для повторения...",
}
QUIZ_DESCRIPTIONS = {
    "Quiz on TypeScript concepts": "Тест по концепциям TypeScript",
}
PAGE_BODIES = {
    "<h1>Assessment Guide</h1><p>Overview of TMAs, CMAs, and exams.</p>":
        "<h1>Руководство по оцениванию</h1><p>Обзор TMA, CMA и экзаменов.</p>",
    "<h1>Academic Integrity</h1><p>University policies on plagiarism.</p>":
        "<h1>Академическая честность</h1><p>Политика университета в отношении плагиата.</p>",
    "<h1>Getting Started</h1><p>How to navigate this course.</p>":
        "<h1>Начало работы</h1><p>Как ориентироваться в этом курсе.</p>",
    "<h1>Resources</h1><p>Recommended reading and materials.</p>":
        "<h1>Материалы</h1><p>Рекомендованная литература и материалы.</p>",
}
SUBMISSION_COMMENTS = {
    "Below expectations. Please attend office hours.":
        "Ниже ожиданий. Пожалуйста, посетите консультацию.",
    "Decent attempt. Focus on methodology next time.":
        "Неплохая попытка. В следующий раз сосредоточьтесь на методологии.",
    "Excellent work!": "Отличная работа!",
    "Good effort, some areas need improvement.":
        "Хорошая работа, но некоторые моменты требуют доработки.",
    "Great understanding of the material.": "Прекрасное понимание материала.",
    "Keep up the excellent work.": "Так держать, отличная работа.",
    "Outstanding analysis.": "Выдающийся анализ.",
    "Please review the material and consider resubmitting.":
        "Пожалуйста, изучите материал и рассмотрите возможность повторной сдачи.",
    "Satisfactory work. Review feedback carefully.":
        "Удовлетворительная работа. Внимательно изучите обратную связь.",
    "Several key concepts were missed. See detailed feedback.":
        "Упущено несколько ключевых концепций. Смотрите подробную обратную связь.",
    "Very well done.": "Очень хорошо.",
}

# ----------------------------------------------------------------------------
# Name composition
# ----------------------------------------------------------------------------
TITLES = {"Dr.": "Д-р", "Prof.": "Проф.", "Mr.": "Г-н", "Ms.": "Г-жа", "Mrs.": "Г-жа"}


def _gender(first_en):
    if first_en in FIRST_F:
        return "f"
    if first_en in FIRST_M:
        return "m"
    return None


def _feminize(last_ru_masc):
    return last_ru_masc + "а"


def map_last(last_en, gender):
    base = LAST.get(last_en)
    if base is None:
        return None
    return _feminize(base) if gender == "f" else base


def map_full_name(name):
    """'[Title] First Last' -> russified, gender-agreed surname.

    Returns the RU string, or None if no first/last atom matched (leave as-is).
    Title (Dr./Prof./...) is mapped if present; unknown middle tokens pass through.
    """
    if not name:
        return None
    toks = name.split()
    title = None
    if toks and toks[0] in TITLES:
        title = TITLES[toks[0]]
        toks = toks[1:]
    if len(toks) < 2:
        return None
    first_en, last_en = toks[0], toks[-1]
    g = _gender(first_en)
    rf = FIRST.get(first_en)
    rl = map_last(last_en, g)
    if rf is None and rl is None:
        return None
    out = f"{rf or first_en} {rl or last_en}"
    return f"{title} {out}" if title else out


def map_short_name(value):
    """First-name-only field."""
    return FIRST.get(value)


def map_sortable_name(value):
    """'Last, First' -> 'Russifiedlast, Russifiedfirst' (gender from first)."""
    if not value or ", " not in value:
        return None
    last_en, first_en = value.split(", ", 1)
    g = _gender(first_en)
    rf = FIRST.get(first_en)
    rl = map_last(last_en, g)
    if rf is None and rl is None:
        return None
    return f"{rl or last_en}, {rf or first_en}"


# ----------------------------------------------------------------------------
# Structured / template fields
# ----------------------------------------------------------------------------
def map_bio(bio):
    if bio is None:
        return None
    if bio in ROLES:
        return ROLES[bio]
    # "Student from {Region}. Education: {Level}."
    import re
    m = re.match(r"^Student from (.+?)\. Education: (.+?)\.?$", bio)
    if m:
        region = REGIONS.get(m.group(1), m.group(1))
        level = EDU_LEVELS.get(m.group(2), m.group(2))
        # profile-style (no preposition) to avoid declension of 13 regions
        return f"Студент. Регион: {region}. Образование: {level}."
    return None


def _norm_amp(s):
    return s.replace("&amp;", "&") if s else s


def map_course_name(name):
    """'{Subject} ({Season Year})' -> RU, handling & / &amp;. Season translated, year kept."""
    if not name:
        return None
    import re
    n = _norm_amp(name)
    m = re.match(r"^(.+?) \((Fall|Spring) (\d{4})\)$", n)
    if m:
        subj = COURSE_SUBJECTS.get(m.group(1))
        if subj is None:
            return None
        return f"{subj} ({SEASONS[m.group(2)]} {m.group(3)})"
    return COURSE_SUBJECTS.get(n)  # bare subject


def map_public_description(text):
    """'Open University {Subject} - {Season Year}'."""
    if not text:
        return None
    import re
    t = _norm_amp(text)
    m = re.match(r"^Open University (.+?) - (Fall|Spring) (\d{4})$", t)
    if m:
        subj = COURSE_SUBJECTS.get(m.group(1))
        if subj is None:
            return None
        return f"Открытый университет, {subj} — {SEASONS[m.group(2)]} {m.group(3)}"
    return None


def map_syllabus_body(html):
    """'<h2>{Subj}</h2><p>Department: {Dept}</p><p>Term: {Season Year}, Duration: N days</p>
    <p>This module covers key concepts and practical applications in {subj}.</p>'"""
    if not html:
        return None
    import re
    h = _norm_amp(html)
    m = re.match(
        r"^<h2>(.+?)</h2>\s*<p>Department: (.+?)</p>\s*<p>Term: (Fall|Spring) (\d{4}), "
        r"Duration: (\d+) days</p>\s*<p>This module covers key concepts and practical "
        r"applications in (.+?)\.</p>$", h, re.DOTALL)
    if not m:
        return None
    subj = COURSE_SUBJECTS.get(m.group(1))
    dept = DEPARTMENTS.get(m.group(2), m.group(2))
    if subj is None:
        return None
    subj_low = subj[0].lower() + subj[1:]
    return (f"<h2>{subj}</h2>\n<p>Кафедра: {dept}</p>\n"
            f"<p>Семестр: {SEASONS[m.group(3)]} {m.group(4)}, Длительность: {m.group(5)} дн.</p>\n"
            f"<p>Этот модуль охватывает ключевые концепции и практическое применение в области "
            f"«{subj_low}».</p>")


def map_page_body(html):
    """'<h1>Course Syllabus</h1><p>Welcome to {Subject}.</p>' and similar finite bodies."""
    if not html:
        return None
    import re
    h = _norm_amp(html)
    m = re.match(r"^<h1>Course Syllabus</h1><p>Welcome to (.+?)\.</p>$", h)
    if m:
        subj = COURSE_SUBJECTS.get(m.group(1))
        if subj:
            return f"<h1>Программа курса</h1><p>Добро пожаловать на курс «{subj}».</p>"
    return PAGE_BODIES.get(h)


def map_assignment_description(desc):
    """'{Type} (weight: N%)' for Assessment types; 'Final Examination (weight: N%)'."""
    if not desc:
        return None
    import re
    m = re.match(r"^(.+?) \(weight: ([\d.]+)%\)$", desc)
    if m and m.group(1) in ASSIGNMENT_TYPES:
        return f"{ASSIGNMENT_TYPES[m.group(1)]} (вес: {m.group(2)}%)"
    return None


def map_quiz_description(desc):
    """'Auto-graded assessment (weight: N%)' or fixed strings."""
    if not desc:
        return None
    import re
    m = re.match(r"^Auto-graded assessment \(weight: ([\d.]+)%\)$", desc)
    if m:
        return f"Автоматически оцениваемая работа (вес: {m.group(1)}%)"
    return QUIZ_DESCRIPTIONS.get(desc)


def map_module_name(name):
    """'Week N' -> 'Неделя N'; 'Introduction' -> 'Введение'."""
    if not name:
        return None
    import re
    m = re.match(r"^Week (\d+)$", name)
    if m:
        return f"Неделя {m.group(1)}"
    return {"Introduction": "Введение"}.get(name)


def map_assignment_group(name):
    return ASSIGNMENT_GROUPS.get(name)


def map_quiz_question_text(text):
    """'Question N for CMA NNNN' -> 'Вопрос N для CMA NNNN' (CMA code kept)."""
    if not text:
        return None
    import re
    m = re.match(r"^Question (\d+) for CMA (\d+)$", text)
    if m:
        return f"Вопрос {m.group(1)} для CMA {m.group(2)}"
    m = re.match(r"^TypeScript question (\d+)$", text)
    if m:
        return f"Вопрос {m.group(1)} по TypeScript"
    return None


def _bracket_phrase(value, phrase_map):
    """'[XXX] phrase' -> '[XXX] <ru>'; bracket code kept."""
    if not value:
        return None
    import re
    m = re.match(r"^(\[[A-Z]+\] )(.+)$", value)
    if m:
        ru = phrase_map.get(m.group(2))
        return f"{m.group(1)}{ru}" if ru else None
    return phrase_map.get(value)


def map_discussion_title(value):
    return _bracket_phrase(value, DISCUSSION_PHRASES)


def map_announcement_title(value):
    return _bracket_phrase(value, ANNOUNCEMENT_PHRASES)


def map_page_title(value):
    return PAGE_TITLES.get(value)


def map_discussion_message(value):
    return DISCUSSION_MSGS.get(value)


def map_announcement_message(value):
    return ANNOUNCEMENT_MSGS.get(value)


def map_submission_comments(jsonb_val):
    """jsonb list of {comment, created_at, author_name}: russify comment + author display.
    Accepts a python list (already parsed) and returns a NEW list; unknown comments pass through.
    """
    if not isinstance(jsonb_val, list):
        return None
    out = []
    changed = False
    for c in jsonb_val:
        if not isinstance(c, dict):
            out.append(c); continue
        nc = dict(c)
        com = c.get("comment")
        if com in SUBMISSION_COMMENTS:
            nc["comment"] = SUBMISSION_COMMENTS[com]; changed = True
        if c.get("author_name") == "Tutor":
            nc["author_name"] = "Преподаватель"; changed = True
        out.append(nc)
    return out if changed else None


# ----------------------------------------------------------------------------
# Dispatcher: map one (table, column, value). Returns RU string or None (leave as-is).
# Used by canvas_gen_seed_sql.py per column.
# ----------------------------------------------------------------------------
def map_value(table, column, value):
    t, c = table, column
    if t == "users":
        if c == "name":          return map_full_name(value)
        if c == "sortable_name": return map_sortable_name(value)
        if c == "short_name":    return map_short_name(value)
        if c == "bio":           return map_bio(value)
    elif t == "courses":
        if c == "name":               return map_course_name(value)
        if c == "public_description": return map_public_description(value)
        if c == "syllabus_body":      return map_syllabus_body(value)
    elif t == "assignments":
        if c == "description":   return map_assignment_description(value)
    elif t == "quizzes":
        if c == "description":   return map_quiz_description(value)
    elif t == "quiz_questions":
        if c == "question_text": return map_quiz_question_text(value)
    elif t == "modules":
        if c == "name":          return map_module_name(value)
    elif t == "assignment_groups":
        if c == "name":          return map_assignment_group(value)
    elif t == "pages":
        if c == "title":         return map_page_title(value)
        if c == "body":          return map_page_body(value)
    elif t == "discussion_topics":
        if c == "title":         return map_discussion_title(value)
        if c == "message":       return map_discussion_message(value)
    elif t == "announcements":
        if c == "title":         return map_announcement_title(value)
        if c == "message":       return map_announcement_message(value)
    return None


# ----------------------------------------------------------------------------
# FLAT_VALUE_MAP: whole-cell English->RU for eval-literal / groundtruth-cell patching.
# Covers all standalone realia EXCEPT person names (compose via map_full_name) and the
# 22 course names (compose via map_course_name; both & and &amp; forms added below).
# ----------------------------------------------------------------------------
FLAT_VALUE_MAP = {}
for _m in (REGIONS, EDU_LEVELS, ROLES, DEPARTMENTS, COURSE_SUBJECTS,
           ASSIGNMENT_GROUPS, PAGE_TITLES, PAGE_BODIES, DISCUSSION_PHRASES, ANNOUNCEMENT_PHRASES,
           DISCUSSION_MSGS, ANNOUNCEMENT_MSGS, QUIZ_DESCRIPTIONS, SUBMISSION_COMMENTS):
    FLAT_VALUE_MAP.update(_m)
# subject names in both & and &amp; entity forms
for _en, _ru in COURSE_SUBJECTS.items():
    FLAT_VALUE_MAP[_en] = _ru
    FLAT_VALUE_MAP[_en.replace("&", "&amp;")] = _ru
# the 22 concrete course names actually present (subject x season/year), both & forms
COURSE_VARIANTS = [
    ("Applied Analytics & Algorithms", "Fall", "2013"), ("Applied Analytics & Algorithms", "Fall", "2014"),
    ("Biochemistry & Bioinformatics", "Fall", "2013"), ("Biochemistry & Bioinformatics", "Fall", "2014"),
    ("Biochemistry & Bioinformatics", "Spring", "2013"), ("Biochemistry & Bioinformatics", "Spring", "2014"),
    ("Creative Computing & Culture", "Fall", "2014"), ("Creative Computing & Culture", "Spring", "2014"),
    ("Data-Driven Design", "Fall", "2013"), ("Data-Driven Design", "Fall", "2014"),
    ("Data-Driven Design", "Spring", "2013"), ("Data-Driven Design", "Spring", "2014"),
    ("Environmental Economics & Ethics", "Fall", "2013"), ("Environmental Economics & Ethics", "Fall", "2014"),
    ("Environmental Economics & Ethics", "Spring", "2014"),
    ("Foundations of Finance", "Fall", "2013"), ("Foundations of Finance", "Fall", "2014"),
    ("Foundations of Finance", "Spring", "2013"), ("Foundations of Finance", "Spring", "2014"),
    ("Global Governance & Geopolitics", "Fall", "2013"), ("Global Governance & Geopolitics", "Fall", "2014"),
    ("Global Governance & Geopolitics", "Spring", "2014"),
]
for _subj, _se, _yr in COURSE_VARIANTS:
    _full = f"{_subj} ({_se} {_yr})"
    _ru = map_course_name(_full)
    FLAT_VALUE_MAP[_full] = _ru
    FLAT_VALUE_MAP[_full.replace("&", "&amp;")] = _ru


if __name__ == "__main__":
    # self-check
    assert map_full_name("Dr. Oliver Murphy") == "Д-р Олег Мурашов"
    assert map_full_name("Grace Baker") == "Галина Баранова"
    assert map_full_name("Victoria Smith") == "Виктория Смирнова"
    assert map_sortable_name("Murphy, Oliver") == "Мурашов, Олег"
    assert map_sortable_name("Baker, Grace") == "Баранова, Галина"
    assert map_course_name("Applied Analytics & Algorithms (Fall 2013)") == "Прикладная аналитика и алгоритмы (Осень 2013)"
    assert map_course_name("Foundations of Finance (Spring 2014)") == "Основы финансов (Весна 2014)"
    assert map_bio("Student from Yorkshire Region. Education: HE Qualification.") == "Студент. Регион: Сибирский регион. Образование: Высшее образование."
    assert map_bio("Faculty member, Professor") == "Преподаватель, профессор"
    assert map_assignment_description("Tutor Marked Assessment (weight: 20.0%)") == "Оценка наставником (вес: 20.0%)"
    assert map_module_name("Week 19") == "Неделя 19"
    assert map_quiz_question_text("Question 17 for CMA 24297") == "Вопрос 17 для CMA 24297"
    assert map_discussion_title("[GGG] Study Tips & Resources") == "[GGG] Советы и материалы для учёбы"
    print("OK; FLAT_VALUE_MAP entries:", len(FLAT_VALUE_MAP))
    print(map_syllabus_body("<h2>Applied Analytics & Algorithms</h2><p>Department: Computer Science</p><p>Term: Fall 2013, Duration: 268 days</p><p>This module covers key concepts and practical applications in applied analytics & algorithms.</p>"))
