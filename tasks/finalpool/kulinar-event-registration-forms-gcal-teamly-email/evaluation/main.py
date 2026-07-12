"""Evaluation для kulinar-event-registration-forms-gcal-teamly-email (RU-стек: kulinar/forms/teamly).

Проверяет четыре поверхности инфраструктуры мероприятия:
  1. Forms (gform.*): регистрационная форма с >=5 вопросами, покрывающими
     ФИО / email / отдел / пищевые ограничения / выбор блюда.
  2. Google Calendar (gcal.events): событие 2026-03-20, старт ~12:00, длительность
     ~1.5 часа, место — кухня.
  3. Teamly (teamly.pages): страница планирования с блюдами, ингредиентами/шагами
     и логистикой (дата + кухня).
  4. Email (email.messages): письмо от hr@ на all-staff@company.com с темой и телом,
     где есть дата/время/место.
  5. Cross-surface: ОДНИ И ТЕ ЖЕ 2 названия блюд из формы фигурируют в описании
     события, на странице Teamly и в теле письма.

Все текстовые проверки расширены на RU+EN. RU-ключевые слова ищутся в .lower()
ОРИГИНАЛЬНОМ тексте (НЕ через normalize, который схлопывает кириллицу в латиницу).

Критические чеки (CRITICAL_CHECKS): любой их fail => задача FAIL, даже если общая
accuracy >= 70%. Порог: accuracy >= 70% И отсутствие критических провалов.
"""

import os
import argparse
import json
import sys
import unicodedata

import psycopg2


DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Канонический набор блюд kulinar (источник: local_servers/kulinar-mcp/src/data/all_recipes.json).
KULINAR_RECIPES = {
    "Салат Оливье", "Винегрет", "Сельдь под шубой", "Салат Мимоза", "Крабовый салат",
    "Греческий салат", "Салат с курицей и грибами", "Холодец", "Икра кабачковая",
    "Грибы маринованные", "Сало солёное", "Селёдка с луком", "Борщ",
    "Щи из квашеной капусты", "Солянка мясная", "Уха", "Окрошка", "Грибной суп",
    "Рассольник", "Куриный бульон с лапшой", "Бефстроганов", "Пельмени домашние",
    "Голубцы", "Котлеты домашние", "Жаркое в горшочках", "Курица в сметане",
    "Рыба запечённая по-русски", "Цыплёнок табака", "Гречка с тушёнкой",
    "Плов узбекский", "Картофельное пюре", "Гречневая каша", "Перловая каша",
    "Картофель отварной с укропом", "Рис отварной", "Пирожки с капустой жареные",
    "Пирожки с мясом печёные", "Блины тонкие", "Кулебяка с капустой и яйцом",
    "Расстегаи с рыбой", "Медовик", "Наполеон", "Сырники", "Пасха творожная",
    "Ватрушки с творогом", "Кисель ягодный", "Морс клюквенный",
    "Компот из сухофруктов", "Сбитень", "Квас домашний",
}

# Критические чеки по имени record()
CRITICAL_CHECKS = {
    "Forms: регистрационная форма с >=5 вопросами (ФИО/email/отдел/ограничения/блюдо)",
    "Calendar: событие 2026-03-20, старт ~12:00, длительность ~1.5ч, место — кухня",
    "Email: письмо от hr@ на all-staff@company.com с темой и логистикой (дата/время/кухня)",
    "Teamly: страница планирования с блюдами, ингредиентами/шагами и логистикой",
    "Cross-surface: 2 блюда из формы согласованы в календаре, Teamly и письме",
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def record(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAILED
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED.append(name)


def norm_recipe(s: str) -> str:
    """Нормализация названия блюда для сравнения (регистр, ё/е, пробелы)."""
    s = (s or "").strip().lower().replace("ё", "е")
    return " ".join(s.split())


CANON_NORM = {norm_recipe(r) for r in KULINAR_RECIPES}


def _option_values(config):
    """Извлекает текстовые значения вариантов из config вопроса (RU forms-mcp).
    Форма config: {'type':'RADIO','options':[{'value': ...}, ...]}."""
    vals = []
    if not config:
        return vals
    if isinstance(config, str):
        try:
            config = json.loads(config)
        except Exception:
            return vals
    opts = config.get("options") if isinstance(config, dict) else None
    if isinstance(opts, list):
        for o in opts:
            if isinstance(o, dict):
                v = o.get("value")
                if v is not None:
                    vals.append(str(v))
            else:
                vals.append(str(o))
    return vals


# ---------------------------------------------------------------------------
# Forms (gform)
# ---------------------------------------------------------------------------
def check_google_form():
    """Проверка регистрационной формы. Возвращает список названий блюд (варианты
    последнего вопроса), которые сверяются на согласованность с остальными
    поверхностями."""
    print("\n=== Проверка Forms (регистрационная форма) ===")
    dish_options = []

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except Exception as e:
        record("Создана хотя бы одна форма", False, str(e))
        record("Forms: регистрационная форма с >=5 вопросами (ФИО/email/отдел/ограничения/блюдо)",
               False, str(e))
        return dish_options

    cur = conn.cursor()
    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    record("Создана хотя бы одна форма", len(forms) > 0, f"Найдено {len(forms)} форм")

    def title_matches(title):
        t = (title or "").lower()
        ru = any(k in t for k in ("обед", "готов", "кулинар", "регистрац"))
        en = any(k in t for k in ("lunch", "cooking", "registration"))
        return ru or en

    target = None
    for fid, title in forms:
        if title_matches(title):
            target = fid
            break
    if target is None and forms:
        # запасной вариант — форма с наибольшим числом вопросов
        best, best_q = None, -1
        for fid, _t in forms:
            cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (fid,))
            qc = cur.fetchone()[0]
            if qc > best_q:
                best, best_q = fid, qc
        target = best

    record("Найдена форма с подходящим заголовком (обед/готов/кулинар/lunch/cooking)",
           any(title_matches(t) for _f, t in forms),
           f"Заголовки: {[t for _f, t in forms]}")

    questions = []
    if target is not None:
        cur.execute("""
            SELECT id, title, question_type, required, config
            FROM gform.questions
            WHERE form_id = %s
            ORDER BY position
        """, (target,))
        questions = cur.fetchall()

    record("Форма содержит >=4 вопроса", len(questions) >= 4,
           f"Найдено {len(questions)} вопросов")

    q_titles_lower = [(q[1] or "").lower() for q in questions]

    has_name = any(("фио" in t or "имя" in t or "name" in t) for t in q_titles_lower)
    has_email = any(("почт" in t or "email" in t or "e-mail" in t) for t in q_titles_lower)
    has_dept = any(("отдел" in t or "department" in t) for t in q_titles_lower)
    has_diet = any(("пищев" in t or "ограничен" in t or "diet" in t or "restrict" in t)
                   for t in q_titles_lower)
    has_dish = any(("блюд" in t or "интересн" in t or "dish" in t or "excited" in t)
                   for t in q_titles_lower)

    record("Есть вопрос про ФИО/имя (фио/имя/name)", has_name, f"Вопросы: {[q[1] for q in questions]}")
    record("Есть вопрос про email (почт/email)", has_email, f"Вопросы: {[q[1] for q in questions]}")
    record("Есть вопрос про отдел (отдел/department)", has_dept, f"Вопросы: {[q[1] for q in questions]}")
    record("Есть вопрос про пищевые ограничения (пищев/ограничен/diet/restrict)",
           has_diet, f"Вопросы: {[q[1] for q in questions]}")
    record("Есть вопрос про выбор блюда (блюд/интересн/dish/excited)",
           has_dish, f"Вопросы: {[q[1] for q in questions]}")

    required_count = sum(1 for q in questions if q[3] is True)
    record("Не менее 3 обязательных вопросов", required_count >= 3,
           f"Обязательных: {required_count}")

    # Извлекаем варианты вопроса про блюдо для cross-surface чека
    for q in questions:
        qt = (q[1] or "").lower()
        if ("блюд" in qt or "интересн" in qt or "dish" in qt or "excited" in qt):
            dish_options = _option_values(q[4])
            break

    # CRITICAL: форма с >=5 вопросами, покрывающими все пять тем
    coverage = has_name and has_email and has_dept and has_diet and has_dish
    record("Forms: регистрационная форма с >=5 вопросами (ФИО/email/отдел/ограничения/блюдо)",
           target is not None and len(questions) >= 5 and coverage,
           f"вопросов={len(questions)}, покрытие={coverage}")

    cur.close()
    conn.close()

    # Оставляем только реальные блюда kulinar среди вариантов
    real_dishes = [d for d in dish_options
                   if norm_recipe(d) in CANON_NORM
                   or any(c in norm_recipe(d) or norm_recipe(d) in c for c in CANON_NORM if c)]
    record("Варианты вопроса про блюдо — реальные блюда kulinar (>=2)",
           len(real_dishes) >= 2,
           f"варианты={dish_options}")
    return real_dishes if len(real_dishes) >= 2 else dish_options


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------
def check_calendar(dish_names):
    print("\n=== Проверка Google Calendar ===")
    crit = "Calendar: событие 2026-03-20, старт ~12:00, длительность ~1.5ч, место — кухня"

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except Exception as e:
        record("Событие календаря (готов/обед/кулинар/lunch/cooking) существует", False, str(e))
        record(crit, False, str(e))
        return

    cur = conn.cursor()
    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime, location
        FROM gcal.events
        WHERE LOWER(summary) LIKE '%%cooking%%'
           OR LOWER(summary) LIKE '%%lunch%%'
           OR LOWER(summary) LIKE '%%готов%%'
           OR LOWER(summary) LIKE '%%обед%%'
           OR LOWER(summary) LIKE '%%кулинар%%'
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    record("Событие календаря (готов/обед/кулинар/lunch/cooking) существует",
           len(events) > 0, "Подходящих событий не найдено")

    if not events:
        record(crit, False, "нет события")
        return

    summary, description, start_dt, end_dt, location = events[0]
    print(f"  Событие: '{summary}'")

    date_ok = False
    hour_ok = False
    if start_dt:
        date_str = start_dt.strftime("%Y-%m-%d")
        date_ok = (date_str == "2026-03-20")
        record("Дата события — 2026-03-20", date_ok, f"Получено {date_str}")
        hour_ok = 8 <= start_dt.hour <= 18  # широкий допуск по таймзоне
        record("Старт около полудня (8-18 для допуска TZ)", hour_ok, f"Час {start_dt.hour}")
    else:
        record("У события есть start_datetime", False, "start_datetime пуст")

    dur_ok = False
    if start_dt and end_dt:
        dur = (end_dt - start_dt).total_seconds() / 3600
        dur_ok = 1.0 <= dur <= 2.0
        record("Длительность ~1.5 часа", dur_ok, f"Получено {dur:.1f} ч")

    desc = (description or "")
    record("Описание события не пустое", len(desc.strip()) > 10,
           f"Длина описания: {len(desc.strip())}")

    loc = (location or "")
    loc_l = loc.lower()
    desc_l = desc.lower()
    loc_ok = ("кухн" in loc_l or "kitchen" in loc_l
              or "кухн" in desc_l or "kitchen" in desc_l)
    record("Место упоминает кухню (кухн/kitchen)", loc_ok,
           f"location='{loc}'")

    # CRITICAL: дата + час + длительность + кухня
    record(crit, date_ok and hour_ok and dur_ok and loc_ok,
           f"date={date_ok}, hour={hour_ok}, dur={dur_ok}, kitchen={loc_ok}")


# ---------------------------------------------------------------------------
# Teamly
# ---------------------------------------------------------------------------
def check_teamly(dish_names):
    print("\n=== Проверка Teamly (страница планирования) ===")
    crit = "Teamly: страница планирования с блюдами, ингредиентами/шагами и логистикой"

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except Exception as e:
        record("Страница планирования в Teamly существует", False, str(e))
        record(crit, False, str(e))
        return

    cur = conn.cursor()
    cur.execute("SELECT title, COALESCE(body, '') FROM teamly.pages")
    pages = cur.fetchall()
    cur.close()
    conn.close()

    def title_matches(title):
        t = (title or "").lower()
        ru = ("планиров" in t or "готов" in t or "обед" in t or "кулинар" in t)
        en = ("planning" in t or "cooking" in t or "event" in t or "lunch" in t)
        return ru or en

    candidates = [(t, b) for t, b in pages if title_matches(t)]
    record("Страница планирования в Teamly существует (заголовок подходит)",
           len(candidates) >= 1, f"Заголовки: {[t for t, _b in pages]}")

    body = "\n\n".join(b for _t, b in candidates)
    bl = body.lower()

    has_logistics = ("20 март" in bl or "2026-03-20" in bl or "march" in bl
                     or "12:00" in bl or "кухн" in bl or "kitchen" in bl)
    record("Страница упоминает логистику (дата/время/место)",
           bool(candidates) and has_logistics, "Нет даты/времени/места в содержимом")

    has_recipe = any(k in bl for k in ("ингредиент", "шаг", "рецепт", "блюд",
                                       "ingredient", "step", "recipe", "dish", "cook"))
    record("Страница упоминает контент рецептов (ингредиенты/шаги/блюдо)",
           bool(candidates) and has_recipe, "Нет контента рецептов")

    has_kitchen = ("кухн" in bl or "kitchen" in bl)
    has_date = ("20 март" in bl or "2026-03-20" in bl or "march 20" in bl)

    # Оба названия блюд присутствуют
    nb = norm_recipe(body)
    dishes_present = sum(1 for d in dish_names if norm_recipe(d) and norm_recipe(d) in nb)

    # CRITICAL: страница есть, логистика (дата+кухня), контент рецептов, оба блюда
    record(crit,
           bool(candidates) and has_date and has_kitchen and has_recipe
           and dishes_present >= 2,
           f"cand={len(candidates)}, date={has_date}, kitchen={has_kitchen}, "
           f"recipe={has_recipe}, dishes={dishes_present}/{len(dish_names)}")


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------
def check_emails(dish_names):
    print("\n=== Проверка Email ===")
    crit = "Email: письмо от hr@ на all-staff@company.com с темой и логистикой (дата/время/кухня)"

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except Exception as e:
        record("Письмо с темой про обед/готовку отправлено", False, str(e))
        record(crit, False, str(e))
        return [], ""

    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE LOWER(subject) LIKE '%%lunch%%'
           OR LOWER(subject) LIKE '%%cooking%%'
           OR LOWER(subject) LIKE '%%обед%%'
           OR LOWER(subject) LIKE '%%готов%%'
           OR LOWER(subject) LIKE '%%кулинар%%'
    """)
    emails = cur.fetchall()
    cur.close()
    conn.close()

    record("Письмо с темой про обед/готовку (обед/готов/кулинар/lunch/cooking) отправлено",
           len(emails) >= 1, f"Найдено {len(emails)} писем")

    if not emails:
        record(crit, False, "нет письма")
        return

    subject, from_addr, to_addr, body_text = emails[0]
    print(f"  Тема: '{subject}'")

    subject_l = (subject or "").lower()
    body_l = (body_text or "").lower()

    has_date = ("20 март" in body_l or "2026-03-20" in body_l
                or "march 20" in body_l or "3/20" in body_l)
    record("Тело письма упоминает дату (20 марта / 2026-03-20)", has_date,
           "Дата не найдена в теле")

    has_time = any(k in body_l for k in ("12:00", "13:30", "1:30", "12 pm", "12 ч"))
    record("Тело письма упоминает время (12:00/13:30)", has_time,
           "Время не найдено в теле")

    has_loc = ("кухн" in body_l or "kitchen" in body_l)
    record("Тело письма упоминает место (Кухня компании / kitchen)", has_loc,
           "Место не найдено в теле")

    from_str = str(from_addr or "").lower()
    from_ok = "hr@" in from_str or from_str.startswith("hr")
    record("Письмо от hr@company.com", from_ok, f"From: '{from_addr}'")

    to_str = str(to_addr or "").lower()
    to_ok = "all-staff@company.com" in to_str or "all-staff" in to_str or "all_staff" in to_str
    record("Письмо на all-staff@company.com", to_ok, f"To: '{to_addr}'")

    # CRITICAL: hr -> all-staff, тема подходит, логистика в теле
    record(crit,
           from_ok and to_ok and has_date and has_time and has_loc,
           f"from={from_ok}, to={to_ok}, date={has_date}, time={has_time}, kitchen={has_loc}")


# ---------------------------------------------------------------------------
# Cross-surface dish consistency
# ---------------------------------------------------------------------------
def check_cross_surface(dish_names):
    print("\n=== Проверка согласованности блюд между поверхностями ===")
    crit = "Cross-surface: 2 блюда из формы согласованы в календаре, Teamly и письме"

    if not dish_names or len(dish_names) < 2:
        record(crit, False, f"в форме не нашлось 2 блюда: {dish_names}")
        return

    names = [norm_recipe(d) for d in dish_names[:2]]

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except Exception as e:
        record(crit, False, str(e))
        return
    cur = conn.cursor()

    # Описания событий календаря
    cur.execute("SELECT COALESCE(description,'') FROM gcal.events")
    cal_text = norm_recipe(" ".join(r[0] for r in cur.fetchall()))
    # Страницы Teamly
    cur.execute("SELECT COALESCE(title,'') || ' ' || COALESCE(body,'') FROM teamly.pages")
    teamly_text = norm_recipe(" ".join(r[0] for r in cur.fetchall()))
    # Тела писем
    cur.execute("SELECT COALESCE(body_text,'') FROM email.messages")
    email_text = norm_recipe(" ".join(r[0] for r in cur.fetchall()))

    cur.close()
    conn.close()

    in_cal = all(n in cal_text for n in names)
    in_teamly = all(n in teamly_text for n in names)
    in_email = all(n in email_text for n in names)

    record("Оба блюда присутствуют в описании события календаря", in_cal,
           f"блюда={names}")
    record("Оба блюда присутствуют на странице Teamly", in_teamly,
           f"блюда={names}")
    record("Оба блюда присутствуют в теле письма", in_email,
           f"блюда={names}")

    record(crit, in_cal and in_teamly and in_email,
           f"cal={in_cal}, teamly={in_teamly}, email={in_email}, блюда={names}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("KULINAR EVENT REGISTRATION (RU: kulinar/forms/teamly) - EVALUATION")
    print("=" * 70)

    dish_names = check_google_form()
    check_calendar(dish_names)
    check_teamly(dish_names)
    check_emails(dish_names)
    check_cross_surface(dish_names)

    total = PASS_COUNT + FAIL_COUNT
    pct = 100.0 * PASS_COUNT / total if total else 0.0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {pct:.1f}%")

    if CRITICAL_FAILED:
        print(f"  CRITICAL FAIL: {CRITICAL_FAILED}")

    success = (pct >= 70.0) and not CRITICAL_FAILED
    print(f"  Overall: {'PASS' if success else 'FAIL'}")

    if args.res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "accuracy": round(pct, 1),
            "critical_failed": CRITICAL_FAILED,
            "success": success,
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)

    if CRITICAL_FAILED:
        sys.exit(1)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
