"""
Evaluation for kulinar-weekly-gsheet-gcal task (RU stack: kulinar).

Проверяет полный недельный план питания, построенный агентом:
  1. Google-таблица "Weekly Meal Plan" (точное совпадение названия) с ровно 21
     строкой данных + заголовок; столбцы Day/Meal_Type/Recipe_Name/Prep_Time/
     Difficulty; каждый день 1–7 встречается 3 раза; Meal_Type ∈ {Breakfast,
     Lunch, Dinner}; Recipe_Name непустой; Prep_Time числовой; Difficulty ∈
     {Easy, Medium, Hard}.
  2. >=15 из 21 названий Recipe_Name — реальные блюда из базы kulinar (агент
     действительно пользовался базой, а не выдумал блюда).
  3. Согласованность Difficulty/Prep_Time: Easy↔15, Medium↔30, Hard↔60.
  4. Ровно 7 событий "Dinner Prep - Day 1..7" на 2026-04-07..04-13, каждое
     18:00-19:00, и в описании — то же блюдо-ужин, что в строке Dinner таблицы.
  5. Письмо на meal_planning@service.com с темой "Weekly Meal Plan Ready",
     тело упоминает несколько реальных блюд и ссылается на таблицу и календарь.

Критические чеки (CRITICAL_CHECKS): любой их fail => задача FAIL, даже если
общая accuracy >= 70%. Структурные чеки — мягкие. Порог: accuracy >= 70% И нет
критических провалов.
"""
import json
import os
import re
import sys
import unicodedata
from argparse import ArgumentParser

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Канонический набор блюд kulinar (источник: local_servers/kulinar-mcp -> all_recipes.json).
# Это ПОЛНАЯ база из 50 блюд, а не подсказка под конкретный инстанс: чек проверяет,
# что названия из таблицы принадлежат реальной базе kulinar.
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

# Соответствие Difficulty <-> Prep_Time (правило из task.md).
DIFF_TO_PREP = {"easy": 15, "medium": 30, "hard": 60}

CRITICAL_CHECKS = {
    "GSheet 'Weekly Meal Plan' — 21 строка, корректные столбцы и значения",
    "Recipe_Name: >=15 из 21 — реальные блюда из базы kulinar",
    "Difficulty/Prep_Time согласованы (Easy↔15, Medium↔30, Hard↔60)",
    "7 событий 'Dinner Prep - Day 1..7' 18:00-19:00 с блюдом-ужином в описании",
    "Письмо на meal_planning@service.com с темой и содержанием плана",
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAILED
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED.append(name)


def norm_recipe(s):
    """Нормализация названия блюда (регистр, ё/е, пробелы) для сравнения с каноном."""
    s = (s or "").strip().lower().replace("ё", "е")
    return " ".join(s.split())


CANON_NORM = {norm_recipe(r) for r in KULINAR_RECIPES}


def load_sheet_rows():
    """Возвращает (title, header_map, rows) для таблицы 'Weekly Meal Plan'.

    header_map: {column_name_lower: col_index}.
    rows: list[dict] по строкам данных (row_index>0): {col_name_lower: value}.
    Точное совпадение названия — без fallback на произвольную таблицу.
    """
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE title = 'Weekly Meal Plan'
        ORDER BY created_at DESC LIMIT 1
    """)
    sp = cur.fetchone()
    if not sp:
        cur.close()
        conn.close()
        return None, {}, []
    sp_id, title = sp

    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s
    """, (sp_id,))
    cells = cur.fetchall()
    cur.close()
    conn.close()

    header_map = {}
    row_cells = {}
    for r_idx, c_idx, value in cells:
        if r_idx == 0:
            if value:
                header_map[str(value).strip().lower()] = c_idx
        else:
            row_cells.setdefault(r_idx, {})[c_idx] = value

    # Развернём по индексу столбца через header_map
    idx_to_name = {v: k for k, v in header_map.items()}
    rows = []
    for r_idx in sorted(row_cells):
        rowmap = {}
        nonempty = False
        for c_idx, value in row_cells[r_idx].items():
            cname = idx_to_name.get(c_idx)
            if cname is not None:
                rowmap[cname] = value
            if value not in (None, ""):
                nonempty = True
        if nonempty:
            rows.append(rowmap)
    return title, header_map, rows


def check_gsheet(rows, header_map):
    print("\n=== Проверка 1: Google-таблица 'Weekly Meal Plan' ===")

    required_cols = {"day", "meal_type", "recipe_name", "prep_time", "difficulty"}
    cols_ok = required_cols.issubset(set(header_map.keys()))
    record("Таблица имеет столбцы Day/Meal_Type/Recipe_Name/Prep_Time/Difficulty",
           cols_ok, f"Найдены столбцы: {sorted(header_map.keys())}")

    # Сводный критический чек структуры/значений
    problems = []
    if not cols_ok:
        problems.append("отсутствуют обязательные столбцы")
    if len(rows) != 21:
        problems.append(f"строк данных {len(rows)} != 21")

    day_counts = {}
    valid_meal = {"breakfast", "lunch", "dinner"}
    valid_diff = {"easy", "medium", "hard"}
    if cols_ok:
        for row in rows:
            day = str(row.get("day", "")).strip()
            meal = str(row.get("meal_type", "")).strip().lower()
            recipe = str(row.get("recipe_name", "")).strip()
            prep = str(row.get("prep_time", "")).strip()
            diff = str(row.get("difficulty", "")).strip().lower()
            try:
                d = int(float(day))
                day_counts[d] = day_counts.get(d, 0) + 1
            except (TypeError, ValueError):
                problems.append(f"Day не число: {day!r}")
            if meal not in valid_meal:
                problems.append(f"Meal_Type вне набора: {meal!r}")
            if not recipe:
                problems.append("пустой Recipe_Name")
            if not re.fullmatch(r"\d+(\.\d+)?", prep):
                problems.append(f"Prep_Time не число: {prep!r}")
            if diff not in valid_diff:
                problems.append(f"Difficulty вне набора: {diff!r}")
        for d in range(1, 8):
            if day_counts.get(d, 0) != 3:
                problems.append(f"День {d} встречается {day_counts.get(d, 0)} раз (нужно 3)")

    record("GSheet 'Weekly Meal Plan' — 21 строка, корректные столбцы и значения",
           len(problems) == 0, "; ".join(problems[:6]))


def check_recipes(rows):
    print("\n=== Проверка 2: Recipe_Name соответствует базе kulinar ===")
    names = [str(r.get("recipe_name", "")).strip() for r in rows]
    names = [n for n in names if n]
    matched = sum(1 for n in names if norm_recipe(n) in CANON_NORM)
    record("Recipe_Name: >=15 из 21 — реальные блюда из базы kulinar",
           matched >= 15,
           f"Совпало {matched}/{len(names)} с базой kulinar")


def check_difficulty(rows):
    print("\n=== Проверка 3: согласованность Difficulty/Prep_Time ===")
    bad = []
    for row in rows:
        diff = str(row.get("difficulty", "")).strip().lower()
        prep = str(row.get("prep_time", "")).strip()
        expected = DIFF_TO_PREP.get(diff)
        try:
            prep_num = int(float(prep))
        except (TypeError, ValueError):
            prep_num = None
        if expected is None or prep_num != expected:
            bad.append(f"{row.get('recipe_name','?')}: {diff!r}/{prep!r}")
    record("Difficulty/Prep_Time согласованы (Easy↔15, Medium↔30, Hard↔60)",
           len(bad) == 0 and len(rows) > 0,
           f"Несогласованных строк: {len(bad)}; примеры: {bad[:5]}")


def get_dinner_recipes(rows):
    """{day:int -> recipe_name(str)} по строкам Dinner."""
    out = {}
    for row in rows:
        if str(row.get("meal_type", "")).strip().lower() == "dinner":
            try:
                d = int(float(str(row.get("day", "")).strip()))
            except (TypeError, ValueError):
                continue
            out[d] = str(row.get("recipe_name", "")).strip()
    return out


def check_gcal(dinner_by_day):
    print("\n=== Проверка 4: события Google Календаря 'Dinner Prep' ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime
        FROM gcal.events
        WHERE start_datetime >= '2026-04-07' AND start_datetime < '2026-04-14'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    # Сопоставим события по дню из заголовка "Dinner Prep - Day X"
    by_day = {}
    for summary, desc, start_dt, end_dt in events:
        m = re.search(r"dinner\s*prep\s*-\s*day\s*(\d)", (summary or "").lower())
        if m:
            by_day[int(m.group(1))] = (summary, desc or "", start_dt, end_dt)

    problems = []
    if len(by_day) != 7 or set(by_day.keys()) != set(range(1, 8)):
        problems.append(f"дни событий: {sorted(by_day.keys())} (нужно 1..7)")

    for day in range(1, 8):
        ev = by_day.get(day)
        if not ev:
            problems.append(f"нет события Day {day}")
            continue
        summary, desc, start_dt, end_dt = ev
        if not start_dt or start_dt.hour != 18:
            problems.append(f"Day {day} старт != 18:00 ({start_dt})")
        if not end_dt or end_dt.hour != 19:
            problems.append(f"Day {day} конец != 19:00 ({end_dt})")
        if start_dt and start_dt.date().isoformat() != f"2026-04-{6 + day:02d}":
            problems.append(f"Day {day} неверная дата ({start_dt.date()})")
        # Кросс-чек: блюдо-ужин дня должно быть в описании
        recipe = dinner_by_day.get(day, "")
        if recipe and norm_recipe(recipe) not in norm_recipe(desc):
            problems.append(f"Day {day}: блюдо '{recipe}' не упомянуто в описании")

    record("7 событий 'Dinner Prep - Day 1..7' 18:00-19:00 с блюдом-ужином в описании",
           len(problems) == 0, "; ".join(problems[:6]))


def check_email(rows):
    print("\n=== Проверка 5: письмо на meal_planning@service.com ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    messages = cur.fetchall()
    cur.close()
    conn.close()

    matching = None
    for subject, from_addr, to_addr, body_text in messages:
        to_str = ""
        if isinstance(to_addr, list):
            to_str = " ".join(str(r).lower() for r in to_addr)
        elif isinstance(to_addr, str):
            try:
                parsed = json.loads(to_addr)
                to_str = " ".join(str(r).lower() for r in parsed) if isinstance(parsed, list) else str(to_addr).lower()
            except Exception:
                to_str = str(to_addr).lower()
        if "meal_planning@service.com" in to_str:
            matching = (subject, body_text)
            break

    if not matching:
        record("Письмо на meal_planning@service.com с темой и содержанием плана",
               False, f"Нет письма на адрес; всего сообщений: {len(messages)}")
        return

    subject, body_text = matching
    subj = (subject or "").strip().lower()
    body = norm_recipe(body_text or "")

    problems = []
    if "weekly meal plan ready" not in subj:
        problems.append(f"тема != 'Weekly Meal Plan Ready' ({subject!r})")
    # тело упоминает несколько реальных блюд из таблицы
    sheet_names = [str(r.get("recipe_name", "")).strip() for r in rows]
    sheet_names = [n for n in sheet_names if n]
    mentioned = sum(1 for n in set(sheet_names) if norm_recipe(n) in body)
    if mentioned < 2:
        problems.append(f"в теле упомянуто блюд из таблицы: {mentioned} (нужно >=2)")
    # ссылается и на таблицу, и на календарь
    refs_sheet = any(k in body for k in ("таблиц", "google", "weekly meal plan"))
    refs_cal = any(k in body for k in ("календар", "событ", "calendar", "dinner prep"))
    if not refs_sheet:
        problems.append("нет упоминания таблицы")
    if not refs_cal:
        problems.append("нет упоминания календаря")

    record("Письмо на meal_planning@service.com с темой и содержанием плана",
           len(problems) == 0, "; ".join(problems[:5]))


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    title, header_map, rows = load_sheet_rows()
    if title is None:
        record("GSheet 'Weekly Meal Plan' — 21 строка, корректные столбцы и значения",
               False, "Таблица 'Weekly Meal Plan' не найдена")
        rows, header_map = [], {}
    else:
        check_gsheet(rows, header_map)

    check_recipes(rows)
    check_difficulty(rows)
    check_gcal(get_dinner_recipes(rows))
    check_email(rows)

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "critical_failed": CRITICAL_FAILED,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)

    if CRITICAL_FAILED:
        print(f"\nCRITICAL FAILURES ({len(CRITICAL_FAILED)}): {CRITICAL_FAILED}")
        print("FAIL")
        sys.exit(1)

    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
