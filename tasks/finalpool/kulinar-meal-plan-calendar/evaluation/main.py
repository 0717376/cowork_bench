"""
Evaluation script for kulinar-meal-plan-calendar task.

Checks via psycopg2:
1. Google Sheet: spreadsheet (title ~ "обед"/"меню"/"lunch"/"menu"),
   2 sheets ("Меню недели" with 5+ data rows, "Список покупок" with data)
2. Google Calendar: 5 events on March 9-13 2026, each noon-1pm, titled by dish
3. Email: sent to team@company.com summarizing the menu

CRITICAL semantic checks (gate the result; any failure => exit 1 BEFORE accuracy):
  C1. Menu dishes span >=4 DISTINCT kulinar categories.
  C2. Menu covers all 5 weekdays (RU/EN), each with a dish + a valid kulinar category.
  C3. >=5 calendar events (one per day Mar 9-13), each ~noon, title matches a menu
      dish name, description non-trivial.
  C4. Notification email to team@company.com lists all 5 weekdays with dishes from menu.
  C5. Shopping List has >=3 ingredient rows whose names belong to the kulinar
      ingredient vocabulary (content consistency, not just row count).

Accuracy threshold for the non-critical part: >= 70%.

Usage:
    python -m evaluation.main \
        --agent_workspace /path/to/workspace \
        --groundtruth_workspace /path/to/groundtruth \
        --launch_time "2026-03-06 10:00:00"
"""

import os
import re
import argparse
import json
import sys

import psycopg2
from zoneinfo import ZoneInfo

NY_TZ = ZoneInfo("America/New_York")

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# Valid kulinar recipe categories (источник правды — kulinar all_recipes.json).
KULINAR_CATEGORIES = {
    "салат", "закуска", "суп", "горячее", "гарнир", "выпечка", "десерт", "напиток",
}

# Weekday names as (EN, RU) pairs — agent legitimately writes Russian.
WEEKDAYS = [
    ("monday", "понедельник"),
    ("tuesday", "вторник"),
    ("wednesday", "среда"),
    ("thursday", "четверг"),
    ("friday", "пятница"),
]

# Checks whose failure forces an overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "CRITICAL: меню охватывает >=4 разных категории kulinar",
    "CRITICAL: лист 'Меню недели' покрывает 5 будних дней с блюдом и категорией",
    "CRITICAL: >=5 событий календаря (по дню), полдень, заголовок = блюдо из меню",
    "CRITICAL: письмо на team@company.com перечисляет 5 дней с блюдами из меню",
    "CRITICAL: 'Список покупок' содержит >=3 ингредиента из словаря kulinar",
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILS = []


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILS.append(name)


# ---------------------------------------------------------------------------
# Загрузка эталонной базы kulinar (для проверки словаря ингредиентов)
# ---------------------------------------------------------------------------

def _norm(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def load_kulinar_recipes():
    """Возвращает список рецептов kulinar или None, если база недоступна."""
    candidates = []
    env = os.environ.get("KULINAR_RECIPES_JSON")
    if env:
        candidates.append(env)
    here = os.path.abspath(__file__)
    cur = os.path.dirname(here)
    rel = os.path.join("local_servers", "kulinar-mcp", "src", "data", "all_recipes.json")
    for _ in range(12):
        candidates.append(os.path.join(cur, rel))
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    for path in candidates:
        try:
            if path and os.path.exists(path):
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list) and data:
                    return data
        except Exception:
            continue
    return None


KULINAR = load_kulinar_recipes()
# Словарь токенов названий ингредиентов из базы kulinar (по полю name).
KULINAR_ING_TOKENS = set()
if KULINAR:
    for r in KULINAR:
        for ing in r.get("ingredients", []):
            nm = _norm(ing.get("name"))
            for tok in re.split(r"[ ,()/]+", nm):
                if len(tok) >= 3:
                    KULINAR_ING_TOKENS.add(tok)


# =========================================================================
# Check 1: Google Sheet  (also extracts menu dish names + categories)
# =========================================================================

# Filled by check_gsheet() so calendar/email checks can cross-reference.
MENU_DISH_NAMES = []          # normalized dish names from "Меню недели"
MENU_CATEGORIES = []          # category cell values from "Меню недели"


def check_gsheet():
    """Verify spreadsheet with weekly menu and shopping list."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gsheet.spreadsheets")
    spreadsheets = cur.fetchall()
    print(f"[check_gsheet] Found {len(spreadsheets)} spreadsheets.")

    # Title may be Russian ("План обедов...") or English ("...Lunch...").
    title_subs = ("обед", "меню", "недел", "lunch", "menu", "plan", "план")
    target_ss = None
    for ss_id, title in spreadsheets:
        tl = (title or "").lower()
        if any(s in tl for s in title_subs):
            target_ss = (ss_id, title)
            break

    if not target_ss:
        cur.close()
        conn.close()
        record("gsheet: spreadsheet found", False,
               f"No matching spreadsheet. Found: {[t for _, t in spreadsheets]}")
        return False, "No spreadsheet"

    ss_id, ss_title = target_ss
    record("gsheet: spreadsheet found", True)
    print(f"[check_gsheet] Using spreadsheet: {ss_title} ({ss_id})")

    cur.execute(
        "SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s ORDER BY index",
        (str(ss_id),)
    )
    sheets = cur.fetchall()
    sheet_titles = [t for _, t in sheets]
    print(f"[check_gsheet] Sheets: {sheet_titles}")

    record("gsheet: has at least 2 sheets", len(sheets) >= 2,
           f"Found {len(sheets)} sheets: {sheet_titles}")

    menu_subs = ("меню", "menu", "недел", "weekly")
    shop_subs = ("покуп", "список", "ингредиент", "shopping", "list", "ingredient")
    menu_sheet = None
    shopping_sheet = None
    for sheet_id, sheet_title in sheets:
        tl = (sheet_title or "").lower()
        if menu_sheet is None and any(s in tl for s in menu_subs):
            menu_sheet = (sheet_id, sheet_title)
        elif any(s in tl for s in shop_subs):
            shopping_sheet = (sheet_id, sheet_title)

    record("gsheet: 'Меню недели' sheet found", menu_sheet is not None,
           f"No menu-like sheet. Sheets: {sheet_titles}")
    record("gsheet: 'Список покупок' sheet found", shopping_sheet is not None,
           f"No shopping-list-like sheet. Sheets: {sheet_titles}")

    days_found = 0
    menu_categories = []
    rows_in_menu = []
    if menu_sheet:
        menu_sheet_id = menu_sheet[0]
        cur.execute(
            "SELECT DISTINCT row_index FROM gsheet.cells WHERE sheet_id = %s ORDER BY row_index",
            (menu_sheet_id,)
        )
        rows = [r[0] for r in cur.fetchall()]
        min_row = min(rows) if rows else 0
        data_rows = [r for r in rows if r > min_row]
        record("gsheet: Меню недели has 5 data rows", len(data_rows) >= 5,
               f"Found {len(data_rows)} data rows (row indices: {rows})")

        cur.execute(
            "SELECT row_index, col_index, value FROM gsheet.cells "
            "WHERE sheet_id = %s ORDER BY row_index, col_index",
            (menu_sheet_id,)
        )
        all_cells = cur.fetchall()

        # Reconstruct rows; row[0] (min) is header. Locate Day/Dish/Category columns.
        by_row = {}
        for ri, ci, v in all_cells:
            by_row.setdefault(ri, {})[ci] = v or ""
        header = by_row.get(min_row, {})

        def find_col(keys):
            for ci, hv in header.items():
                hl = _norm(hv)
                if any(k in hl for k in keys):
                    return ci
            return None

        col_day = find_col(("день", "day"))
        col_dish = find_col(("блюдо", "dish", "название"))
        col_cat = find_col(("категор", "categor"))

        data_values = " ".join(_norm(v) for ri, r in by_row.items()
                               if ri > min_row for v in r.values())
        for en, ru in WEEKDAYS:
            if en in data_values or ru in data_values:
                days_found += 1
        record("gsheet: Меню недели mentions all 5 weekdays", days_found >= 5,
               f"Found {days_found}/5 weekday names in menu data")

        for ri in sorted(by_row):
            if ri == min_row:
                continue
            row = by_row[ri]
            dish = _norm(row.get(col_dish, "")) if col_dish is not None else ""
            cat = _norm(row.get(col_cat, "")) if col_cat is not None else ""
            # Fallback: scan whole row for a known category if no Category column.
            if not cat:
                for v in row.values():
                    if _norm(v) in KULINAR_CATEGORIES:
                        cat = _norm(v)
                        break
            if not dish:
                # Fallback: first non-day, non-category cell as dish name.
                for ci in sorted(row):
                    vv = _norm(row[ci])
                    if not vv or vv in KULINAR_CATEGORIES:
                        continue
                    if any(vv == ru or vv == en for en, ru in WEEKDAYS):
                        continue
                    dish = vv
                    break
            if dish:
                MENU_DISH_NAMES.append(dish)
            if cat:
                menu_categories.append(cat)
            rows_in_menu.append((dish, cat))

        MENU_CATEGORIES.extend(menu_categories)

    # ---- CRITICAL C1: category diversity (>=4 distinct kulinar categories) ----
    distinct_cats = {c for c in MENU_CATEGORIES if c in KULINAR_CATEGORIES}
    record("CRITICAL: меню охватывает >=4 разных категории kulinar",
           len(distinct_cats) >= 4,
           f"Распознано категорий kulinar: {sorted(distinct_cats)} "
           f"(значения категорий: {MENU_CATEGORIES})")

    # ---- CRITICAL C2: 5 weekdays each with dish + valid category ----
    valid_rows = [
        (d, c) for (d, c) in rows_in_menu if d and c in KULINAR_CATEGORIES
    ]
    record("CRITICAL: лист 'Меню недели' покрывает 5 будних дней с блюдом и категорией",
           days_found >= 5 and len(valid_rows) >= 5,
           f"weekdays={days_found}/5, строк с блюдом+категорией={len(valid_rows)}")

    # ---- Shopping List structural + CRITICAL C5 content ----
    shopping_data_rows = []
    shop_ing_values = []
    if shopping_sheet:
        shopping_sheet_id = shopping_sheet[0]
        cur.execute(
            "SELECT row_index, col_index, value FROM gsheet.cells "
            "WHERE sheet_id = %s ORDER BY row_index, col_index",
            (shopping_sheet_id,)
        )
        shop_cells = cur.fetchall()
        shop_rows = {}
        for ri, ci, v in shop_cells:
            shop_rows.setdefault(ri, {})[ci] = v or ""
        if shop_rows:
            smin = min(shop_rows)
            shopping_data_rows = [ri for ri in shop_rows if ri > smin]
            # Ingredient column = first column by default.
            sheader = shop_rows.get(smin, {})
            ing_col = None
            for ci, hv in sheader.items():
                if any(k in _norm(hv) for k in ("ингредиент", "ingredient")):
                    ing_col = ci
                    break
            if ing_col is None:
                ing_col = min(sheader) if sheader else 0
            for ri in shopping_data_rows:
                shop_ing_values.append(_norm(shop_rows[ri].get(ing_col, "")))

        record("gsheet: Список покупок has ingredient rows", len(shopping_data_rows) >= 3,
               f"Found {len(shopping_data_rows)} data rows in Shopping List")

    # CRITICAL C5: ingredient content consistency against kulinar vocabulary.
    if KULINAR_ING_TOKENS:
        matched = 0
        for val in shop_ing_values:
            toks = [t for t in re.split(r"[ ,()/]+", val) if len(t) >= 3]
            if any(t in KULINAR_ING_TOKENS for t in toks):
                matched += 1
        record("CRITICAL: 'Список покупок' содержит >=3 ингредиента из словаря kulinar",
               matched >= 3,
               f"Совпало ингредиентов с базой kulinar: {matched} из {len(shop_ing_values)}")
    else:
        # База недоступна в окружении — не блокируем (структурная проверка остаётся).
        record("CRITICAL: 'Список покупок' содержит >=3 ингредиента из словаря kulinar",
               len(shopping_data_rows) >= 3,
               "база kulinar недоступна — проверка по числу строк")

    cur.close()
    conn.close()

    all_ok = (
        len(sheets) >= 2
        and menu_sheet is not None
        and shopping_sheet is not None
    )
    return all_ok, None if all_ok else "Some gsheet checks failed"


# =========================================================================
# Check 2: Google Calendar
# =========================================================================

def check_gcal():
    """Verify 5 calendar events for March 9-13, 2026 around noon."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime
        FROM gcal.events
        WHERE start_datetime >= '2026-03-09 00:00:00'
          AND start_datetime < '2026-03-14 00:00:00'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_gcal] Found {len(events)} calendar events in March 9-13.")
    for ev in events:
        print(f"  Event: {ev[0]} | {ev[2]} - {ev[3]}")

    record("gcal: at least 5 events in March 9-13", len(events) >= 5,
           f"Found {len(events)} events")

    days_covered = set()
    noon_events = 0
    events_with_description = 0
    title_matches = 0

    for summary, description, start_dt, end_dt in events:
        if start_dt:
            days_covered.add(start_dt.strftime("%Y-%m-%d"))
            # start_datetime is timestamptz (stored UTC). Task requires 12:00 in
            # America/New_York, which is 16:00 UTC during EDT. Convert to NY before
            # reading the hour so a timezone-correct submission passes the noon gate.
            if start_dt.tzinfo is not None:
                local_dt = start_dt.astimezone(NY_TZ)
            else:
                local_dt = start_dt.replace(tzinfo=ZoneInfo("UTC")).astimezone(NY_TZ)
            if 11 <= local_dt.hour <= 13:
                noon_events += 1
        if description and len(description.strip()) > 5:
            events_with_description += 1
        s = _norm(summary)
        if MENU_DISH_NAMES and any(d and (d in s or s in d) for d in MENU_DISH_NAMES):
            title_matches += 1

    expected_days = {"2026-03-09", "2026-03-10", "2026-03-11", "2026-03-12", "2026-03-13"}
    missing_days = expected_days - days_covered
    record("gcal: events cover all 5 weekdays", len(missing_days) == 0,
           f"Missing days: {missing_days}")
    record("gcal: events start around noon", noon_events >= 5,
           f"Only {noon_events}/5 events start between 11:00-13:00")
    record("gcal: events have descriptions with ingredients",
           events_with_description >= 5,
           f"Only {events_with_description}/5 events have non-trivial descriptions")

    # CRITICAL C3: 5 per-day noon events whose titles match menu dishes + descriptions.
    # If no menu dishes were parsed, fall back to requiring non-trivial titles.
    if MENU_DISH_NAMES:
        title_ok = title_matches >= 5
        title_detail = f"Заголовков, совпавших с блюдами меню: {title_matches}/5"
    else:
        title_ok = len(events) >= 5
        title_detail = "меню не распознано — проверка по числу событий"
    record("CRITICAL: >=5 событий календаря (по дню), полдень, заголовок = блюдо из меню",
           len(events) >= 5 and len(missing_days) == 0 and noon_events >= 5
           and events_with_description >= 5 and title_ok,
           f"events={len(events)}, missing={missing_days}, noon={noon_events}, "
           f"desc={events_with_description}, {title_detail}")

    all_ok = (
        len(events) >= 5
        and len(missing_days) == 0
        and noon_events >= 5
        and events_with_description >= 5
    )
    return all_ok, None if all_ok else "Some gcal checks failed"


# =========================================================================
# Check 3: Email
# =========================================================================

def check_email():
    """Verify notification email sent to team@company.com."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_email] Found {len(all_emails)} total emails.")

    # Subject may be Russian ("План обедов: 9-13 марта 2026") or English.
    def subj_match(subject):
        sl = (subject or "").lower()
        has_topic = any(t in sl for t in ("обед", "lunch", "меню", "menu", "план"))
        has_month = ("март" in sl) or ("march" in sl)
        return has_topic and has_month

    target_email = None
    for subject, from_addr, to_addr, body_text in all_emails:
        if subj_match(subject):
            target_email = (subject, from_addr, to_addr, body_text)
            break
    if not target_email:
        for subject, from_addr, to_addr, body_text in all_emails:
            if "team@company.com" in _to_addr_str(to_addr):
                target_email = (subject, from_addr, to_addr, body_text)
                break

    record("email: notification email found", target_email is not None,
           f"No matching email. Found subjects: {[s for s, _, _, _ in all_emails]}")

    if not target_email:
        return False, "Notification email not found"

    subject, from_addr, to_addr, body_text = target_email
    print(f"[check_email] Found email: {subject}")

    to_str = _to_addr_str(to_addr)
    record("email: sent to team@company.com", "team@company.com" in to_str,
           f"Recipient: {to_addr}")

    sl = (subject or "").lower()
    record("email: subject mentions march/март", ("march" in sl or "март" in sl),
           f"Subject: {subject}")

    body_lower = (body_text or "").lower()
    days_in_body = 0
    for en, ru in WEEKDAYS:
        if en in body_lower or ru in body_lower:
            days_in_body += 1
    record("email: body mentions weekdays", days_in_body >= 3,
           f"Found {days_in_body}/5 weekday names in body")
    record("email: body has substantial content", len(body_text or "") > 50,
           f"Body length: {len(body_text or '')} chars")

    # CRITICAL C4: lists all 5 weekdays AND dishes from the menu.
    dishes_in_body = 0
    if MENU_DISH_NAMES:
        for d in MENU_DISH_NAMES:
            if d and d in body_lower:
                dishes_in_body += 1
        dishes_ok = dishes_in_body >= 5
        dish_detail = f"блюд меню в теле: {dishes_in_body}/5"
    else:
        dishes_ok = len(body_text or "") > 50
        dish_detail = "меню не распознано — проверка по объёму тела"
    record("CRITICAL: письмо на team@company.com перечисляет 5 дней с блюдами из меню",
           ("team@company.com" in to_str) and days_in_body >= 5 and dishes_ok,
           f"recipient_ok={'team@company.com' in to_str}, "
           f"weekdays={days_in_body}/5, {dish_detail}")

    all_ok = (
        target_email is not None
        and "team@company.com" in to_str
        and days_in_body >= 3
    )
    return all_ok, None if all_ok else "Some email checks failed"


def _to_addr_str(to_addr):
    """Convert to_addr (JSONB or string) to a lowercase search string."""
    if isinstance(to_addr, list):
        return " ".join(str(r).lower() for r in to_addr)
    elif isinstance(to_addr, str):
        try:
            parsed = json.loads(to_addr)
            if isinstance(parsed, list):
                return " ".join(str(r).lower() for r in parsed)
            return str(to_addr).lower()
        except (json.JSONDecodeError, TypeError):
            return str(to_addr).lower()
    return str(to_addr or "").lower()


# =========================================================================
# Main
# =========================================================================

def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    """Run all evaluation checks."""

    print("\n=== Checking Google Sheet ===")
    gsheet_pass, _ = check_gsheet()

    print("\n=== Checking Google Calendar ===")
    gcal_pass, _ = check_gcal()

    print("\n=== Checking Email ===")
    email_pass, _ = check_email()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    accuracy_ok = accuracy >= 70.0
    critical_ok = len(CRITICAL_FAILS) == 0
    all_passed = critical_ok and accuracy_ok

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}% (threshold 70%)")
    if CRITICAL_FAILS:
        print(f"  CRITICAL FAILURES: {CRITICAL_FAILS}")
    print(f"  Overall: {'PASS' if all_passed else 'FAIL'}")

    if res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "accuracy": accuracy,
            "critical_failures": CRITICAL_FAILS,
            "success": all_passed,
            "details": {
                "gsheet": gsheet_pass,
                "gcal": gcal_pass,
                "email": email_pass,
            },
        }
        with open(res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    return all_passed, f"Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}, Accuracy: {accuracy:.1f}%"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    success, message = run_evaluation(
        args.agent_workspace,
        args.groundtruth_workspace,
        args.launch_time,
        args.res_log_file,
    )
    print(message)

    # Critical failures force FAIL regardless of accuracy.
    if CRITICAL_FAILS:
        print(f"[CRITICAL] Failing due to: {CRITICAL_FAILS}")
        sys.exit(1)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
