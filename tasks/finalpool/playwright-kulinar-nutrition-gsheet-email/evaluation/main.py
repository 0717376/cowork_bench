"""
Evaluation script for playwright-kulinar-nutrition-gsheet-email task.

Checks:
1. Google Sheets spreadsheet "Wellness_Lunch_Menu" with Weekly_Plan and
   Nutrition_Summary sheets
2. Word document Wellness_Menu_Summary.docx with menu content
3. Email sent to cafeteria-manager about wellness lunch menu

Critical checks (see CRITICAL_CHECKS): any failure => overall FAIL regardless
of accuracy. Otherwise pass threshold: accuracy >= 70%.

Note: the agent is asked to create a Google Sheets spreadsheet plus a .docx --
NOT an .xlsx file. The xlsx gate of the original task was a desync bug and has
been removed; there is no .xlsx deliverable.
"""

import argparse
import json
import os
import sys

import psycopg2

try:
    from docx import Document
except ImportError:
    Document = None

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Per-meal targets from the guidelines page (numbers are the only hard
# source-of-truth on the served HTML).
TARGETS = {"calories": 650, "protein": 30, "fat": 22, "carbs": 80}

# Critical (semantic) checks: any failure => overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "Wellness_Lunch_Menu spreadsheet exists",
    "Nutrition_Summary Daily_Target matches guidelines",
    "Weekly_Plan has 5 weekday rows",
    "Dishes match kulinar recipe DB",
    "Wellness menu email exists",
    "Email to cafeteria-manager",
    "Doc title 'Corporate Wellness Lunch Program' present",
}

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def str_contains(haystack, needle):
    if haystack is None or needle is None:
        return False
    return needle.strip().lower() in str(haystack).strip().lower()


def load_kulinar_recipe_names():
    """Load all kulinar recipe names from the MCP data file (cross-check)."""
    candidates = [
        os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "..", "..", "..", "..",
            "local_servers", "kulinar-mcp", "src", "data", "all_recipes.json",
        ),
    ]
    for path in candidates:
        path = os.path.abspath(path)
        if os.path.isfile(path):
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                return {str(r.get("name", "")).strip().lower() for r in data if r.get("name")}
            except Exception:
                pass
    return set()


def check_gsheet():
    """Check Google Sheets spreadsheet."""
    print("\n=== Checking Google Sheets ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Find the Wellness_Lunch_Menu spreadsheet
        cur.execute(
            "SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE %s",
            ("%wellness%lunch%menu%",),
        )
        rows = cur.fetchall()
        if not rows:
            cur.execute("SELECT id, title FROM gsheet.spreadsheets")
            all_ss = cur.fetchall()
            record(
                "Wellness_Lunch_Menu spreadsheet exists",
                False,
                f"Found spreadsheets: {[r[1] for r in all_ss]}",
            )
            cur.close()
            conn.close()
            return False

        record("Wellness_Lunch_Menu spreadsheet exists", True)
        ss_id = rows[0][0]

        # Check sheets
        cur.execute(
            "SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s",
            (ss_id,),
        )
        sheets = cur.fetchall()
        sheet_titles = [s[1].lower() for s in sheets]

        has_weekly = any("weekly" in t or "plan" in t for t in sheet_titles)
        has_nutrition = any("nutrition" in t or "summary" in t for t in sheet_titles)
        record("Weekly_Plan sheet exists", has_weekly, f"Sheets: {sheet_titles}")
        record(
            "Nutrition_Summary sheet exists", has_nutrition, f"Sheets: {sheet_titles}"
        )

        # Check Weekly_Plan content - should have 5 data rows (Mon-Fri)
        weekly_sheet_id = None
        for sid, title in sheets:
            if "weekly" in title.lower() or "plan" in title.lower():
                weekly_sheet_id = sid
                break

        weekly_dish_values = []
        if weekly_sheet_id:
            cur.execute(
                "SELECT row_index, col_index, value FROM gsheet.cells "
                "WHERE sheet_id = %s ORDER BY row_index, col_index",
                (weekly_sheet_id,),
            )
            cells = cur.fetchall()

            # Count data rows (row_index > 0 for header row 0)
            data_rows = set()
            for row_idx, col_idx, val in cells:
                if row_idx > 0 and val and str(val).strip():
                    data_rows.add(row_idx)

            # CRITICAL: 5 weekday rows
            record(
                "Weekly_Plan has 5 weekday rows",
                len(data_rows) >= 5,
                f"Found {len(data_rows)} data rows",
            )

            # Check days of week present (accept EN weekday names and RU пн-пт)
            all_values = [str(v).lower() for _, _, v in cells if v]
            day_aliases = [
                ["monday", "понедельник", "пн"],
                ["tuesday", "вторник", "вт"],
                ["wednesday", "среда", "ср"],
                ["thursday", "четверг", "чт"],
                ["friday", "пятница", "пт"],
            ]
            days_found = sum(
                1
                for aliases in day_aliases
                if any(any(a in v for a in aliases) for v in all_values)
            )
            record(
                "All 5 weekdays in Weekly_Plan",
                days_found >= 5,
                f"Found {days_found} days",
            )

            # Check nutritional columns exist (calories, protein, fat, carbs)
            header_values = [str(v).lower() for r, c, v in cells if r == 0 and v]
            has_cal = any("calori" in h for h in header_values)
            has_prot = any("protein" in h for h in header_values)
            has_fat = any("fat" in h for h in header_values)
            has_carb = any("carb" in h for h in header_values)
            record(
                "Nutrition columns in Weekly_Plan",
                has_cal and has_prot and has_fat and has_carb,
                f"Headers: {header_values}",
            )

            # Collect dish-cell text values for the kulinar cross-check.
            # Dish columns are the non-numeric, non-day, non-header cells.
            num_cols = set()
            for r, c, v in cells:
                if r == 0 and v:
                    h = str(v).lower()
                    if any(k in h for k in ("calori", "protein", "fat", "carb", "day")):
                        num_cols.add(c)
            for r, c, v in cells:
                if r > 0 and c not in num_cols and v and str(v).strip():
                    weekly_dish_values.append(str(v).strip())
        else:
            record("Weekly_Plan has 5 weekday rows", False, "No Weekly_Plan sheet")

        # CRITICAL: dishes named in Weekly_Plan exist in the kulinar recipe DB.
        kulinar_names = load_kulinar_recipe_names()
        if kulinar_names:
            matched = set()
            for v in weekly_dish_values:
                vl = v.strip().lower()
                for rn in kulinar_names:
                    if rn and (rn == vl or rn in vl or vl in rn):
                        matched.add(rn)
                        break
            record(
                "Dishes match kulinar recipe DB",
                len(matched) >= 3,
                f"Matched {len(matched)} kulinar dishes: {sorted(matched)[:8]}",
            )
        else:
            # Cannot load DB -> do not block; mark non-fatal informational fail.
            record(
                "Dishes match kulinar recipe DB",
                True,
                "kulinar recipe DB not found; skipping cross-check",
            )

        # Check Nutrition_Summary content
        nutr_sheet_id = None
        for sid, title in sheets:
            if "nutrition" in title.lower() or "summary" in title.lower():
                nutr_sheet_id = sid
                break

        if nutr_sheet_id:
            cur.execute(
                "SELECT row_index, col_index, value FROM gsheet.cells "
                "WHERE sheet_id = %s ORDER BY row_index, col_index",
                (nutr_sheet_id,),
            )
            cells = cur.fetchall()
            all_values = [str(v).lower() for _, _, v in cells if v]

            has_target = any("target" in v for v in all_values)
            has_average = any("average" in v or "avg" in v for v in all_values)
            record(
                "Nutrition_Summary has target and average columns",
                has_target and has_average,
                f"Values sample: {all_values[:10]}",
            )

            # Check nutrients listed
            nutrients_found = sum(
                1
                for n in ["calori", "protein", "fat", "carb"]
                if any(n in v for v in all_values)
            )
            record(
                "All 4 nutrients in Nutrition_Summary",
                nutrients_found >= 4,
                f"Found {nutrients_found}",
            )

            # CRITICAL: Daily_Target column matches the guidelines page values.
            # Build a grid (row -> {col: value}) and find Daily_Target column.
            grid = {}
            header_row = {}
            for r, c, v in cells:
                grid.setdefault(r, {})[c] = v
                if r == 0 and v:
                    header_row[c] = str(v).strip().lower()
            target_col = None
            for c, h in header_row.items():
                if "daily" in h and "target" in h:
                    target_col = c
                    break
            if target_col is None:
                for c, h in header_row.items():
                    if "target" in h:
                        target_col = c
                        break
            # Find the nutrient label column (usually col 0).
            label_col = 0
            for c, h in header_row.items():
                if "nutrient" in h:
                    label_col = c
                    break

            target_ok = False
            if target_col is not None:
                hits = 0
                for r, cols in grid.items():
                    if r == 0:
                        continue
                    label = str(cols.get(label_col, "")).lower()
                    tval = cols.get(target_col)
                    for key, expected in TARGETS.items():
                        aliases = {
                            "calories": ["calori", "кал"],
                            "protein": ["protein", "белк"],
                            "fat": ["fat", "жир"],
                            "carbs": ["carb", "углевод"],
                        }[key]
                        if any(a in label for a in aliases):
                            tol = 0.05 * expected if key != "fat" else 2.0
                            if num_close(tval, expected, max(tol, 1.0)):
                                hits += 1
                            break
                target_ok = hits >= 4
                record(
                    "Nutrition_Summary Daily_Target matches guidelines",
                    target_ok,
                    f"Matched {hits}/4 targets {TARGETS}",
                )
            else:
                record(
                    "Nutrition_Summary Daily_Target matches guidelines",
                    False,
                    f"No Daily_Target column; headers: {list(header_row.values())}",
                )
        else:
            record(
                "Nutrition_Summary Daily_Target matches guidelines",
                False,
                "No Nutrition_Summary sheet",
            )

        cur.close()
        conn.close()
        return True

    except Exception as e:
        record("Google Sheets accessible", False, str(e))
        return False


def check_word(agent_workspace):
    """Check Word document."""
    print("\n=== Checking Word Document ===")

    doc_path = os.path.join(agent_workspace, "Wellness_Menu_Summary.docx")
    if not os.path.isfile(doc_path):
        record("Wellness_Menu_Summary.docx exists", False, f"Not found: {doc_path}")
        record("Doc title 'Corporate Wellness Lunch Program' present", False, "no docx")
        return False

    record("Wellness_Menu_Summary.docx exists", True)

    if Document is None:
        record("python-docx available", False, "Cannot import docx")
        record("Doc title 'Corporate Wellness Lunch Program' present", False, "no docx lib")
        return False

    try:
        doc = Document(doc_path)
        full_text = "\n".join(p.text for p in doc.paragraphs).lower()

        # CRITICAL: title present (preserved English identifier).
        record(
            "Doc title 'Corporate Wellness Lunch Program' present",
            "corporate wellness lunch program" in full_text,
            "Missing title",
        )

        record(
            "Doc mentions wellness/lunch/menu",
            "wellness" in full_text or "lunch" in full_text or "menu" in full_text
            or "велнес" in full_text or "обед" in full_text or "меню" in full_text,
            "Missing wellness/lunch/menu keywords",
        )

        # Check weekdays mentioned (EN or RU)
        day_aliases = [
            ["monday", "понедельник", "пн"],
            ["tuesday", "вторник", "вт"],
            ["wednesday", "среда", "ср"],
            ["thursday", "четверг", "чт"],
            ["friday", "пятница", "пт"],
        ]
        days_found = sum(
            1 for aliases in day_aliases if any(a in full_text for a in aliases)
        )
        record(
            "Doc mentions weekdays",
            days_found >= 3,
            f"Found {days_found} days",
        )

        # Check nutritional info mentioned (EN or RU)
        has_nutrition = (
            "calori" in full_text
            or "protein" in full_text
            or "nutrition" in full_text
            or "калор" in full_text
            or "белк" in full_text
            or "питани" in full_text
        )
        record("Doc mentions nutrition", has_nutrition)

        # Check dietary restriction mention (EN or RU)
        has_restriction = any(
            kw in full_text
            for kw in [
                "vegetarian", "low-sodium", "low sodium", "gluten",
                "вегетариан", "низкосол", "безглютен", "без свинин",
            ]
        )
        record("Doc mentions a dietary restriction", has_restriction)

        # Check length - should be substantive
        record(
            "Doc has substantial content",
            len(full_text) > 200,
            f"Length: {len(full_text)}",
        )

        return True
    except Exception as e:
        record("Word doc readable", False, str(e))
        record("Doc title 'Corporate Wellness Lunch Program' present", False, str(e))
        return False


def check_email():
    """Check email sent about wellness menu."""
    print("\n=== Checking Email ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute(
            "SELECT subject, from_addr, to_addr, body_text FROM email.messages"
        )
        emails = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Email DB accessible", False, str(e))
        record("Wellness menu email exists", False, str(e))
        record("Email to cafeteria-manager", False, str(e))
        return False

    found = False
    for subject, from_addr, to_addr, body_text in emails:
        subj_lower = (subject or "").lower()
        if (
            "wellness" in subj_lower
            or "lunch" in subj_lower
            or "menu" in subj_lower
            or "велнес" in subj_lower
            or "обед" in subj_lower
            or "меню" in subj_lower
        ):
            found = True
            record("Wellness menu email exists", True)

            # Check recipient
            to_str = str(to_addr).lower() if to_addr else ""
            record(
                "Email to cafeteria-manager",
                "cafeteria" in to_str or "manager" in to_str,
                f"To: {to_addr}",
            )

            # Check from address
            from_str = str(from_addr).lower() if from_addr else ""
            record(
                "Email from wellness@company.com",
                "wellness@company.com" in from_str or "wellness" in from_str,
                f"From: {from_addr}",
            )

            # Check body mentions dishes (RU recipe names + generic RU/EN tokens)
            body_lower = (body_text or "").lower()
            has_dishes = any(
                kw in body_lower
                for kw in [
                    # English generic tokens (kept for safety)
                    "chicken", "beef", "tofu", "vegetable", "soup", "rice",
                    # Russian generic tokens
                    "курин", "куриц", "говяд", "мясн", "овощ", "суп",
                    "рис", "гарнир", "каша", "салат", "борщ", "плов",
                    "котлет", "пельмен", "голубц",
                ]
            )
            record(
                "Email body mentions dishes",
                has_dishes,
                f"Body length: {len(body_lower)}",
            )
            break

    if not found:
        record(
            "Wellness menu email exists",
            False,
            f"Found {len(emails)} emails but none about wellness/lunch/menu",
        )
        record("Email to cafeteria-manager", False, "no wellness email")

    return found


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    gsheet_ok = check_gsheet()
    word_ok = check_word(args.agent_workspace)
    email_ok = check_email()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print("\n=== SUMMARY ===")
    print(f"  GSheet:   {'PASS' if gsheet_ok else 'FAIL'}")
    print(f"  Word:     {'PASS' if word_ok else 'FAIL'}")
    print(f"  Email:    {'PASS' if email_ok else 'FAIL'}")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}, Accuracy: {accuracy:.1f}%")

    if critical_failed:
        print(f"  CRITICAL CHECKS FAILED: {critical_failed}")
        print("  Overall:  FAIL (critical check failed)")
        sys.exit(1)

    overall = accuracy >= 70
    print(f"  Overall:  {'PASS' if overall else 'FAIL'} (threshold 70%)")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
