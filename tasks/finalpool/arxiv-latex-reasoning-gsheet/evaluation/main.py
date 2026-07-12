"""
Evaluation for arxiv-latex-reasoning-gsheet task.
Checks Google Sheet and Word document.

Порог прохождения: accuracy >= 70 И ни одна CRITICAL-проверка не провалена.
CRITICAL-проверки (семантические, любой провал => немедленный FAIL):
  - В листе "Papers" присутствуют все 5 ожидаемых статей (по названию).
  - NOISE-исключение: ни таблица, ни Word-документ не упоминают статьи про
    векторные представления слов (word2vec / glove / skip-gram / word representation).
  - Корректность содержимого колонки Method: для каждой из 5 строк ячейка Method
    называет правильный подход (per-paper), а не просто наличие заголовка колонки.
  - В Word-документе есть точный заголовок и дата.
Структурные/мягкие проверки (лист/колонка/файл существуют, "5 строк", читается)
помечены как НЕ критические.
Свободный текст (дата, описания вклада) сопоставляется по RU+EN — агент
легитимно пишет Key_Contribution и прозу .docx на русском.
"""
import argparse
import json
import os
import sys

import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# The 5 reasoning papers from arxiv_latex.papers
EXPECTED_REASONING_PAPERS = [
    {"id": "2201.11903", "title": "Chain-of-Thought Prompting Elicits Reasoning in Large Language Models"},
    {"id": "2203.11171", "title": "Self-Consistency Improves Chain of Thought Reasoning in Language Models"},
    {"id": "2205.11916", "title": "Large Language Models are Zero-Shot Reasoners"},
    {"id": "2210.03493", "title": "Automatic Chain of Thought Prompting in Large Language Models"},
    {"id": "2305.10601", "title": "Tree of Thoughts: Deliberate Problem Solving with Large Language Models"},
]

# Keywords identifying the contribution / method per paper (for Word prose).
PAPER_KEYWORDS = [
    ["chain-of-thought prompting", "chain of thought prompting"],
    ["self-consistency", "self consistency"],
    ["zero-shot", "zero shot"],
    ["automatic chain", "auto-cot", "auto cot"],
    ["tree of thoughts"],
]

# Per-paper acceptable Method-column values (content correctness, EN method names).
METHOD_KEYWORDS = [
    ["chain-of-thought prompting", "chain of thought prompting", "cot prompting"],
    ["self-consistency", "self consistency"],
    ["zero-shot", "zero shot"],
    ["automatic", "auto-cot", "auto cot"],
    ["tree of thoughts", "tree-of-thoughts"],
]

# Word embedding papers that should NOT be included
NOISE_KEYWORDS = ["word2vec", "glove", "word representation", "skip-gram", "skip gram"]

# Семантические проверки: любой провал => немедленный FAIL независимо от accuracy.
CRITICAL_CHECKS = {
    "Лист 'Papers' содержит все 5 ожидаемых статей",
    "Содержимое колонки Method корректно для всех 5 статей",
    "NOISE-исключение: таблица не содержит статей про word embeddings",
    "NOISE-исключение: Word-документ не содержит статей про word embeddings",
    "Word-документ содержит корректный заголовок",
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_gsheet():
    """Check Google Sheet exists with correct data."""
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Find spreadsheet
        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        target_ss = None
        for sid, title in spreadsheets:
            if title and "reasoning" in title.lower():
                target_ss = sid
                break

        record("Google Sheet 'Reasoning Methods Comparison' exists",
               target_ss is not None,
               f"Found spreadsheets: {[t for _, t in spreadsheets]}")

        if target_ss is None:
            conn.close()
            return

        # Check "Papers" sheet exists
        cur.execute("""
            SELECT id, title FROM gsheet.sheets
            WHERE spreadsheet_id = %s
        """, (target_ss,))
        sheets = cur.fetchall()
        sheet_names = [t for _, t in sheets]

        papers_sheet_id = None
        for sid, sname in sheets:
            if sname and sname.strip().lower() == "papers":
                papers_sheet_id = sid
                break

        record("Sheet 'Papers' exists", papers_sheet_id is not None,
               f"Found sheets: {sheet_names}")

        if papers_sheet_id is None:
            conn.close()
            return

        # Read cells from Papers sheet
        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
            ORDER BY row_index, col_index
        """, (target_ss, papers_sheet_id))
        cells = cur.fetchall()

        # Build grid
        grid = {}
        for row_idx, col_idx, val in cells:
            if row_idx not in grid:
                grid[row_idx] = {}
            grid[row_idx][col_idx] = val

        if not grid:
            record("Papers sheet has data", False, "No cells found")
            conn.close()
            return

        min_row = min(grid.keys())
        header_row = grid.get(min_row, {})
        header_vals = [header_row.get(i, "") for i in range(max(header_row.keys()) + 1)] if header_row else []

        # Find Title column
        title_col = None
        for i, h in enumerate(header_vals):
            if h and str(h).strip().lower() == "title":
                title_col = i
                break

        record("Title column exists", title_col is not None, f"Header: {header_vals}")

        # Check Method column
        method_col = None
        for i, h in enumerate(header_vals):
            if h and str(h).strip().lower() == "method":
                method_col = i
                break
        record("Method column exists", method_col is not None, f"Header: {header_vals}")

        # Check Key_Contribution column
        kc_col = None
        for i, h in enumerate(header_vals):
            if h and "contribution" in str(h).strip().lower():
                kc_col = i
                break
        record("Key_Contribution column exists", kc_col is not None, f"Header: {header_vals}")

        # Check Year column (structural)
        year_col = None
        for i, h in enumerate(header_vals):
            if h and str(h).strip().lower() == "year":
                year_col = i
                break
        record("Year column exists", year_col is not None, f"Header: {header_vals}")

        # Data rows
        data_rows = {r: grid[r] for r in grid if r > min_row}
        record("Papers sheet has 5 data rows", len(data_rows) == 5,
               f"Found {len(data_rows)} rows")

        # --- CRITICAL: all 5 expected paper titles present ---
        all_titles_found = True
        per_title_lower = {}  # row -> title text lower
        if title_col is not None:
            found_titles = []
            for r in sorted(data_rows.keys()):
                val = data_rows[r].get(title_col, "")
                if val:
                    found_titles.append(str(val).lower())
                    per_title_lower[r] = str(val).lower()

            missing = []
            for paper in EXPECTED_REASONING_PAPERS:
                t_lower = paper["title"].lower()
                found = any(t_lower in t or t in t_lower for t in found_titles)
                if not found:
                    missing.append(paper["title"][:40])
                    all_titles_found = False
        else:
            all_titles_found = False
            missing = ["<no Title column>"]
        record("Лист 'Papers' содержит все 5 ожидаемых статей",
               all_titles_found, f"Отсутствуют: {missing}")

        # --- CRITICAL: Method column content correctness (per-paper) ---
        method_ok = True
        method_detail = []
        if method_col is not None and title_col is not None:
            for paper, kws in zip(EXPECTED_REASONING_PAPERS, METHOD_KEYWORDS):
                t_lower = paper["title"].lower()
                # find the row whose title matches this paper
                matched_row = None
                for r, tl in per_title_lower.items():
                    if t_lower in tl or tl in t_lower:
                        matched_row = r
                        break
                if matched_row is None:
                    method_ok = False
                    method_detail.append(f"{paper['title'][:30]}: row not found")
                    continue
                mval = str(data_rows[matched_row].get(method_col, "") or "").lower()
                if not any(kw in mval for kw in kws):
                    method_ok = False
                    method_detail.append(f"{paper['title'][:30]}: method='{mval[:40]}'")
        else:
            method_ok = False
            method_detail.append("Method/Title column missing")
        record("Содержимое колонки Method корректно для всех 5 статей",
               method_ok, "; ".join(method_detail))

        # --- CRITICAL: NOISE exclusion in sheet (no word-embedding papers) ---
        all_sheet_text = " ".join(
            str(v or "").lower()
            for row in grid.values() for v in row.values()
        )
        noise_in_sheet = [n for n in NOISE_KEYWORDS if n in all_sheet_text]
        record("NOISE-исключение: таблица не содержит статей про word embeddings",
               len(noise_in_sheet) == 0, f"Найдены: {noise_in_sheet}")

        conn.close()
    except Exception as e:
        record("GSheet connection", False, str(e))


def check_word(agent_workspace):
    """Check Word document."""
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "Reasoning_Methods_Review.docx")

    if not os.path.isfile(doc_path):
        record("Word file exists", False, f"Not found: {doc_path}")
        return

    record("Word file exists", True)

    try:
        doc = Document(doc_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        return

    record("Word file readable", True)

    full_text = "\n".join(p.text for p in doc.paragraphs)
    full_lower = full_text.lower()

    # CRITICAL: title (EN, preserved)
    has_title = "chain-of-thought" in full_lower and "reasoning" in full_lower and "comparison" in full_lower
    if not has_title:
        has_title = "chain of thought" in full_lower and "methods" in full_lower
    record("Word-документ содержит корректный заголовок", has_title)

    # Date: accept ISO or Russian wording.
    has_date = (
        "2026-03-06" in full_text
        or "march 6, 2026" in full_lower
        or "march 2026" in full_lower
        or "6 марта 2026" in full_lower
        or "март 2026" in full_lower
        or "марта 2026" in full_lower
    )
    record("Word-документ содержит дату (2026-03-06 / 6 марта 2026)", has_date)

    # Each reasoning paper is mentioned (contribution keyword present).
    for i, paper in enumerate(EXPECTED_REASONING_PAPERS):
        found = any(kw in full_lower for kw in PAPER_KEYWORDS[i])
        record(f"Word упоминает: {paper['title'][:50]}...", found,
               f"Keywords: {PAPER_KEYWORDS[i]}")

    # Документ содержит непустую прозу (вклад/заключение) — мягкая проверка
    # содержательности, текст может быть на русском.
    record("Word-документ содержит развёрнутую прозу",
           len(full_text.strip()) >= 400,
           f"len={len(full_text.strip())}")

    # CRITICAL: noise papers NOT mentioned (single aggregated critical check).
    noise_in_doc = [n for n in NOISE_KEYWORDS if n in full_lower]
    record("NOISE-исключение: Word-документ не содержит статей про word embeddings",
           len(noise_in_doc) == 0, f"Найдены: {noise_in_doc}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_word(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    print(f"\n=== Итого: {PASS_COUNT}/{total} проверок пройдено ({accuracy:.1f}%) ===")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "critical_failed": critical_failed,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if critical_failed:
        print(f"CRITICAL FAILURES: {critical_failed}")
        print("FAIL (провалена критическая проверка)")
        sys.exit(1)

    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
