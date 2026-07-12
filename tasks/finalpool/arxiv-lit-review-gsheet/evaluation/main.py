"""
Evaluation script for arxiv-lit-review-gsheet task.

Checks:
1. Google Sheet spreadsheet exists with "prompt engineering" or "literature review" in title
2. "Paper Comparison" sheet exists with at least 5 data rows
3. Paper IDs match the 5 injected target papers
4. Citation counts approximately match expected values
5. "Technique Analysis" sheet exists with at least 3 rows
6. review_summary.txt exists in workspace
7. Memory file has been updated with entities

CRITICAL checks (semantic substance, any failure => hard FAIL via sys.exit(1)):
  - Paper Comparison contains exactly the 5 target IDs and ZERO noise IDs
  - >=4 of 5 papers have Citation_Count matching seed (tol=300) tied to correct Paper_ID
  - Technique Analysis Requires_Examples values are correct per source
  - Each of the 5 Paper Comparison rows has a non-empty Methodology_Summary
  - review_summary.txt names all 5 target papers and contains a RU/EN synthesis
"""

import argparse
import json
import os
import re
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

TARGET_IDS = ["2201.11903", "2203.11171", "2210.03493", "2205.11916", "2305.10601"]
NOISE_IDS = ["1301.03781", "1310.04546", "1405.01512"]

TARGET_TITLES_LOWER = [
    "chain-of-thought prompting",
    "self-consistency",
    "automatic chain of thought",
    "zero-shot reasoners",
    "tree of thoughts",
]

EXPECTED_CITATIONS = {
    "2201.11903": 6500,
    "2203.11171": 3200,
    "2210.03493": 1800,
    "2205.11916": 4100,
    "2305.10601": 2400,
}

# Map each target paper to its correct Requires_Examples value (per source manuscripts).
EXPECTED_REQUIRES_EXAMPLES = {
    "2201.11903": "yes",  # Chain-of-Thought (few-shot)
    "2203.11171": "yes",  # Self-Consistency (built on CoT few-shot)
    "2210.03493": "no",   # Auto-CoT (automatic, no manual examples)
    "2205.11916": "no",   # Zero-shot-CoT
    "2305.10601": "yes",  # Tree of Thoughts (generalizes CoT; hand-crafted few-shot propose/value prompts)
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILURES = []


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS]{' [CRITICAL]' if critical else ''} {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL]{' [CRITICAL]' if critical else ''} {name}{detail_str}")
        if critical:
            CRITICAL_FAILURES.append(name)


def num_close(a, b, tol=500):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_gsheet_data():
    """Read Google Sheet data from the database."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        ORDER BY created_at DESC
    """)
    spreadsheets = cur.fetchall()

    result = {"spreadsheet": None, "sheets": {}, "cells": {}}

    for ss_id, ss_title in spreadsheets:
        title_lower = ss_title.lower()
        if "prompt" in title_lower or "literature" in title_lower or "engineering" in title_lower:
            result["spreadsheet"] = (ss_id, ss_title)
            break

    if not result["spreadsheet"]:
        if spreadsheets:
            result["spreadsheet"] = spreadsheets[0]

    if result["spreadsheet"]:
        ss_id = result["spreadsheet"][0]

        cur.execute("""
            SELECT id, title FROM gsheet.sheets
            WHERE spreadsheet_id = %s
            ORDER BY index
        """, (ss_id,))
        for sheet_id, sheet_title in cur.fetchall():
            result["sheets"][sheet_title.lower()] = sheet_id

            cur.execute("""
                SELECT row_index, col_index, value
                FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, sheet_id))

            cells = {}
            for row_idx, col_idx, value in cur.fetchall():
                if row_idx not in cells:
                    cells[row_idx] = {}
                cells[row_idx][col_idx] = value

            result["cells"][sheet_title.lower()] = cells

    cur.close()
    conn.close()
    return result


def _header_map(cells):
    """Return {column_name_lower: col_index} from header row 0."""
    header = cells.get(0, {})
    return {str(v).strip().lower(): c for c, v in header.items() if v is not None}


def check_gsheet():
    """Check Google Sheet content."""
    print("\n=== Checking Google Sheet ===")

    data = get_gsheet_data()

    check("Spreadsheet exists", data["spreadsheet"] is not None,
          "No spreadsheet found")
    if not data["spreadsheet"]:
        return

    ss_id, ss_title = data["spreadsheet"]
    print(f"  Found spreadsheet: {ss_title}")

    # --- Paper Comparison sheet ---
    paper_sheet_key = None
    for key in data["sheets"]:
        if "paper" in key or "comparison" in key or "index" in key:
            paper_sheet_key = key
            break

    check("Paper Comparison sheet exists", paper_sheet_key is not None,
          f"Sheets found: {list(data['sheets'].keys())}")

    if paper_sheet_key and paper_sheet_key in data["cells"]:
        cells = data["cells"][paper_sheet_key]
        data_rows = {r: cells[r] for r in cells if r > 0}
        check("Paper sheet has at least 5 data rows",
              len(data_rows) >= 5,
              f"Found {len(data_rows)} data rows")

        all_values = " ".join(
            str(v).lower() for row in data_rows.values() for v in row.values()
        )

        # Loose found-papers check (non-critical, original behavior)
        found_papers = 0
        for title_kw in TARGET_TITLES_LOWER:
            if title_kw in all_values:
                found_papers += 1
        for pid in TARGET_IDS:
            if pid in all_values:
                found_papers += 1
        found_papers = min(found_papers, 5)
        check("At least 3 target papers in Paper sheet",
              found_papers >= 3,
              f"Found {found_papers}/5 target papers")

        # Loose noise check (non-critical, original behavior)
        noise_found = 0
        for nid in NOISE_IDS:
            if nid in all_values:
                noise_found += 1
        noise_titles = ["word2vec", "glove", "word representations", "distributed representations"]
        for nt in noise_titles:
            if nt in all_values:
                noise_found += 1
        check("Noise papers excluded (at most 1 noise)",
              noise_found <= 1,
              f"Found {noise_found} noise paper references")

        # --- CRITICAL: exactly the 5 target IDs and ZERO noise IDs ---
        ids_present = [pid for pid in TARGET_IDS if pid in all_values]
        noise_ids_present = [nid for nid in NOISE_IDS if nid in all_values]
        check("CRITICAL: all 5 target IDs present and 0 noise IDs",
              len(ids_present) == 5 and len(noise_ids_present) == 0,
              f"targets={ids_present}, noise={noise_ids_present}",
              critical=True)

        # --- CRITICAL: >=4 of 5 citation counts tied to correct Paper_ID (tol=300) ---
        # Exclude 4-digit years (19xx/20xx) from the citation number scan to avoid
        # a Year cell (e.g. 2022) spuriously satisfying the tolerance.
        citation_checks = 0
        for row_data in data_rows.values():
            row_text = " ".join(str(v) for v in row_data.values())
            for pid, expected_count in EXPECTED_CITATIONS.items():
                if pid in row_text:
                    numbers = re.findall(r'\d+', row_text)
                    for num_str in numbers:
                        if re.fullmatch(r'(19|20)\d{2}', num_str):
                            continue  # skip plausible 4-digit years
                        if num_close(num_str, expected_count, 300):
                            citation_checks += 1
                            break
        check("CRITICAL: >=4 papers have correct Citation_Count by Paper_ID (tol=300)",
              citation_checks >= 4,
              f"Found {citation_checks}/5 papers with matching citations",
              critical=True)

        # --- CRITICAL: each of the 5 rows has a non-empty Methodology_Summary ---
        hmap = _header_map(cells)
        methodology_col = None
        for cname, cidx in hmap.items():
            if "methodology" in cname:
                methodology_col = cidx
                break
        if methodology_col is None:
            check("CRITICAL: Methodology_Summary column present and filled", False,
                  f"No Methodology_Summary column; headers={list(hmap.keys())}",
                  critical=True)
        else:
            # Count target-paper rows whose methodology cell is non-trivially filled.
            filled = 0
            for row_data in data_rows.values():
                row_text = " ".join(str(v) for v in row_data.values())
                if not any(pid in row_text for pid in TARGET_IDS):
                    continue
                meth = str(row_data.get(methodology_col, "") or "").strip()
                if len(meth) >= 15:
                    filled += 1
            check("CRITICAL: >=5 target rows have non-empty Methodology_Summary (>=15 chars)",
                  filled >= 5,
                  f"{filled} rows with substantive methodology summary",
                  critical=True)

    # --- Technique Analysis sheet ---
    technique_sheet_key = None
    for key in data["sheets"]:
        if "technique" in key or "analysis" in key or "method" in key:
            technique_sheet_key = key
            break

    check("Technique Analysis sheet exists", technique_sheet_key is not None,
          f"Sheets found: {list(data['sheets'].keys())}")

    if technique_sheet_key and technique_sheet_key in data["cells"]:
        cells = data["cells"][technique_sheet_key]
        data_rows = {r: cells[r] for r in cells if r > 0}
        check("Technique sheet has at least 3 data rows",
              len(data_rows) >= 3,
              f"Found {len(data_rows)} data rows")

        all_values = " ".join(
            str(v).lower() for row in data_rows.values() for v in row.values()
        )
        has_technique_content = any(
            kw in all_values for kw in [
                "chain", "thought", "self-consistency", "zero-shot",
                "tree", "auto", "prompting", "reasoning"
            ]
        )
        check("Technique sheet has prompting-related content",
              has_technique_content,
              "No prompting technique keywords found")

        # --- CRITICAL: Requires_Examples values correct per source ---
        hmap = _header_map(cells)
        req_col = None
        for cname, cidx in hmap.items():
            if "requires_examples" in cname or "requires examples" in cname:
                req_col = cidx
                break
        if req_col is None:
            check("CRITICAL: Requires_Examples values correct per source", False,
                  f"No Requires_Examples column; headers={list(hmap.keys())}",
                  critical=True)
        else:
            correct = 0
            checked = 0
            for row_data in data_rows.values():
                row_text = " ".join(str(v) for v in row_data.values())
                for pid, expected in EXPECTED_REQUIRES_EXAMPLES.items():
                    if pid in row_text:
                        checked += 1
                        cell_val = str(row_data.get(req_col, "") or "").strip().lower()
                        # accept yes/no in EN; also tolerate RU да/нет if used
                        if expected == "yes":
                            ok = cell_val.startswith("yes") or cell_val.startswith("да")
                        else:
                            ok = cell_val.startswith("no") or cell_val.startswith("нет")
                        if ok:
                            correct += 1
                        break
            check("CRITICAL: all 5 Requires_Examples values correct per source",
                  checked >= 5 and correct >= 5,
                  f"{correct}/{checked} rows correct (need 5/5)",
                  critical=True)


def check_review_summary(agent_workspace):
    """Check review_summary.txt exists and has content."""
    print("\n=== Checking review_summary.txt ===")

    summary_path = os.path.join(agent_workspace, "review_summary.txt")
    check("review_summary.txt exists", os.path.isfile(summary_path),
          f"Not found at {summary_path}")

    if os.path.isfile(summary_path):
        with open(summary_path, "r") as f:
            content = f.read()

        check("review_summary.txt has at least 200 characters",
              len(content.strip()) >= 200,
              f"File has {len(content.strip())} characters")

        content_lower = content.lower()
        papers_mentioned = sum(1 for kw in TARGET_TITLES_LOWER if kw in content_lower)
        check("review_summary mentions at least 3 papers",
              papers_mentioned >= 3,
              f"Found {papers_mentioned}/5 paper references")

        # --- CRITICAL: names >=4 of 5 papers AND contains a synthesis (RU or EN) ---
        # RU keyword check operates on the ORIGINAL lowercased text (never normalized).
        ru_kw = ["промпт", "рассужден", "обзор", "модел", "техник", "цепоч", "мысл"]
        en_kw = ["prompt", "reasoning", "review", "synthesis", "technique"]
        has_synthesis = any(k in content_lower for k in ru_kw) or \
                        any(k in content_lower for k in en_kw)
        check("CRITICAL: review_summary names >=4 papers and has a RU/EN synthesis",
              papers_mentioned >= 4 and has_synthesis,
              f"papers={papers_mentioned}/5, synthesis_present={has_synthesis}",
              critical=True)


def check_memory(agent_workspace):
    """Check memory file has been updated with entities."""
    print("\n=== Checking Memory ===")

    memory_path = os.path.join(agent_workspace, "memory", "memory.json")
    check("memory.json exists", os.path.isfile(memory_path),
          f"Not found at {memory_path}")

    if os.path.isfile(memory_path):
        with open(memory_path, "r") as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                check("memory.json is valid JSON", False, "JSON parse error")
                return

        check("memory.json is valid JSON", True)

        entities = data.get("entities", [])
        check("memory has at least 1 entity", len(entities) >= 1,
              f"Found {len(entities)} entities")

        has_observations = any(
            len(e.get("observations", [])) > 0 for e in entities
        )
        check("At least one entity has observations", has_observations,
              "No entities with observations found")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_review_summary(args.agent_workspace)
    check_memory(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks performed.")
        sys.exit(1)

    pass_rate = PASS_COUNT / total
    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Pass Rate: {pass_rate:.1%}")
    if CRITICAL_FAILURES:
        print(f"  CRITICAL FAILURES: {CRITICAL_FAILURES}")

    accuracy = pass_rate * 100
    success = (not CRITICAL_FAILURES) and accuracy >= 70

    result = {
        "passed": PASS_COUNT,
        "failed": FAIL_COUNT,
        "pass_rate": round(pass_rate, 3),
        "accuracy": round(accuracy, 1),
        "critical_failures": CRITICAL_FAILURES,
        "success": success,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # Any critical failure => hard fail before the accuracy gate.
    if CRITICAL_FAILURES:
        print("\nFAIL: one or more CRITICAL checks failed.")
        sys.exit(1)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
