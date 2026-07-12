"""
Evaluation for terminal-arxiv-latex-pdf-gsheet-word task.

Checks:
1. Google Sheet "Paper Review Matrix" with 3 sheets (Review Scores, Methodology Comparison, Rankings)
2. Conference_Review_Summary.docx
3. Intermediate JSON files (methodology_analysis.json, comparison_matrix.json, final_rankings.json)
"""
import argparse
import json
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

TARGET_IDS = {"2401.00101", "2401.00102", "2401.00103", "2401.00104"}
NOISE_IDS = {"2401.00201", "2401.00202"}

CRITERIA = ["novelty", "methodology_rigor", "experimental_completeness",
            "clarity", "significance"]


def rubric_rec(total):
    """Rubric mapping (from the RU/EN PDF): >=20 Accept, 15..19 Revise, <15 Reject."""
    if total >= 20:
        return "accept"
    if total >= 15:
        return "revise"
    return "reject"


# CRITICAL checks: any failure here => overall FAIL regardless of accuracy.
# These verify SUBSTANCE / internal consistency, NOT pinned subjective scores
# (per-paper scores are the agent's own judgement; we only enforce the rubric
# and the noise-exclusion that the task mandates).
CRITICAL_CHECKS = {
    "Review Scores contains exactly the 4 target papers and NO noise papers",
    "comparison_matrix scores are integers 1..5 and total == sum of criteria",
    "Rankings recommendations follow the rubric mapping derived from totals",
    "final_rankings.json is sorted by total_score descending",
    "Word doc references all 4 target papers and NO noise paper",
}


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def num_close(a, b, tol=2.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def check_gsheet():
    """Check Google Sheet via database."""
    print("\n=== Checking Google Sheet 'Paper Review Matrix' ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()

        # Find spreadsheet
        cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%paper review%'")
        rows = cur.fetchall()
        check("Spreadsheet 'Paper Review Matrix' exists", len(rows) >= 1,
              f"Found {len(rows)} matching spreadsheets")
        if not rows:
            cur.close(); conn.close()
            return

        ss_id = rows[0][0]

        # Check sheets
        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
        sheets = cur.fetchall()
        sheet_names = {s[1].strip().lower(): s[0] for s in sheets}

        # Sheet 1: Review Scores
        review_key = None
        for name, sid in sheet_names.items():
            if "review" in name and "score" in name:
                review_key = sid
                break
        check("Sheet 'Review Scores' exists", review_key is not None,
              f"Sheets: {list(sheet_names.keys())}")

        if review_key is not None:
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, review_key))
            cells = cur.fetchall()

            # Build grid
            grid = {}
            for r, c, v in cells:
                grid[(r, c)] = v

            # Check header row (row 0)
            headers = [grid.get((0, c), "") for c in range(8)]
            header_lower = [str(h).lower() for h in headers]
            check("Review Scores has Paper_ID column",
                  any("paper" in h and "id" in h for h in header_lower),
                  f"Headers: {headers}")
            check("Review Scores has Total_Score column",
                  any("total" in h for h in header_lower),
                  f"Headers: {headers}")

            # Check 4 data rows
            data_rows = set()
            for (r, c), v in grid.items():
                if r >= 1 and c == 0 and v:
                    data_rows.add(str(v).strip())
            check("Review Scores has 4 target papers",
                  data_rows.issuperset(TARGET_IDS),
                  f"Found IDs: {data_rows}")

            # CRITICAL: exactly the 4 target IDs and ZERO noise IDs.
            noise_in_review = data_rows.intersection(NOISE_IDS)
            check("Review Scores contains exactly the 4 target papers and NO noise papers",
                  data_rows == TARGET_IDS,
                  f"Found IDs: {sorted(data_rows)} (noise present: {sorted(noise_in_review)})")

            # Check total scores are reasonable (between 5 and 25)
            total_col = None
            for c in range(8):
                if "total" in str(grid.get((0, c), "")).lower():
                    total_col = c
                    break
            if total_col is not None:
                for r in range(1, 5):
                    val = grid.get((r, total_col))
                    if val is not None:
                        check(f"Row {r} total score in range",
                              5 <= float(val) <= 25,
                              f"Got {val}")

        # Sheet 2: Methodology Comparison
        method_key = None
        for name, sid in sheet_names.items():
            if "method" in name:
                method_key = sid
                break
        check("Sheet 'Methodology Comparison' exists", method_key is not None,
              f"Sheets: {list(sheet_names.keys())}")

        if method_key is not None:
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, method_key))
            cells = cur.fetchall()
            grid = {}
            for r, c, v in cells:
                grid[(r, c)] = v

            data_rows = set()
            for (r, c), v in grid.items():
                if r >= 1 and c == 0 and v:
                    data_rows.add(str(v).strip())
            check("Methodology Comparison has target papers",
                  len(data_rows.intersection(TARGET_IDS)) >= 4,
                  f"Found: {data_rows}")

            # Check methods column has content
            methods_populated = 0
            for r in range(1, 5):
                val = grid.get((r, 2))  # Methods_Used column
                if val and len(str(val).strip()) > 0:
                    methods_populated += 1
            check("Methodology rows have methods content",
                  methods_populated >= 3,
                  f"Populated: {methods_populated}")

        # Sheet 3: Rankings
        rank_key = None
        for name, sid in sheet_names.items():
            if "rank" in name:
                rank_key = sid
                break
        check("Sheet 'Rankings' exists", rank_key is not None,
              f"Sheets: {list(sheet_names.keys())}")

        if rank_key is not None:
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, rank_key))
            cells = cur.fetchall()
            grid = {}
            for r, c, v in cells:
                grid[(r, c)] = v

            headers = [str(grid.get((0, c), "")).lower() for c in range(6)]
            check("Rankings has Recommendation column",
                  any("recommend" in h for h in headers),
                  f"Headers: {headers}")

            # Locate relevant columns by header.
            rec_col = total_col = rank_col = None
            for c in range(6):
                h = str(grid.get((0, c), "")).lower()
                if rec_col is None and "recommend" in h:
                    rec_col = c
                if total_col is None and "total" in h:
                    total_col = c
                if rank_col is None and "rank" in h:
                    rank_col = c

            if rec_col is not None:
                recs = set()
                for r in range(1, 5):
                    val = grid.get((r, rec_col))
                    if val:
                        recs.add(str(val).strip().lower())
                check("Rankings has Accept/Revise/Reject values",
                      recs.issubset({"accept", "revise", "reject"}) and len(recs) >= 2,
                      f"Found: {recs}")
                check("At least one paper recommended Accept",
                      "accept" in recs,
                      f"Recommendations: {recs}")

            # CRITICAL: recommendation follows the rubric mapping derived from
            # each row's OWN total (not pinned expected values).
            if rec_col is not None and total_col is not None:
                mismatches = []
                for r in range(1, 5):
                    tv = grid.get((r, total_col))
                    rv = grid.get((r, rec_col))
                    if tv is None or rv is None:
                        continue
                    try:
                        expected = rubric_rec(float(tv))
                    except (TypeError, ValueError):
                        continue
                    if str(rv).strip().lower() != expected:
                        mismatches.append((r, tv, str(rv).strip(), expected))
                check("Rankings recommendations follow the rubric mapping derived from totals",
                      len(mismatches) == 0,
                      f"Mismatches (row,total,got,expected): {mismatches}")
            else:
                check("Rankings recommendations follow the rubric mapping derived from totals",
                      False, "Could not locate Recommendation/Total_Score columns")

            # NON-critical: Rank column ordering matches totals descending.
            if rank_col is not None and total_col is not None:
                ordered = []
                for r in range(1, 5):
                    rk = grid.get((r, rank_col))
                    tv = grid.get((r, total_col))
                    if rk is not None and tv is not None:
                        try:
                            ordered.append((int(float(rk)), float(tv)))
                        except (TypeError, ValueError):
                            pass
                ordered.sort(key=lambda x: x[0])
                totals_by_rank = [t for _, t in ordered]
                check("Rankings rows ordered by Rank have non-increasing Total_Score",
                      all(totals_by_rank[i] >= totals_by_rank[i + 1]
                          for i in range(len(totals_by_rank) - 1)),
                      f"Totals by rank: {totals_by_rank}")

        cur.close()
        conn.close()
    except Exception as e:
        check("GSheet check", False, str(e))


def check_word(agent_workspace):
    """Check Conference_Review_Summary.docx."""
    print("\n=== Checking Conference_Review_Summary.docx ===")
    docx_path = os.path.join(agent_workspace, "Conference_Review_Summary.docx")
    check("Conference_Review_Summary.docx exists", os.path.isfile(docx_path))
    if not os.path.isfile(docx_path):
        return
    try:
        from docx import Document
        doc = Document(docx_path)
        text = " ".join(p.text for p in doc.paragraphs).lower()
        check("Document has substantial content", len(text) > 500, f"Length: {len(text)}")

        # Check sections (agent may write RU or EN headings -> accept both).
        check("Contains 'overview' section (Обзор)",
              "overview" in text or "обзор" in text)
        check("Contains 'per-paper review' or individual reviews",
              "per-paper" in text or "per paper" in text
              or "покритериальный" in text or "по статьям" in text
              or "2401.00101" in text)
        check("Contains 'comparative analysis' (Сравнительный анализ)",
              "comparative" in text or "comparison" in text
              or "сравнительн" in text or "сравнение" in text)
        check("Contains 'recommendation' section (Рекомендации)",
              "recommendation" in text or "рекомендац" in text)

        # CRITICAL: references all 4 target papers and NO noise paper.
        all_targets = all(pid in text for pid in TARGET_IDS)
        no_noise = all(nid not in text for nid in NOISE_IDS)
        check("Word doc references all 4 target papers and NO noise paper",
              all_targets and no_noise,
              f"targets_present={all_targets} no_noise={no_noise}")

        # Non-critical per-paper mentions (visibility).
        for pid in TARGET_IDS:
            check(f"Mentions paper {pid}", pid in text, "Not found in document")

        # Check key terms (RU+EN).
        check("Mentions strengths/weaknesses",
              "strength" in text or "weakness" in text
              or "сильны" in text or "слабы" in text)
        check("Contains accept/revise/reject recommendations",
              ("accept" in text and "revise" in text) or "reject" in text)

    except ImportError:
        check("python-docx available", False)
    except Exception as e:
        check("Word document readable", False, str(e))


def _entry_id(e):
    """Case-insensitive paper-id lookup (paper_id / Paper_ID / id / paperid)."""
    if not isinstance(e, dict):
        return ""
    return next((str(v).strip() for k, v in e.items()
                 if str(k).lower() in ("paper_id", "id", "paperid")), "")


def check_json_files(agent_workspace):
    """Check intermediate JSON files."""
    print("\n=== Checking Intermediate JSON Files ===")

    # methodology_analysis.json
    ma_path = os.path.join(agent_workspace, "methodology_analysis.json")
    check("methodology_analysis.json exists", os.path.isfile(ma_path))
    if os.path.isfile(ma_path):
        try:
            with open(ma_path) as f:
                ma = json.load(f)
            if isinstance(ma, list):
                ids = {_entry_id(e) for e in ma}
            elif isinstance(ma, dict):
                ids = set(ma.keys())
            else:
                ids = set()
            check("methodology_analysis has 4 target papers",
                  ids.issuperset(TARGET_IDS),
                  f"Found: {ids}")
        except Exception as e:
            check("methodology_analysis readable", False, str(e))

    # comparison_matrix.json
    cm_path = os.path.join(agent_workspace, "comparison_matrix.json")
    check("comparison_matrix.json exists", os.path.isfile(cm_path))
    if os.path.isfile(cm_path):
        try:
            with open(cm_path) as f:
                cm = json.load(f)
            if isinstance(cm, list):
                ids = {_entry_id(e) for e in cm}
            elif isinstance(cm, dict):
                ids = set(cm.keys())
            else:
                ids = set()
            check("comparison_matrix has 4 target papers",
                  ids.issuperset(TARGET_IDS),
                  f"Found: {ids}")

            # CRITICAL: internal consistency — each target paper has integer
            # scores 1..5 on all five criteria and total == sum (NOT pinned values).
            entries = cm if isinstance(cm, list) else [
                dict(v, paper_id=k) if isinstance(v, dict) else {} for k, v in cm.items()
            ]
            problems = []
            seen = set()
            for e in entries:
                pid = _entry_id(e)
                if pid not in TARGET_IDS:
                    continue
                seen.add(pid)
                scores = e.get("scores", e)  # scores may be nested or flat
                # Case-insensitive key lookup: task.md does not pin the casing,
                # and the mandated Sheet headers use Novelty/Methodology_Rigor/...
                scores_lc = {str(k).lower(): v for k, v in scores.items()} if isinstance(scores, dict) else {}
                entry_lc = {str(k).lower(): v for k, v in e.items()} if isinstance(e, dict) else {}
                vals = {}
                for crit in CRITERIA:
                    v = scores_lc.get(crit)
                    if v is None:
                        v = entry_lc.get(crit)
                    vals[crit] = v
                # integer 1..5
                ok_range = all(
                    isinstance(vals[c], (int, float)) and float(vals[c]).is_integer()
                    and 1 <= int(vals[c]) <= 5 for c in CRITERIA)
                if not ok_range:
                    problems.append((pid, "scores not int 1..5", vals))
                    continue
                total = e.get("total", e.get("total_score", e.get("score")))
                expected_total = sum(int(vals[c]) for c in CRITERIA)
                if total is None or int(float(total)) != expected_total:
                    problems.append((pid, f"total {total} != sum {expected_total}", vals))
            check("comparison_matrix scores are integers 1..5 and total == sum of criteria",
                  seen == TARGET_IDS and len(problems) == 0,
                  f"seen={sorted(seen)} problems={problems}")
        except Exception as e:
            check("comparison_matrix readable", False, str(e))
            check("comparison_matrix scores are integers 1..5 and total == sum of criteria",
                  False, str(e))

    # final_rankings.json
    fr_path = os.path.join(agent_workspace, "final_rankings.json")
    fr_exists = os.path.isfile(fr_path)
    check("final_rankings.json exists", fr_exists)
    sorted_ok = False
    if fr_exists:
        try:
            with open(fr_path) as f:
                fr = json.load(f)
            if isinstance(fr, list):
                check("final_rankings has 4 entries", len(fr) >= 4, f"Got {len(fr)}")
                scores = []
                for e in fr:
                    s = e.get("total_score", e.get("total", e.get("score", 0)))
                    scores.append(float(s) if s else 0)
                sorted_ok = all(scores[i] >= scores[i + 1] for i in range(len(scores) - 1))
                # Check recommendations present
                recs = {str(e.get("recommendation", "")).lower() for e in fr}
                check("final_rankings has recommendations",
                      bool(recs.intersection({"accept", "revise", "reject"})),
                      f"Found: {recs}")
        except Exception as e:
            check("final_rankings readable", False, str(e))
    # CRITICAL: sorted descending (registers FAIL if missing/unsorted).
    check("final_rankings.json is sorted by total_score descending",
          sorted_ok,
          "final_rankings.json missing or not sorted descending")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_word(args.agent_workspace)
    check_json_files(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    print(f"\nOverall: {PASS_COUNT}/{total} ({accuracy:.1f}%)")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print(f"CRITICAL FAILURES ({len(critical_failed)}):")
        for n in critical_failed:
            print(f"  - {n}")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "critical_failed": critical_failed,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # Any critical failure => overall FAIL regardless of accuracy.
    if critical_failed:
        print("FAIL (critical check failed)")
        sys.exit(1)
    sys.exit(0 if accuracy >= 70 else 1)


if __name__ == "__main__":
    main()
