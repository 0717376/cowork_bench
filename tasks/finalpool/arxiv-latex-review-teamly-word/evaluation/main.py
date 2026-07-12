"""
Evaluation for arxiv-latex-review task (Teamly variant).
Checks: Teamly knowledge-base page, Word document, Google Sheet.

The agent writes Russian prose for descriptions/results, so keyword matching
accepts both English and Russian forms where appropriate. Paper titles, arxiv
IDs, author names and column headers stay English (real arxiv identifiers).

Critical checks (see CRITICAL_CHECKS): any failure there => overall FAIL
regardless of accuracy. Otherwise pass threshold: accuracy >= 70%.

The semantic substance of the task is: include the 3 relevant LLM papers and
EXCLUDE the robotics noise paper (2309.16349). Both are enforced critically.
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

# Relevant papers (must appear) and the noise paper (must NOT appear).
RELEVANT_IDS = ["2305.20050", "2307.09288", "2310.06825"]
RELEVANT_DATES = ["2023-05-30", "2023-07-18", "2023-10-10"]
NOISE_TOKENS = ["affordance", "robot learning", "2309.16349"]
REQUIRED_HEADERS = ["arxiv_id", "title", "authors", "published_date",
                    "key_contribution", "method_category"]

# Critical checks: any failure => overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "Teamly KB page exists and covers at least 2 relevant methods",
    "Robotics noise paper is correctly EXCLUDED from all deliverables",
    "GSheet has the 6 required column headers and the 3 relevant arxiv IDs",
    "Word doc names all 3 relevant papers plus a correct author",
    "All 3 published dates appear in the GSheet",
}


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


# Collected deliverable text for the cross-cutting noise-exclusion check.
_NOISE_CORPUS = {"teamly": "", "word": "", "gsheet": ""}


def check_teamly():
    print("\n=== Checking Teamly Knowledge Base Page ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT title, COALESCE(body, '') FROM teamly.pages")
        rows = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Teamly KB page exists and covers at least 2 relevant methods",
               False, f"Query failed: {e}")
        record("Teamly page has content for at least 2 relevant papers",
               False, f"Query failed: {e}")
        return

    all_text = " ".join((str(t) + " " + str(b)).lower() for t, b in rows)
    _NOISE_CORPUS["teamly"] = all_text

    # Target KB page exists (title marker).
    has_kb_title = any(
        ("knowledge base" in (t or "").lower())
        or ("llm" in (t or "").lower() and "fine" in (t or "").lower())
        for t, _ in rows
    )
    record("Teamly KB page 'LLM Fine-Tuning Knowledge Base' exists",
           has_kb_title, f"Total pages: {len(rows)}")

    # Method coverage (English identifiers; RU prose allowed around them).
    has_dpo = "dpo" in all_text or "direct preference" in all_text
    has_llama = "llama" in all_text
    has_mistral = "mistral" in all_text
    methods = sum([has_dpo, has_llama, has_mistral])

    # CRITICAL: page exists AND covers >= 2 of the 3 relevant methods.
    record("Teamly KB page exists and covers at least 2 relevant methods",
           has_kb_title and methods >= 2,
           f"kb_title={has_kb_title}, dpo={has_dpo}, llama={has_llama}, mistral={has_mistral}")

    record("Teamly page has content for at least 2 relevant papers",
           methods >= 2,
           f"dpo={has_dpo}, llama={has_llama}, mistral={has_mistral}")


def check_word(agent_workspace):
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "LLM_Paper_Synthesis.docx")
    if not os.path.isfile(doc_path):
        record("Word file LLM_Paper_Synthesis.docx exists", False, f"Not found at: {doc_path}")
        record("Word doc names all 3 relevant papers plus a correct author", False,
               "Word doc missing")
        return
    record("Word file LLM_Paper_Synthesis.docx exists", True)

    try:
        doc = Document(doc_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        record("Word doc names all 3 relevant papers plus a correct author", False, str(e))
        return
    record("Word file readable", True)

    full_text = "\n".join(p.text for p in doc.paragraphs).lower()
    _NOISE_CORPUS["word"] = full_text

    has_heading = ("llm" in full_text or "fine-tun" in full_text) and \
        ("survey" in full_text or "synthesis" in full_text or "alignment" in full_text)
    record("Word has heading mentioning LLM fine-tuning/alignment", has_heading)

    has_intro = len(full_text) > 400
    record("Word has substantial content", has_intro, f"Text length: {len(full_text)}")

    has_dpo = "dpo" in full_text or "direct preference" in full_text
    has_llama = "llama" in full_text
    has_mistral = "mistral" in full_text
    papers_mentioned = sum([has_dpo, has_llama, has_mistral])

    record("Word mentions DPO paper", has_dpo)
    record("Word mentions Llama 2 paper", has_llama)
    record("Word mentions Mistral 7B paper", has_mistral)

    # At least one correct author name across the 3 relevant papers.
    authors = ["rafailov", "archit sharma", "touvron", "louis martin",
               "albert jiang", "sablayrolles"]
    has_author = any(a in full_text for a in authors)
    record("Word mentions at least one correct author name", has_author)

    # CRITICAL: all 3 relevant papers named + at least one correct author.
    record("Word doc names all 3 relevant papers plus a correct author",
           papers_mentioned >= 3 and has_author,
           f"papers={papers_mentioned}/3, author={has_author}")


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        target_ss = None
        for sid, title in spreadsheets:
            t = (title or "").lower()
            if ("llm" in t or "paper" in t) and ("registry" in t or "paper" in t):
                target_ss = sid
                break

        record("GSheet 'LLM Paper Registry' exists",
               target_ss is not None,
               f"Found sheets: {[t for _, t in spreadsheets]}")

        if target_ss is None:
            record("GSheet has the 6 required column headers and the 3 relevant arxiv IDs",
                   False, "Spreadsheet not found")
            record("All 3 published dates appear in the GSheet", False, "Spreadsheet not found")
            conn.close()
            return

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (target_ss,))
        sheets = cur.fetchall()
        if not sheets:
            record("GSheet has at least one sheet", False)
            record("GSheet has the 6 required column headers and the 3 relevant arxiv IDs",
                   False, "No sheet")
            record("All 3 published dates appear in the GSheet", False, "No sheet")
            conn.close()
            return

        sheet_id = sheets[0][0]
        cur.execute("""
            SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 0
        """, (target_ss, sheet_id))
        data_rows = cur.fetchone()[0]
        record("GSheet 'LLM Paper Registry' has at least 3 data rows",
               data_rows >= 3, f"Found {data_rows} data rows")

        # Header row (row 0).
        cur.execute("""
            SELECT LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index = 0
        """, (target_ss, sheet_id))
        header_text = " ".join(r[0] for r in cur.fetchall() if r[0])
        missing_headers = [h for h in REQUIRED_HEADERS if h not in header_text]
        record("GSheet header row has the 6 required column headers",
               len(missing_headers) == 0,
               f"Missing: {missing_headers}")

        # All cells.
        cur.execute("""
            SELECT LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
        """, (target_ss, sheet_id))
        cell_values = [row[0] for row in cur.fetchall() if row[0]]
        all_text = " ".join(cell_values)
        _NOISE_CORPUS["gsheet"] = all_text

        has_dpo = "dpo" in all_text or "direct preference" in all_text
        has_llama = "llama" in all_text
        has_mistral = "mistral" in all_text
        record("GSheet contains DPO paper entry", has_dpo)
        record("GSheet contains Llama paper entry", has_llama)
        record("GSheet contains Mistral paper entry", has_mistral)

        # CRITICAL: 6 headers present AND all 3 relevant arxiv IDs present.
        ids_present = [aid for aid in RELEVANT_IDS if aid in all_text]
        record("GSheet has the 6 required column headers and the 3 relevant arxiv IDs",
               len(missing_headers) == 0 and len(ids_present) == 3,
               f"missing_headers={missing_headers}, ids_present={ids_present}")

        # CRITICAL: all 3 published dates present.
        dates_present = [d for d in RELEVANT_DATES if d in all_text]
        record("All 3 published dates appear in the GSheet",
               len(dates_present) == 3,
               f"dates_present={dates_present}")

        conn.close()
    except Exception as e:
        record("GSheet connection", False, str(e))


def check_noise_exclusion():
    print("\n=== Checking Robotics Noise Paper Exclusion ===")
    corpus = " ".join(_NOISE_CORPUS.values())
    leaked = [tok for tok in NOISE_TOKENS if tok in corpus]
    record("Robotics noise paper is correctly EXCLUDED from all deliverables",
           len(leaked) == 0,
           f"Leaked tokens: {leaked}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_teamly()
    check_word(args.agent_workspace)
    check_gsheet()
    check_noise_exclusion()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print(f"CRITICAL FAILURES: {len(critical_failed)}")
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

    if critical_failed:
        print("FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
