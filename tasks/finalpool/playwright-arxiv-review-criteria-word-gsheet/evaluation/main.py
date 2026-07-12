"""Evaluation for playwright-arxiv-review-criteria-word-gsheet (RU localized)."""
import argparse
import os
import re
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")

# Recommendation labels accepted bilingually (EN canonical + RU gloss).
RECO_ALIASES = {
    "accept": ["accept", "принять", "принима"],
    "weak accept": ["weak accept", "слабое принятие", "слабое принят"],
}

# Section heading keywords broadened to RU+EN. Each entry is a list of
# acceptable substrings (lowercased); presence of ANY satisfies the section.
SECTION_KEYS = {
    "summary": ["summary", "краткое описание", "кратко", "резюме"],
    "technical": ["technical soundness", "technical", "техническ", "обоснованност"],
    "novelty": ["novelty", "новизна"],
    "clarity": ["clarity", "ясност"],
}


def any_in(text, options):
    return any(o in text for o in options)


def score_near_label(text, label_options, score):
    """Return True if `score` (e.g. 5) appears near any of the label words.

    Accepts flexible formats: '5', '5/5', '5 из 5', 'оценка: 5', etc.
    We search a window of text after each label occurrence for a standalone
    digit equal to `score` (optionally followed by '/5' or 'из 5').
    """
    s = str(score)
    # A standalone integer score: not part of a longer number and NOT the
    # integer part of a decimal like '4.3' (that's an average, not a criterion
    # score). Allows '5', '5/5', '5 из 5', 'оценка: 5'.
    pat = re.compile(r"(?<!\d)" + re.escape(s) + r"(?!\d)(?!\.\d)")
    for lab in label_options:
        idx = 0
        while True:
            pos = text.find(lab, idx)
            if pos == -1:
                break
            # window = label + remainder of its line (stop at next newline) +
            # a short tail; this keeps us within the same criterion block and
            # avoids bleeding into the next section / the average line.
            # window = label line + the following line (justification), so a
            # score written either inline ('Технич...: 5') or in the next
            # paragraph is caught, but we never reach the section after that
            # (which could hold the average like '4.3').
            start = pos
            nl1 = text.find("\n", pos)
            if nl1 == -1:
                end = len(text)
            else:
                nl2 = text.find("\n", nl1 + 1)
                end = nl2 if nl2 != -1 else len(text)
            window = text[start:end]
            if pat.search(window):
                return True
            idx = pos + len(lab)
    return False


def check_word_review(agent_workspace, filename, expected_title_fragment,
                      expected_tech, expected_novelty, expected_clarity,
                      expected_recommendation, errors_out, critical_out):
    """Validate one review docx.

    Non-critical: file/section presence (structure).
    Critical: correct per-rubric numeric scores near their labels + correct
    recommendation label (RU or EN).
    """
    path = os.path.join(agent_workspace, filename)
    if not os.path.exists(path):
        errors_out.append(f"{filename} not found")
        critical_out.append(f"{filename} not found")
        return
    try:
        from docx import Document
        doc = Document(path)
        full_text = "\n".join(p.text for p in doc.paragraphs).lower()

        # --- structural (non-critical) ---
        if expected_title_fragment.lower() not in full_text:
            errors_out.append(f"{filename}: missing paper title fragment '{expected_title_fragment}'")
        for name, keys in SECTION_KEYS.items():
            if not any_in(full_text, keys):
                errors_out.append(f"{filename}: missing {name} section")

        # --- recommendation (critical) ---
        reco_key = expected_recommendation.lower()
        aliases = RECO_ALIASES.get(reco_key, [reco_key])
        if not any_in(full_text, aliases):
            msg = f"{filename}: missing/wrong recommendation '{expected_recommendation}' (RU+EN)"
            errors_out.append(msg)
            critical_out.append(msg)
        # For 'accept' make sure it isn't only matched inside 'weak accept' when
        # we expect a plain Accept (and vice versa). Disambiguate.
        if reco_key == "accept":
            # must have an 'accept'/'принять' that is NOT part of weak accept
            plain = bool(re.search(r"(?<!weak )accept", full_text)) or \
                    bool(re.search(r"(?<!слабое )принят", full_text)) or \
                    ("принять" in full_text)
            if not plain:
                msg = f"{filename}: expected plain Accept but only Weak Accept found"
                errors_out.append(msg)
                critical_out.append(msg)
        if reco_key == "weak accept":
            if not (("weak accept" in full_text) or ("слабое принят" in full_text)):
                msg = f"{filename}: expected Weak Accept recommendation"
                errors_out.append(msg)
                critical_out.append(msg)

        # --- scores near labels (critical) ---
        if not score_near_label(full_text, SECTION_KEYS["technical"], expected_tech):
            msg = f"{filename}: Technical Soundness score {expected_tech} not found near label"
            errors_out.append(msg)
            critical_out.append(msg)
        if not score_near_label(full_text, SECTION_KEYS["novelty"], expected_novelty):
            msg = f"{filename}: Novelty score {expected_novelty} not found near label"
            errors_out.append(msg)
            critical_out.append(msg)
        if not score_near_label(full_text, SECTION_KEYS["clarity"], expected_clarity):
            msg = f"{filename}: Clarity score {expected_clarity} not found near label"
            errors_out.append(msg)
            critical_out.append(msg)

    except Exception as e:
        errors_out.append(f"Error reading {filename}: {e}")
        critical_out.append(f"Error reading {filename}: {e}")


def check_gsheet(errors_out, critical_out):
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("""
            SELECT s.id, s.title FROM gsheet.spreadsheets s
            WHERE LOWER(s.title) LIKE '%review%' OR LOWER(s.title) LIKE '%conference%'
            ORDER BY s.id DESC LIMIT 5
        """)
        spreadsheets = cur.fetchall()
        if not spreadsheets:
            msg = "No review tracker spreadsheet found"
            errors_out.append(msg); critical_out.append(msg)
            cur.close(); conn.close(); return

        ss_id = spreadsheets[0][0]

        cur.execute("""
            SELECT id FROM gsheet.sheets
            WHERE spreadsheet_id = %s AND LOWER(title) LIKE '%%review%%'
            LIMIT 1
        """, (ss_id,))
        sheet_row = cur.fetchone()
        if not sheet_row:
            msg = "No 'Reviews' sheet found in spreadsheet"
            errors_out.append(msg); critical_out.append(msg)
            cur.close(); conn.close(); return
        sheet_id = sheet_row[0]

        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
            ORDER BY row_index, col_index
        """, (ss_id, sheet_id))
        cells = cur.fetchall()
        cur.close(); conn.close()

        if len(cells) < 24:  # header + 3 rows * 8 cols = 32
            errors_out.append(f"Too few cells in Reviews sheet: {len(cells)}, expected ~32")

        # Build per-row text blobs (row_index -> concatenated lowercase values).
        rows = {}
        for ri, ci, val in cells:
            rows.setdefault(ri, {})[ci] = str(val).lower() if val is not None else ""
        row_text = {ri: " ".join(d[c] for c in sorted(d)) for ri, d in rows.items()}

        all_text = " ".join(row_text.values())

        # Paper presence (non-critical, OR id/keyword).
        if "2301.07041" not in all_text and "scaling" not in all_text:
            errors_out.append("Scaling Laws paper not found in GSheet")
        if "2203.11171" not in all_text and "instruct" not in all_text:
            errors_out.append("InstructGPT paper not found in GSheet")
        if "2205.01068" not in all_text and "opt" not in all_text:
            errors_out.append("OPT paper not found in GSheet")

        # Status present (non-critical), RU+EN.
        if not any(k in all_text for k in ("completed", "выполнено", "завершено", "завершен")):
            errors_out.append("Review status 'Completed'/'Выполнено' not found in GSheet")

        # ---- CRITICAL: row-anchored averages by arxiv ID ----
        # Each arxiv ID must appear in a row whose blob also contains the
        # correct average score.
        expected = {
            "2301.07041": "4.3",   # Scaling Laws
            "2203.11171": "4.7",   # InstructGPT
            "2205.01068": "4.0",   # OPT
        }
        avg_row_index = {}  # arxiv_id -> row_index (for ordering check)
        for arxiv_id, avg in expected.items():
            found = False
            for ri, txt in row_text.items():
                if arxiv_id in txt:
                    # the average must be in the SAME row, as a standalone token
                    if re.search(r"(?<!\d)" + re.escape(avg) + r"(?!\d)", txt):
                        found = True
                        avg_row_index[arxiv_id] = ri
                        break
            if not found:
                msg = (f"GSheet: row for {arxiv_id} missing correct Average_Score "
                       f"{avg} (row-anchored)")
                errors_out.append(msg); critical_out.append(msg)

        # ---- CRITICAL: status Completed per row + descending sort ----
        for arxiv_id in expected:
            ri = None
            for r, txt in row_text.items():
                if arxiv_id in txt:
                    ri = r; break
            if ri is not None:
                txt = row_text[ri]
                if not any(k in txt for k in ("completed", "выполнено", "завершено", "завершен")):
                    msg = f"GSheet: row for {arxiv_id} not marked Completed/Выполнено"
                    errors_out.append(msg); critical_out.append(msg)

        # Descending order by Average_Score: InstructGPT(4.7) < Scaling(4.3) < OPT(4.0)
        # i.e. the row index of InstructGPT should be above (smaller) Scaling, etc.
        if all(k in avg_row_index for k in expected):
            ig = avg_row_index["2203.11171"]
            sl = avg_row_index["2301.07041"]
            op = avg_row_index["2205.01068"]
            if not (ig < sl < op):
                msg = ("GSheet: rows not sorted by Average_Score descending "
                       "(InstructGPT 4.7 > Scaling 4.3 > OPT 4.0)")
                errors_out.append(msg); critical_out.append(msg)

    except Exception as e:
        errors_out.append(f"Error checking GSheet: {e}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    agent_ws = args.agent_workspace or os.path.join(
        os.path.dirname(__file__), "..", "groundtruth_workspace")

    all_errors = []
    critical_errors = []

    checks = [
        ("Review_Scaling_Laws.docx", "Scaling Laws", 5, 4, 4, "Accept"),
        ("Review_InstructGPT.docx", "follow instructions", 5, 5, 4, "Accept"),
        ("Review_OPT.docx", "OPT", 4, 3, 5, "Weak Accept"),
    ]
    for fn, frag, t, n, c, reco in checks:
        print(f"  Checking {fn}...")
        before = len(all_errors)
        check_word_review(agent_ws, fn, frag, t, n, c, reco,
                          all_errors, critical_errors)
        new = all_errors[before:]
        if new:
            for e in new[:3]:
                print(f"    ERROR: {e}")
        else:
            print("    PASS")

    print("  Checking Google Sheet...")
    before = len(all_errors)
    check_gsheet(all_errors, critical_errors)
    new = all_errors[before:]
    if new:
        for e in new[:3]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    # ---- Critical gate: any critical failure => FAIL regardless of accuracy ----
    if critical_errors:
        print(f"\n=== CRITICAL FAILURE ({len(critical_errors)}) ===")
        for e in critical_errors[:10]:
            print(f"  CRITICAL: {e}")
        print("=== RESULT: FAIL ===")
        sys.exit(1)

    # ---- Accuracy gate (>=70%) over all (non-critical) checks ----
    # Total checks: 3 docs * (1 title + 4 sections + 1 reco + 3 scores) = 27
    #             + gsheet (3 presence + 1 status + 3 avg + 3 row-status + 1 sort) = 11
    TOTAL = 38
    failed = len(all_errors)
    passed = max(0, TOTAL - failed)
    accuracy = 100.0 * passed / TOTAL
    print(f"\nAccuracy: {accuracy:.1f}% ({passed}/{TOTAL}), failures={failed}")

    if accuracy >= 70.0:
        print("=== RESULT: PASS ===")
        sys.exit(0)
    else:
        print(f"=== RESULT: FAIL ({failed} errors) ===")
        for e in all_errors[:10]:
            print(f"  {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
