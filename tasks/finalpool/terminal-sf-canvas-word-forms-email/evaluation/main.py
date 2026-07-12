"""Evaluation for terminal-sf-canvas-word-gform-email.
Checks:
1. Training_Effectiveness_Report.docx content
2. Google Form "Training Feedback Survey" with 5 questions
3. Emails to hr_director and training_team
4. Script files exist (training_matches.py, effectiveness_analysis.py, survey_analysis.py)
5. JSON output files exist
"""
import argparse
import json
import os
import sys

import psycopg2
from docx import Document

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRITICAL " if critical else ""
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {tag}{name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {tag}{name}: {str(detail)[:200]}")
        if critical:
            CRITICAL_FAILED.append(name)


def num_close(a, b, tol=2.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_values():
    """Query DB for expected values."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Course avg scores
    cur.execute("""
        SELECT a.course_id, ROUND(AVG(s.score)::numeric, 2)
        FROM canvas.assignments a
        JOIN canvas.submissions s ON s.assignment_id = a.id
        WHERE a.course_id IN (9, 10) AND s.score IS NOT NULL
        GROUP BY a.course_id
    """)
    course_avgs = {int(r[0]): float(r[1]) for r in cur.fetchall()}

    # Enrollment counts
    cur.execute("""
        SELECT course_id, COUNT(DISTINCT user_id)
        FROM canvas.enrollments
        WHERE course_id IN (9, 10) AND type='StudentEnrollment'
        GROUP BY course_id
    """)
    enrollments = {int(r[0]): r[1] for r in cur.fetchall()}

    # SF dept ratings
    cur.execute("""
        SELECT "DEPARTMENT", ROUND(AVG("PERFORMANCE_RATING")::numeric, 2)
        FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
        WHERE "DEPARTMENT" IN ('Engineering', 'R&D')
        GROUP BY "DEPARTMENT"
    """)
    dept_ratings = {r[0]: float(r[1]) for r in cur.fetchall()}

    cur.close()
    conn.close()

    c9_avg = course_avgs.get(9, 69.59)
    c10_avg = course_avgs.get(10, 71.53)
    eng_rating = dept_ratings.get("Engineering", 3.21)
    rnd_rating = dept_ratings.get("R&D", 3.20)
    eng_impr = eng_rating - 3.00
    rnd_impr = rnd_rating - 2.95
    avg_impr = (eng_impr + rnd_impr) / 2
    overall_avg_score = (c9_avg + c10_avg) / 2

    return {
        "c9_avg": c9_avg, "c10_avg": c10_avg,
        "eng_enrolled": enrollments.get(9, 1938),
        "rnd_enrolled": enrollments.get(10, 1803),
        "eng_rating": eng_rating, "rnd_rating": rnd_rating,
        "eng_impr": eng_impr, "rnd_impr": rnd_impr,
        "avg_impr": avg_impr,
        "overall_avg_score": overall_avg_score,
    }


def check_word(workspace):
    """Check Training_Effectiveness_Report.docx."""
    print("\n=== Check 1: Word Document ===")
    path = os.path.join(workspace, "Training_Effectiveness_Report.docx")
    if not os.path.exists(path):
        check("Word document exists", False, f"Not found: {path}")
        return
    check("Word document exists", True)

    doc = Document(path)
    full_text = " ".join(p.text for p in doc.paragraphs).lower()

    check("Has Executive Summary section",
          "executive summary" in full_text)
    check("Has Methodology section",
          "methodology" in full_text)
    check("Has Performance Impact section",
          "performance impact" in full_text or "impact analysis" in full_text)
    check("Has Survey Findings section",
          "survey findings" in full_text or "survey" in full_text)
    check("Has ROI section",
          "roi" in full_text or "return on investment" in full_text)
    check("Has Recommendations section",
          "recommendation" in full_text)
    check("Mentions Engineering department",
          "engineering" in full_text)
    check("Mentions R&D department",
          "r&d" in full_text or "r & d" in full_text)
    check("Mentions Проектирование на основе данных",
          "data-driven design" in full_text or "data driven design" in full_text)
    check("Has substantial content", len(full_text) > 500, f"Length: {len(full_text)}")

    ev = get_expected_values()

    # CRITICAL: key live numbers (course avg scores) must appear, not just any number.
    check("Mentions course 9 avg score",
          str(round(ev["c9_avg"], 1)) in full_text or str(round(ev["c9_avg"], 2)) in full_text
          or str(int(round(ev["c9_avg"]))) in full_text,
          f"Expected ~{ev['c9_avg']:.2f}", critical=True)
    check("Mentions course 10 avg score",
          str(round(ev["c10_avg"], 1)) in full_text or str(round(ev["c10_avg"], 2)) in full_text
          or str(int(round(ev["c10_avg"]))) in full_text,
          f"Expected ~{ev['c10_avg']:.2f}", critical=True)

    # CRITICAL: conditional recommendation must match the live avg improvement branch.
    # RU + EN keywords (agent body is Russian).
    if ev["avg_impr"] < 0.15:
        check("Recommends restructuring (improvement < 0.15)",
              "restructur" in full_text or "реструктур" in full_text or "переработ" in full_text
              or "пересмотр" in full_text,
              f"Avg improvement: {ev['avg_impr']:.2f}", critical=True)
    else:
        check("Recommends expanding (improvement >= 0.15)",
              "expand" in full_text or "расшир" in full_text or "масштаб" in full_text,
              f"Avg improvement: {ev['avg_impr']:.2f}", critical=True)


def check_gform():
    """Check Google Form creation."""
    print("\n=== Check 2: Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()

    form_id = None
    for fid, title in forms:
        t = (title or "").lower()
        if "training" in t and ("feedback" in t or "survey" in t):
            form_id = fid
            break

    check("Training Feedback Survey form exists", form_id is not None,
          f"Forms: {[f[1] for f in forms]}", critical=True)

    if form_id:
        cur.execute("SELECT title, question_type, config FROM gform.questions WHERE form_id = %s ORDER BY position",
                    (form_id,))
        questions = cur.fetchall()
        check("Form has exactly 5 questions", len(questions) == 5,
              f"Found {len(questions)}", critical=True)

        if len(questions) >= 5:
            q_titles = [q[0].lower() for q in questions]
            q_types = [(q[1] or "").upper() for q in questions]
            q_configs = [q[2] for q in questions]
            check("Q1 about satisfaction",
                  any("satisfaction" in t or "rating" in t for t in q_titles[:2]),
                  f"Q titles: {q_titles}", critical=True)
            check("Q4 about recommendation",
                  any("recommend" in t for t in q_titles),
                  f"Q titles: {q_titles}")
            check("Q5 about format",
                  any("format" in t for t in q_titles),
                  f"Q titles: {q_titles}")

            # CRITICAL: question types/choices reflect the spec (Q1 rating 1-5,
            # Q4 Yes/No, Q5 Online/In-Person/Hybrid). Accept by choice content so
            # the forms MCP's question-type naming does not break the check.
            def _choices(cfg):
                # forms MCP (add_multiple_choice_question) stores choice options as
                #   config = {"type": "RADIO", "options": [{"value": "..."}, ...]}
                # Accept that real schema, plus a couple of legacy shapes, so an
                # agent-created form passes.
                if not cfg:
                    return []
                try:
                    c = cfg if isinstance(cfg, dict) else json.loads(cfg)
                except Exception:
                    return []
                if not isinstance(c, dict):
                    return []
                out = []
                opts = c.get("options")
                if isinstance(opts, list):
                    for o in opts:
                        if isinstance(o, dict):
                            out.append(str(o.get("value", "")).lower())
                        else:
                            out.append(str(o).lower())
                for x in c.get("choices", []) or []:
                    out.append(str(x).lower())
                return [o for o in out if o]
            all_choices = [_choices(c) for c in q_configs]
            flat = [opt for ch in all_choices for opt in ch]
            q1_rating = any(set(["1", "2", "3", "4", "5"]).issubset(set(ch)) for ch in all_choices) \
                or q_types[0] in ("SCALE", "RATING", "LINEAR_SCALE")
            check("Q1 is a 1-5 rating", q1_rating,
                  f"Q1 type={q_types[0]} choices={all_choices[0] if all_choices else []}",
                  critical=True)
            yes_no = any(set(["yes", "no"]).issubset(set(ch)) for ch in all_choices)
            check("Q4 has Yes/No choices", yes_no,
                  f"choices={all_choices}", critical=True)
            fmt = any(("online" in ch and "hybrid" in ch
                       and any("person" in o for o in ch)) for ch in all_choices)
            check("Q5 has Online/In-Person/Hybrid choices", fmt,
                  f"choices={all_choices}", critical=True)

        # NOTE: survey response data lives in the workspace input file
        # (survey_responses.json), not in gform.responses, so the agent-created
        # deliverable form legitimately has no gform responses. survey_results.json
        # content (derived from those responses) is validated in check_json_outputs.

    cur.close()
    conn.close()


def check_emails():
    """Check emails to hr_director and training_team."""
    print("\n=== Check 3: Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()

    # Check email to hr_director
    hr_email = None
    training_email = None
    for subj, from_addr, to_addr, body in all_emails:
        to_str = str(to_addr).lower() if to_addr else ""
        if "hr_director" in to_str:
            hr_email = (subj, from_addr, to_addr, body)
        if "training_team" in to_str:
            training_email = (subj, from_addr, to_addr, body)

    check("Email sent to hr_director@company.com", hr_email is not None,
          f"Total emails: {len(all_emails)}", critical=True)
    if hr_email:
        subj, from_addr, to_addr, body = hr_email
        subj_l = (subj or "").lower()
        # Subject stays English per task.md.
        check("HR email subject mentions effectiveness/training",
              "effectiveness" in subj_l or "training" in subj_l,
              f"Subject: {subj}")
        check("HR email from training_analytics@company.com",
              "training_analytics" in (from_addr or "").lower(),
              f"From: {from_addr}", critical=True)
        body_lower = (body or "").lower()
        # Body is Russian: accept RU + EN keywords.
        perf_ok = ("performance" in body_lower or "rating" in body_lower
                   or "эффективност" in body_lower or "оценк" in body_lower
                   or "рейтинг" in body_lower)
        impr_ok = ("improvement" in body_lower or "baseline" in body_lower
                   or "improv" in body_lower or "улучшен" in body_lower
                   or "прирост" in body_lower or "базов" in body_lower)
        check("HR email mentions performance (RU/EN)", perf_ok,
              "Expected performance/rating in body")
        check("HR email mentions improvement or baseline (RU/EN)", impr_ok,
              "Expected improvement mention")
        # CRITICAL: subject names the analysis AND body covers per-dept perf improvement.
        check("HR email: subject 'Effectiveness' AND body covers perf improvement (RU/EN)",
              "effectiveness" in subj_l and perf_ok and impr_ok,
              f"subj={subj} perf={perf_ok} impr={impr_ok}", critical=True)

    check("Email sent to training_team@company.com", training_email is not None,
          f"Total emails: {len(all_emails)}", critical=True)
    if training_email:
        subj, from_addr, to_addr, body = training_email
        subj_l = (subj or "").lower()
        check("Training email subject mentions survey/feedback",
              "survey" in subj_l or "feedback" in subj_l,
              f"Subject: {subj}")
        body_lower = (body or "").lower()
        check("Training email mentions satisfaction (RU/EN)",
              "satisfaction" in body_lower or "rating" in body_lower
              or "удовлетвор" in body_lower or "оценк" in body_lower,
              "Expected satisfaction mention")
        check("Training email mentions format or recommend (RU/EN)",
              "format" in body_lower or "recommend" in body_lower
              or "формат" in body_lower or "рекоменд" in body_lower
              or "порекоменд" in body_lower,
              "Expected format/recommend mention")

    cur.close()
    conn.close()


def check_scripts(workspace):
    """Check that required scripts exist."""
    print("\n=== Check 4: Scripts ===")
    for script in ["training_matches.py", "effectiveness_analysis.py", "survey_analysis.py"]:
        path = os.path.join(workspace, script)
        check(f"{script} exists", os.path.exists(path), f"Not found: {path}")


def check_json_outputs(workspace):
    """Check JSON output files."""
    print("\n=== Check 5: JSON Outputs ===")
    for jfile in ["training_matches.json", "effectiveness_analysis.json", "survey_results.json"]:
        path = os.path.join(workspace, jfile)
        if not os.path.exists(path):
            check(f"{jfile} exists", False, f"Not found: {path}")
            continue
        check(f"{jfile} exists", True)
        try:
            with open(path) as f:
                data = json.load(f)
            check(f"{jfile} is valid JSON", True)
            check(f"{jfile} is non-empty", len(data) > 0, "Empty JSON")
        except json.JSONDecodeError as e:
            check(f"{jfile} is valid JSON", False, str(e))

    # Check effectiveness_analysis.json content
    ea_path = os.path.join(workspace, "effectiveness_analysis.json")
    if os.path.exists(ea_path):
        try:
            with open(ea_path) as f:
                ea = json.load(f)
            ea_str = json.dumps(ea).lower()
            check("effectiveness_analysis mentions Engineering or R&D",
                  "engineering" in ea_str or "r&d" in ea_str or "r_d" in ea_str,
                  "Expected department names")
            # CRITICAL: per-department completion_rate AND performance_gap present,
            # and the performance gap matches the live HR rating vs baseline.
            ev = get_expected_values()
            has_cr = "completion_rate" in ea_str or "completion rate" in ea_str
            has_gap = "performance_gap" in ea_str or "performance gap" in ea_str \
                or "gap" in ea_str

            def _find_nums(obj, out):
                if isinstance(obj, dict):
                    for v in obj.values():
                        _find_nums(v, out)
                elif isinstance(obj, list):
                    for v in obj:
                        _find_nums(v, out)
                elif isinstance(obj, (int, float)):
                    out.append(float(obj))

            nums = []
            _find_nums(ea, nums)
            # Live performance gaps per dept (current avg rating - baseline).
            gap_eng = ev["eng_impr"]
            gap_rnd = ev["rnd_impr"]
            gap_ok = (any(abs(n - gap_eng) <= 0.05 for n in nums)
                      and any(abs(n - gap_rnd) <= 0.05 for n in nums))
            check("effectiveness_analysis has completion_rate AND performance_gap "
                  "matching live HR gaps (Eng/R&D vs baselines 3.00/2.95)",
                  has_cr and has_gap and gap_ok,
                  f"has_cr={has_cr} has_gap={has_gap} gap_ok={gap_ok} "
                  f"expected gaps eng={gap_eng:.2f} rnd={gap_rnd:.2f}",
                  critical=True)
        except Exception as e:
            check("effectiveness_analysis content readable", False, str(e), critical=True)

    # Check survey_results.json content derived from the seeded survey responses
    # (survey_responses.json: 15 responses). Expected: avg satisfaction ~3.87,
    # 11 respondents recommended, most common preferred format = Hybrid.
    sr_path = os.path.join(workspace, "survey_results.json")
    if os.path.exists(sr_path):
        try:
            with open(sr_path) as f:
                sr = json.load(f)
            sr_str = json.dumps(sr).lower()

            def _find_nums2(obj, out):
                if isinstance(obj, dict):
                    for v in obj.values():
                        _find_nums2(v, out)
                elif isinstance(obj, list):
                    for v in obj:
                        _find_nums2(v, out)
                elif isinstance(obj, (int, float)):
                    out.append(float(obj))

            nums = []
            _find_nums2(sr, nums)
            avg_sat_ok = any(abs(n - 3.87) <= 0.1 for n in nums)
            recommend_ok = any(abs(n - 11) <= 0.5 for n in nums)
            fmt_ok = "hybrid" in sr_str
            check("survey_results reflects seeded responses "
                  "(avg satisfaction ~3.87, 11 recommend, format Hybrid)",
                  avg_sat_ok and recommend_ok and fmt_ok,
                  f"avg_sat_ok={avg_sat_ok} recommend_ok={recommend_ok} "
                  f"fmt_ok={fmt_ok}")
        except Exception as e:
            check("survey_results content readable", False, str(e))


def check_reverse_validation():
    """Verify noise data not misused."""
    print("\n=== Reverse Validation ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # No emails to noise recipients
    cur.execute("""
        SELECT to_addr FROM email.messages
        WHERE from_addr ILIKE '%%training_analytics%%'
    """)
    sent = cur.fetchall()
    noise_addrs = ["all@company.com", "managers@company.com", "leadership@company.com"]
    for row in sent:
        to_str = str(row[0]).lower()
        for noise in noise_addrs:
            if noise in to_str:
                check("No emails sent to noise recipients", False, f"Sent to {noise}")
                cur.close()
                conn.close()
                return
    check("No emails sent to noise recipients", True)

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("TERMINAL-SF-CANVAS-WORD-GFORM-EMAIL - EVALUATION")
    print("=" * 70)

    check_word(args.agent_workspace)
    check_gform()
    check_emails()
    check_scripts(args.agent_workspace)
    check_json_outputs(args.agent_workspace)
    check_reverse_validation()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    print(f"\nOverall: {PASS_COUNT}/{total} ({accuracy:.1f}%)")

    result = {"total_passed": PASS_COUNT, "total_checks": total, "accuracy": accuracy,
              "critical_failed": CRITICAL_FAILED}
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # Any CRITICAL failure => hard FAIL, independent of the accuracy gate.
    if CRITICAL_FAILED:
        print(f"\nCRITICAL checks failed ({len(CRITICAL_FAILED)}): {CRITICAL_FAILED}")
        sys.exit(1)

    sys.exit(0 if accuracy >= 70 else 1)


if __name__ == "__main__":
    main()
