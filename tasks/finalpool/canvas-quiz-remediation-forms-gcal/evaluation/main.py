"""Evaluation for canvas-quiz-remediation-gform-gcal (RU stack: forms + canvas + gcal).

The agent (a Russian tutor of the "Основы финансов" course, course ID 16) must:
1. Recompute per-quiz average scores from the LIVE canvas data, find the quizzes
   whose average is below 80 (these need remediation).
2. Schedule one 1-hour 'Remediation Session: <quiz title>' Google Calendar event
   per underperforming quiz, starting at 15:00 on consecutive days beginning the
   Monday of the week AFTER --launch_time.
3. Create a Google Form titled 'Finance Quiz Self-Assessment' with exactly 4
   questions (unclear/difficult topics, study hours, confidence level, preferred
   study method) on the RU forms-mcp backend (schema gform.*).
4. Email remediation@financeou.example.com from tutor@financeou.example.com with
   subject containing 'Remediation Study Sessions' whose body contains the form
   link, the session count, all underperforming quiz titles, and the session dates.

The set of underperforming quizzes is NOT trusted from a constant — it is recomputed
from canvas.quizzes / canvas.quiz_submissions. The expected fixed identifiers below
are the stable seeded quiz codes; they are used to cross-check the recomputed set.

CRITICAL_CHECKS reflect substance: a single critical failure => overall FAIL
(sys.exit(1)) regardless of accuracy. The accuracy>=70 gate applies afterwards.
"""
import argparse
import datetime as dt
import json
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

COURSE_ID = 16
AVG_THRESHOLD = 80.0

# Stable seeded quiz codes expected to be underperforming (avg < 80). Used to
# cross-check the recomputed set, NOT as the sole source of truth.
EXPECTED_UNDERPERFORMING = [
    "CMA 34880",
    "CMA 34881",
    "CMA 34882",
    "CMA 34883",
    "CMA 34884",
]

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILS = []


def record(name, passed, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRITICAL " if critical else ""
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {tag}{name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {tag}{name}{msg}")
        if critical:
            CRITICAL_FAILS.append(name)


# ---------------------------------------------------------------------------
# Canvas: recompute underperforming quizzes from LIVE data
# ---------------------------------------------------------------------------
def compute_underperforming():
    """Return (list_of_quiz_titles_avg_below_80, debug_rows) from canvas live data."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT q.title,
               ROUND(AVG(100.0 * qs.score / NULLIF(q.points_possible, 0))::numeric, 2) AS avg_pct,
               COUNT(*) AS n
        FROM canvas.quizzes q
        JOIN canvas.quiz_submissions qs ON q.id = qs.quiz_id
        WHERE q.course_id = %s
        GROUP BY q.id, q.title
        ORDER BY avg_pct
        """,
        (COURSE_ID,),
    )
    rows = cur.fetchall()
    conn.close()
    under = [title for title, avg_pct, _n in rows if avg_pct is not None and float(avg_pct) < AVG_THRESHOLD]
    return under, rows


# ---------------------------------------------------------------------------
# Launch-time -> Monday of the following week
# ---------------------------------------------------------------------------
def monday_after_launch(launch_time):
    """Return date of the Monday of the week AFTER launch_time."""
    if not launch_time:
        return None
    s = launch_time.strip().replace("Z", "+00:00")
    try:
        base = dt.datetime.fromisoformat(s)
    except ValueError:
        # try date-only
        try:
            base = dt.datetime.strptime(s[:10], "%Y-%m-%d")
        except ValueError:
            return None
    d = base.date()
    # Monday of current week
    this_monday = d - dt.timedelta(days=d.weekday())
    return this_monday + dt.timedelta(days=7)


# ---------------------------------------------------------------------------
# Google Form (RU forms-mcp backend, schema gform.*)
# ---------------------------------------------------------------------------
def check_gform():
    print("\n=== Checking Google Form (gform.*) ===")
    crit = ("CRITICAL: форма 'Finance Quiz Self-Assessment' существует и содержит "
            "ровно 4 вопроса")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()

    record("Хотя бы одна форма создана", len(forms) >= 1, f"Found {len(forms)} forms")

    # Strict: exact title (case-insensitive) 'Finance Quiz Self-Assessment'.
    form_id = None
    for fid, title in forms:
        if (title or "").strip().lower() == "finance quiz self-assessment":
            form_id = fid
            break
    record("Форма с точным названием 'Finance Quiz Self-Assessment'",
           form_id is not None, f"Found titles: {[f[1] for f in forms]}")

    q_count = -1
    if form_id is not None:
        cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (form_id,))
        q_count = cur.fetchone()[0]
        record("Форма содержит ровно 4 вопроса", q_count == 4, f"Found {q_count} questions")
    conn.close()

    record(crit, form_id is not None and q_count == 4,
           f"form_found={form_id is not None} q_count={q_count}", critical=True)


# ---------------------------------------------------------------------------
# Google Calendar (kept-foreign infra; gcal.* data layer)
# ---------------------------------------------------------------------------
def check_gcal(under_titles, expected_monday):
    print("\n=== Checking Google Calendar (gcal.*) ===")
    crit_cover = ("CRITICAL: ровно по одной часовой сессии в формате "
                  "'Remediation Session: <quiz title>' для каждого отстающего теста, "
                  "начало в 15:00")
    crit_sched = ("CRITICAL: сессии запланированы на последовательные дни, начиная с "
                  "понедельника недели после launch_time")

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT id, summary, start_datetime, end_datetime FROM gcal.events ORDER BY start_datetime")
    events = cur.fetchall()
    conn.close()

    rem = [e for e in events if "remediation session" in (e[1] or "").lower()]
    record("Создано хотя бы одно событие пересдачи", len(rem) >= 1,
           f"Found {len(rem)} remediation events of {len(events)} total")

    n_under = len(under_titles)
    record(f"Количество сессий пересдач == числу отстающих тестов ({n_under})",
           len(rem) == n_under,
           f"events={len(rem)} expected={n_under}")

    # --- per-quiz coverage in the exact 'Remediation Session: <title>' format ---
    covered = []
    missing = []
    for qt in under_titles:
        want = f"remediation session: {qt}".lower()
        if any((e[1] or "").strip().lower() == want for e in rem):
            covered.append(qt)
        else:
            missing.append(qt)

    # 15:00 start + ~1h duration for every remediation event.
    times_ok = True
    durations_ok = True
    for _id, summ, sdt, edt in rem:
        if sdt is None:
            times_ok = False
            continue
        if sdt.hour != 15 or sdt.minute != 0:
            times_ok = False
        if edt is not None:
            dur_min = round((edt - sdt).total_seconds() / 60.0)
            if abs(dur_min - 60) > 1:
                durations_ok = False
        else:
            durations_ok = False

    record("Все сессии начинаются в 15:00", times_ok)
    record("Все сессии длятся 1 час", durations_ok)
    record(crit_cover,
           n_under > 0 and not missing and len(rem) == n_under and times_ok and durations_ok,
           f"covered={covered} missing={missing} times_ok={times_ok} dur_ok={durations_ok}",
           critical=True)

    # --- consecutive days starting Monday-after-launch ---
    sched_ok = False
    detail = ""
    if expected_monday is None:
        detail = "launch_time not provided / unparseable"
    elif not rem:
        detail = "no remediation events"
    else:
        days = sorted({e[2].date() for e in rem if e[2] is not None})
        expected_days = [expected_monday + dt.timedelta(days=i) for i in range(len(rem))]
        sched_ok = (days == expected_days)
        detail = f"actual_days={[str(d) for d in days]} expected={[str(d) for d in expected_days]}"
    record(crit_sched, sched_ok, detail, critical=True)


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------
def check_emails(under_titles, under_dates):
    print("\n=== Checking Emails ===")
    crit = ("CRITICAL: письмо на remediation@financeou.example.com от "
            "tutor@financeou.example.com с темой 'Remediation Study Sessions', "
            "в теле — ссылка на форму, количество сессий, названия всех тестов и даты")

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()
    conn.close()

    def recipients_of(to_addr):
        if not to_addr:
            return []
        if isinstance(to_addr, list):
            return [str(r).strip().lower() for r in to_addr]
        try:
            parsed = json.loads(to_addr)
            if isinstance(parsed, list):
                return [str(r).strip().lower() for r in parsed]
        except (json.JSONDecodeError, TypeError):
            pass
        return [str(to_addr).strip().lower()]

    result = None
    for subj, from_addr, to_addr, body in all_emails:
        if "remediation@financeou.example.com" in recipients_of(to_addr):
            result = (subj, from_addr, to_addr, body)
            break

    record("Письмо отправлено на remediation@financeou.example.com", result is not None,
           f"Total emails: {len(all_emails)}")

    subj_ok = from_ok = link_ok = count_ok = titles_ok = dates_ok = False
    if result:
        subj, from_addr, to_addr, body = result
        body_lower = (body or "").lower()
        subj_ok = "remediation study sessions" in (subj or "").lower()
        from_ok = "tutor@financeou.example.com" in (from_addr or "").lower()

        # Form link substring (the forms MCP produces URLs containing 'form').
        link_ok = ("http" in body_lower and "form" in body_lower)

        # Session count present (digit or RU words). Accept the numeric count of
        # underperforming quizzes anywhere in the body.
        n = len(under_titles)
        count_ok = (str(n) in (body or "")) or any(
            w in body_lower for w in ["сесси", "пересдач", "ремедиаци"]
        ) and str(n) in (body or "")
        # be lenient: count is satisfied if the numeric session count appears.
        count_ok = str(n) in (body or "")

        # All underperforming quiz titles present (CMA codes are exact English).
        titles_ok = all((qt.lower() in body_lower) for qt in under_titles) if under_titles else False

        # All scheduled session dates present. task.md prescribes no date format,
        # so accept numeric (ISO / d.m.yyyy RU) AND month-name forms in EN and RU.
        EN_MONTHS = ["", "January", "February", "March", "April", "May", "June",
                     "July", "August", "September", "October", "November", "December"]
        RU_MONTHS = ["", "января", "февраля", "марта", "апреля", "мая", "июня",
                     "июля", "августа", "сентября", "октября", "ноября", "декабря"]

        def date_present(d):
            iso = d.strftime("%Y-%m-%d")
            ru = f"{d.day:02d}.{d.month:02d}.{d.year}"
            ru2 = f"{d.day}.{d.month}.{d.year}"
            blw = body_lower
            # English long forms: "June 15, 2026" and "June 15 2026"
            en1 = f"{EN_MONTHS[d.month]} {d.day}, {d.year}".lower()
            en2 = f"{EN_MONTHS[d.month]} {d.day} {d.year}".lower()
            # Russian long form: "15 июня 2026" (optional leading zero)
            ru_long = f"{d.day} {RU_MONTHS[d.month]} {d.year}".lower()
            ru_long0 = f"{d.day:02d} {RU_MONTHS[d.month]} {d.year}".lower()
            return (
                iso in (body or "") or ru in (body or "") or ru2 in (body or "")
                or en1 in blw or en2 in blw or ru_long in blw or ru_long0 in blw
            )
        dates_ok = all(date_present(d) for d in under_dates) if under_dates else False

        record("Тема письма содержит 'Remediation Study Sessions'", subj_ok, f"Subject: {subj}")
        record("Отправитель tutor@financeou.example.com", from_ok, f"From: {from_addr}")
        record("В теле письма есть ссылка на форму", link_ok)
        record("В теле письма указано количество сессий", count_ok)
        record("В теле письма перечислены все названия отстающих тестов", titles_ok)
        record("В теле письма указаны все даты сессий", dates_ok)

    record(crit,
           result is not None and subj_ok and from_ok and link_ok and count_ok and titles_ok and dates_ok,
           "", critical=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("CANVAS QUIZ REMEDIATION GFORM GCAL - EVALUATION")
    print("=" * 70)

    # Recompute underperforming quizzes from live canvas data.
    try:
        under_titles, quiz_rows = compute_underperforming()
        print(f"\n[canvas] per-quiz avg%: {[(t, str(a), n) for t, a, n in quiz_rows]}")
        print(f"[canvas] underperforming (avg<80): {under_titles}")
    except Exception as e:
        print(f"[canvas] WARNING could not recompute averages ({e}); "
              f"falling back to expected codes")
        under_titles = list(EXPECTED_UNDERPERFORMING)
        quiz_rows = []

    # Sanity cross-check against the stable seeded codes (non-critical).
    matches_expected = sorted(under_titles) == sorted(EXPECTED_UNDERPERFORMING)
    record("Набор отстающих тестов совпадает с ожидаемым (CMA 34880..34884)",
           matches_expected, f"recomputed={under_titles} expected={EXPECTED_UNDERPERFORMING}")
    if not under_titles:
        under_titles = list(EXPECTED_UNDERPERFORMING)

    expected_monday = monday_after_launch(args.launch_time)
    under_dates = []
    if expected_monday is not None:
        under_dates = [expected_monday + dt.timedelta(days=i) for i in range(len(under_titles))]

    check_gform()
    check_gcal(under_titles, expected_monday)
    check_emails(under_titles, under_dates)

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)
    accuracy = PASS_COUNT / total * 100.0
    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")
    print(f"  Critical failures: {CRITICAL_FAILS}")

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({
                "total_passed": PASS_COUNT,
                "total_checks": total,
                "accuracy": accuracy,
                "critical_failures": CRITICAL_FAILS,
            }, f, indent=2)

    if CRITICAL_FAILS:
        print(f"\nFAIL: critical checks failed: {CRITICAL_FAILS}")
        sys.exit(1)
    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    print("FAIL")
    sys.exit(1)


if __name__ == "__main__":
    main()
