"""Evaluation for insales-customer-survey-form (RU / InSales + Forms).

Critical gate model:
  * A small set of SEMANTIC checks (CRITICAL_CHECKS) must all pass. Any critical
    failure => sys.exit(1) immediately, regardless of the overall accuracy.
  * Remaining (structural / soft) checks contribute to an accuracy score; PASS
    additionally requires accuracy >= 70.

Forms backend reality (local_servers/forms-mcp): gform.questions.question_type is
only 'textQuestion' or 'choiceQuestion'. Choice options live in config JSONB as
{"type":"RADIO","options":[{"value": ...}]}. There is no native linear-scale or
paragraph type, so scale questions are represented as choice questions with the
numeric option lists 1..5 / 1..10, and free-text questions as textQuestion.

Customer first names in wc.customers are russified centrally (db/zzz_wc_after_init.sql);
the per-recipient personalization check reads the actual first_name for each email,
so it works for both EN and RU seeds without hand-translation here.
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

# Customers with at least one completed order (frozen groundtruth set).
EXPECTED_EMAILS = {
    "emily.johnson@x.dummyjson.com",
    "michael.williams@x.dummyjson.com",
    "sophia.brown@x.dummyjson.com",
    "olivia.wilson@x.dummyjson.com",
    "ava.taylor@x.dummyjson.com",
    "ethan.martinez@x.dummyjson.com",
    "liam.garcia@x.dummyjson.com",
    "mia.rodriguez@x.dummyjson.com",
    "noah.hernandez@x.dummyjson.com",
    "charlotte.lopez@x.dummyjson.com",
    "william.gonzalez@x.dummyjson.com",
    "evelyn.sanchez@x.dummyjson.com",
    "abigail.rivera@x.dummyjson.com",
    "chloe.morales@x.dummyjson.com",
    "mateo.nguyen@x.dummyjson.com",
    "evelyn.gonzalez@x.dummyjson.com",
    "daniel.cook@x.dummyjson.com",
    "lily.lee@x.dummyjson.com",
    "henry.hill@x.dummyjson.com",
    "addison.wright@x.dummyjson.com",
    "gabriel.adams@x.dummyjson.com",
    "natalie.harris@x.dummyjson.com",
    "carter.baker@x.dummyjson.com",
    "harper.turner@x.dummyjson.com",
    "mason.parker@x.dummyjson.com",
    "lucas.gray@x.dummyjson.com",
    "ella.adams@x.dummyjson.com",
    "evan.reed@x.dummyjson.com",
    "avery.carter@x.dummyjson.com",
    "benjamin.foster@x.dummyjson.com",
    "scarlett.wright@x.dummyjson.com",
    "lincoln.kelly@x.dummyjson.com",
    "hannah.robinson@x.dummyjson.com",
    "nicholas.bailey@x.dummyjson.com",
    "luna.russell@x.dummyjson.com",
    "stella.hughes@x.dummyjson.com",
}

# Per-position question title alternatives (EN + RU). Matched as substrings
# against the ORIGINAL question title (.lower()), NOT via any .split() hack.
# A position passes if ANY of its alternatives is contained in the title.
EXPECTED_QUESTIONS = [
    ["overall shopping experience", "опыт покуп", "общий опыт"],
    ["satisfied are you with the product quality", "качеств", "довольны качеств"],
    ["likely are you to recommend our store", "рекоменд", "порекоменд"],
    ["like most about your purchase", "понравил", "больше всего"],
    ["suggestions for improvement", "предложен", "улучшен"],
]

# Form title / description alternatives (EN + RU).
TITLE_ALTS = ["customer satisfaction survey", "опрос удовлетвор", "удовлетворённости клиентов",
              "удовлетворенности клиентов"]
DESC_ALTS = ["value your feedback", "ценим вашу обратную связь", "ценим ваш"]

# Email subject keywords (EN + RU).
SUBJECT_ALTS = ["feedback", "survey", "satisfaction", "отзыв", "опрос", "удовлетвор", "обратн"]
# Greeting keywords (EN + RU).
GREETING_ALTS = ["dear", "уважаем", "здравствуй", "добрый день", "привет"]

SURVEY_SENDER = "survey@store.com"


def str_contains_any(haystack, needles):
    if haystack is None:
        return False
    h = str(haystack).lower()
    return any(n.lower() in h for n in needles)


def config_options(cfg):
    """Return list of lowercased option value strings from a choice config JSONB."""
    if cfg is None:
        return []
    if isinstance(cfg, str):
        try:
            cfg = json.loads(cfg)
        except Exception:
            return []
    if not isinstance(cfg, dict):
        return []
    opts = cfg.get("options") or []
    out = []
    for o in opts:
        if isinstance(o, dict):
            v = o.get("value", o.get("label", ""))
        else:
            v = o
        out.append(str(v).strip().lower())
    return [v for v in out if v]


def numeric_options_cover(opts, lo, hi):
    """True if option values include every integer in [lo, hi]."""
    nums = set()
    for o in opts:
        m = re.search(r"-?\d+", o)
        if m:
            nums.add(int(m.group()))
    return all(n in nums for n in range(lo, hi + 1))


def build_email_to_first_name(cur):
    """email (lower) -> russified first_name (lower), from wc.customers.

    Prefers the canonical top-level `email` column; falls back to billing->>'email'.
    """
    mapping = {}
    try:
        cur.execute("SELECT email, first_name, billing FROM wc.customers")
        for email, first_name, billing in cur.fetchall():
            if not first_name:
                continue
            fn = str(first_name).strip().lower()
            if email:
                mapping[str(email).strip().lower()] = fn
            if isinstance(billing, str):
                try:
                    billing = json.loads(billing)
                except Exception:
                    billing = {}
            if isinstance(billing, dict):
                be = (billing.get("email") or "").strip().lower()
                if be:
                    mapping.setdefault(be, fn)
    except Exception:
        pass
    return mapping


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    critical_failures = []   # any => hard FAIL
    soft_results = []        # (label, passed) -> accuracy

    def critical(label, passed, detail=""):
        if not passed:
            critical_failures.append(f"{label}{(' :: ' + detail) if detail else ''}")
        print(f"  [CRIT] {'PASS' if passed else 'FAIL'} - {label}"
              f"{(' :: ' + detail) if (detail and not passed) else ''}")

    def soft(label, passed, detail=""):
        soft_results.append((label, bool(passed)))
        print(f"  [soft] {'PASS' if passed else 'fail'} - {label}"
              f"{(' :: ' + detail) if (detail and not passed) else ''}")

    # ---- Form + questions ----
    print("Checking Form...")
    cur.execute("SELECT id, title, description FROM gform.forms ORDER BY created_at")
    forms = cur.fetchall()

    questions = []
    if not forms:
        critical("Form created", False, "no form rows in gform.forms")
    else:
        form = forms[0]
        form_id = form[0]
        critical("Form title matches (EN/RU)", str_contains_any(form[1], TITLE_ALTS),
                 f"title={form[1]!r}")
        soft("Form description matches (EN/RU)", str_contains_any(form[2], DESC_ALTS),
             f"desc={form[2]!r}")

        cur.execute(
            "SELECT title, question_type, required, position, config "
            "FROM gform.questions WHERE form_id=%s ORDER BY position, id",
            (form_id,),
        )
        questions = cur.fetchall()

        # CRITICAL: exactly 5 questions.
        critical("Exactly 5 questions in the form", len(questions) == 5,
                 f"found {len(questions)}")

        # CRITICAL: per-position titles match EN/RU alternatives (original .lower()).
        if len(questions) >= 5:
            for i, alts in enumerate(EXPECTED_QUESTIONS):
                title = questions[i][0]
                critical(f"Q{i+1} title matches expected (EN/RU)",
                         str_contains_any(title, alts), f"got {title!r}")
        else:
            critical("Per-position question titles", False,
                     "fewer than 5 questions; cannot verify order")

        if len(questions) >= 5:
            q1, q2, q3, q4, q5 = questions[:5]

            # CRITICAL: question types match the forms-MCP representation.
            # Q1 rating -> choiceQuestion with the 5 rating options (EN or RU).
            q1_opts = config_options(q1[4])
            q1_is_choice = q1[1] == "choiceQuestion"
            rating_en = ["excellent", "good", "average", "poor", "very poor"]
            rating_ru_hits = sum(
                1 for kw in ["отличн", "хорош", "средн", "плох", "очень плох"]
                if any(kw in o for o in q1_opts)
            )
            rating_en_hits = sum(1 for kw in rating_en if any(kw in o for o in q1_opts))
            critical("Q1 is choice with 5 rating options (Excellent..Very Poor / RU)",
                     q1_is_choice and len(q1_opts) == 5
                     and (rating_en_hits >= 4 or rating_ru_hits >= 4),
                     f"type={q1[1]} opts={q1_opts}")

            # Q2 scale 1-5 -> choiceQuestion covering integers 1..5.
            q2_opts = config_options(q2[4])
            critical("Q2 is a 1-5 scale (choice options covering 1..5)",
                     q2[1] == "choiceQuestion" and numeric_options_cover(q2_opts, 1, 5),
                     f"type={q2[1]} opts={q2_opts}")

            # Q3 scale 1-10 -> choiceQuestion covering integers 1..10.
            q3_opts = config_options(q3[4])
            critical("Q3 is a 1-10 scale (choice options covering 1..10)",
                     q3[1] == "choiceQuestion" and numeric_options_cover(q3_opts, 1, 10),
                     f"type={q3[1]} opts={q3_opts}")

            # Q4 / Q5 free text -> textQuestion.
            critical("Q4 is a free-text question (textQuestion)",
                     q4[1] == "textQuestion", f"type={q4[1]}")
            critical("Q5 is a free-text question (textQuestion)",
                     q5[1] == "textQuestion", f"type={q5[1]}")

            # CRITICAL: required flags. Q1-Q4 required, Q5 (suggestions) optional.
            req_ok = (bool(q1[2]) and bool(q2[2]) and bool(q3[2]) and bool(q4[2])
                      and not bool(q5[2]))
            critical("Required flags: Q1-Q4 required, Q5 optional",
                     req_ok,
                     f"required={[bool(q[2]) for q in questions[:5]]}")

    # ---- Emails ----
    print("Checking emails...")
    email_first_name = build_email_to_first_name(cur)

    cur.execute("SELECT to_addr, subject, body_text, from_addr FROM email.messages")
    messages = cur.fetchall()

    sent_to = set()
    survey_messages = []  # (recipient_email, subject, body, from_addr)
    for to_addr, subject, body, from_addr in messages:
        if not to_addr:
            continue
        found = re.findall(r"[\w.+-]+@[\w.-]+", str(to_addr).lower())
        for addr in found:
            sent_to.add(addr)
            if addr in EXPECTED_EMAILS:
                survey_messages.append((addr, subject, body, from_addr))

    missing = EXPECTED_EMAILS - sent_to

    # CRITICAL: every qualifying customer received the invite (exact recipient set).
    critical("All 36 qualifying customers received the invite",
             len(missing) == 0,
             f"missing {len(missing)}: {sorted(list(missing))[:8]}")

    # CRITICAL: subject keyword present (EN/RU) on survey emails.
    subj_ok = bool(survey_messages) and any(
        str_contains_any(m[1], SUBJECT_ALTS) for m in survey_messages
    )
    critical("Email subject mentions feedback/survey/satisfaction (EN/RU)",
             subj_ok,
             f"sample subject={survey_messages[0][1]!r}" if survey_messages else "no survey emails")

    # CRITICAL: per-recipient personalization (greeting present + recipient's first
    # name present in the body) AND sent from survey@store.com.
    personalized_ok = 0
    sender_ok = 0
    checked = 0
    for addr, subject, body, from_addr in survey_messages:
        checked += 1
        b = (body or "").lower()
        greet = str_contains_any(b, GREETING_ALTS)
        fname = email_first_name.get(addr, "")
        name_ok = bool(fname) and fname in b
        if greet and name_ok:
            personalized_ok += 1
        if from_addr and SURVEY_SENDER in str(from_addr).lower():
            sender_ok += 1

    # Require personalization on (almost) all delivered survey emails.
    if checked:
        critical("Email bodies personalized (greeting + recipient first name)",
                 personalized_ok >= max(1, checked - 1),
                 f"{personalized_ok}/{checked} personalized")
        critical("Emails sent from survey@store.com",
                 sender_ok >= max(1, checked - 1),
                 f"{sender_ok}/{checked} from {SURVEY_SENDER}")
    else:
        critical("Email bodies personalized (greeting + recipient first name)", False,
                 "no survey emails to inspect")
        critical("Emails sent from survey@store.com", False, "no survey emails")

    # ---- Soft / structural extras (contribute to accuracy) ----
    soft("No spam: not wildly more recipients than expected",
         len(sent_to - EXPECTED_EMAILS) <= 5,
         f"extra={len(sent_to - EXPECTED_EMAILS)}")
    soft("Form has a non-empty description", bool(forms) and bool(forms[0][2]))

    cur.close()
    conn.close()

    # ---- Verdict ----
    print("\n=== SUMMARY ===")
    if critical_failures:
        print(f"CRITICAL FAILURES ({len(critical_failures)}):")
        for f in critical_failures:
            print(f"  - {f}")
        print("\n=== RESULT: FAIL (critical) ===")
        sys.exit(1)

    total = len(soft_results)
    passed = sum(1 for _, ok in soft_results if ok)
    accuracy = 100.0 if total == 0 else (passed / total) * 100.0
    print(f"Soft checks: {passed}/{total} passed (accuracy={accuracy:.1f}%)")

    if accuracy >= 70:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)
    else:
        print("\n=== RESULT: FAIL (accuracy < 70) ===")
        sys.exit(1)


if __name__ == "__main__":
    main()
