"""Evaluation for insales-product-review-gform (InSales + RU Forms)."""
import argparse
import json
import os
import sys

import psycopg2

# Required email subject line (kept English per task.md — it is the literal grepped value).
REQUIRED_SUBJECT = "We Value Your Feedback - Product Quality Survey"
# Title keywords accepted in either Russian or English.
TITLE_KEYWORDS = ["product", "quality", "feedback", "качеств", "продукц", "отзыв", "опрос"]


def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")
PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def record(name, passed, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRITICAL " if critical else ""
    if passed:
        PASS_COUNT += 1; print(f"  [PASS] {tag}{name}")
    else:
        FAIL_COUNT += 1; print(f"  [FAIL] {tag}{name}: {str(detail)[:300]}")
        if critical:
            CRITICAL_FAILED.append(name)


def get_expected():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""SELECT id, name, ROUND(average_rating::numeric,2), total_sales
        FROM wc.products WHERE average_rating IS NOT NULL AND average_rating::numeric > 0
        ORDER BY average_rating ASC, total_sales DESC LIMIT 5""")
    products = [{"id": r[0], "name": r[1], "rating": float(r[2]), "sales": r[3]} for r in cur.fetchall()]
    product_ids = [p["id"] for p in products]

    placeholders = ",".join(["%s"] * len(product_ids))
    cur.execute(f"""SELECT DISTINCT o.customer_id, c.email, c.first_name, c.last_name
        FROM wc.orders o JOIN wc.customers c ON c.id=o.customer_id,
        LATERAL jsonb_array_elements(o.line_items) AS item
        WHERE (item->>'product_id')::int IN ({placeholders}) AND o.customer_id > 0
        ORDER BY o.customer_id""", product_ids)
    customers = [{"id": r[0], "email": r[1], "first_name": r[2], "last_name": r[3]} for r in cur.fetchall()]
    conn.close()
    return {"products": products, "customers": customers}


def check_gform(expected):
    print("\n=== Checking Forms ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title, description FROM gform.forms")
    forms = cur.fetchall()
    record("At least 1 form created", len(forms) >= 1, f"Found {len(forms)}")

    def title_match(t):
        tl = (t or "").lower()
        return any(k in tl for k in TITLE_KEYWORDS)

    target_form = None
    for fid, title, desc in forms:
        if title_match(title):
            target_form = fid
            break
    if target_form is None and forms:
        target_form = forms[0][0]

    if target_form is None:
        record("Survey form with description and 5 required rating questions", False,
               "No form found", critical=True)
        conn.close()
        return

    record("Form titled with product/quality/feedback (RU+EN)",
           any(title_match(f[1]) for f in forms))

    # Description of the chosen form must be non-empty.
    desc = None
    for fid, title, d in forms:
        if fid == target_form:
            desc = d
            break
    has_desc = bool((desc or "").strip())

    cur.execute("SELECT title, question_type, required FROM gform.questions WHERE form_id=%s ORDER BY position", (target_form,))
    questions = cur.fetchall()
    record("At least 5 questions", len(questions) >= 5, f"Found {len(questions)}")

    required_count = sum(1 for q in questions if q[2])

    # CRITICAL: survey form with non-empty description AND >=5 required questions.
    record("Survey form with non-empty description and >=5 required questions",
           has_desc and required_count >= 5,
           f"description_set={has_desc}, required_questions={required_count}", critical=True)

    # CRITICAL: the questions cover the correct 5 lowest-rated products (>=4 of 5 names present).
    matched_products = 0
    for prod in expected["products"]:
        prod_name_lower = prod["name"].lower()[:30]
        prod_name_20 = prod["name"][:20].lower()
        for q_title, q_type, q_req in questions:
            qt = (q_title or "").lower()
            if prod_name_lower in qt or prod_name_20 in qt:
                matched_products += 1
                break
    record("Questions cover correct 5 lowest-rated products (>=4/5 names)",
           matched_products >= 4,
           f"Found {matched_products}/5 product names in {len(questions)} questions", critical=True)

    cur.close()
    conn.close()


def check_emails(expected):
    print("\n=== Checking Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("Emails sent", len(emails) >= len(expected["customers"]),
           f"Found {len(emails)}, expected >= {len(expected['customers'])}")

    def recipients(to):
        out = []
        if isinstance(to, list):
            out.extend([str(r).lower() for r in to])
        elif isinstance(to, str):
            try:
                parsed = json.loads(to)
                if isinstance(parsed, list):
                    out.extend([str(r).lower() for r in parsed])
                else:
                    out.append(to.lower())
            except Exception:
                out.append(to.lower())
        return out

    all_to = []
    for subj, to, body in emails:
        all_to.extend(recipients(to))
    all_to_blob = " ".join(all_to)

    # CRITICAL: emails reach ALL affected customers (computed from live order/line-item join).
    matched_customers = 0
    for cust in expected["customers"]:
        if cust["email"].lower() in all_to_blob:
            matched_customers += 1
    record("Emails reach ALL affected customers",
           matched_customers == len(expected["customers"]),
           f"Matched {matched_customers}/{len(expected['customers'])}", critical=True)

    # CRITICAL: every sent email subject equals the required subject line.
    bad_subjects = [s for (s, t, b) in emails if (s or "").strip() != REQUIRED_SUBJECT]
    record("Every email subject equals required subject line",
           len(emails) > 0 and len(bad_subjects) == 0,
           f"{len(bad_subjects)} mismatched subject(s); required={REQUIRED_SUBJECT!r}", critical=True)

    # CRITICAL: each email body is personalized with the recipient's first_name.
    by_email = {c["email"].lower(): c for c in expected["customers"]}
    personalized = 0
    checked = 0
    for subj, to, body in emails:
        body_l = (body or "").lower()
        for r in recipients(to):
            cust = by_email.get(r)
            if not cust:
                continue
            checked += 1
            fn = (cust["first_name"] or "").strip().lower()
            if fn and fn in body_l:
                personalized += 1
            break
    record("Each email body personalized with recipient first_name",
           checked > 0 and personalized == checked,
           f"Personalized {personalized}/{checked} bodies", critical=True)

    # Non-critical: body has substantive content.
    if emails:
        record("Email body has content", len((emails[0][2] or "")) > 20,
               f"Body length: {len(emails[0][2] or '')}")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", default=".")
    parser.add_argument("--groundtruth_workspace", default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    expected = get_expected()
    check_gform(expected)
    check_emails(expected)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    print(f"\n=== SUMMARY: {PASS_COUNT} passed, {FAIL_COUNT} failed, accuracy={accuracy:.1f}% ===")

    success = (not CRITICAL_FAILED) and accuracy >= 70
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT,
                       "accuracy": accuracy, "critical_failed": CRITICAL_FAILED,
                       "success": success}, f)

    if CRITICAL_FAILED:
        print(f"CRITICAL checks failed: {CRITICAL_FAILED}")
        sys.exit(1)
    sys.exit(0 if accuracy >= 70 else 1)


if __name__ == "__main__":
    main()
