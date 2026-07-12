"""Evaluation для market-news-digest-moex-teamly-forms-email.

Проверки:
1. Teamly: страница «Рыночный новостной дайджест» (Market News Digest) создана,
   ссылается минимум на 3 из 5 тикеров MOEX (SBER/GAZP/LKOH/MGNT/MTSS) и содержит
   не менее 5 новостных записей с метками полей Title/Symbol/Publisher/Published_Date/Summary.
2. Forms: опрос «Опрос настроений инвесторов» с ровно 4 вопросами, семантически
   покрывающими: перспективы рынка, сектор, инвестиционные опасения, увеличение
   доли акций; у каждого вопроса присутствуют обязательные варианты ответа.
3. Email: письмо на subscribers@newsletter.example.com с темой, содержащей маркеры
   Weekly + Market + Digest, и телом, упоминающим минимум 3 из 5 тикеров.

КРИТИЧЕСКИЕ чеки (CRITICAL_CHECKS): любой их провал => FAIL независимо от accuracy.
Порог: accuracy >= 70% И нет критических провалов => PASS.
"""
import argparse
import json
import os
import sys

import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Базовые символы тикеров (без .ME) — для grep по тексту дайджеста и тела письма.
TICKERS = ["SBER", "GAZP", "LKOH", "MGNT", "MTSS"]

CRITICAL_CHECKS = {
    "Teamly: дайджест создан и ссылается на >=3 тикера MOEX",
    "Forms: ровно 4 вопроса, покрывающих outlook/sector/concern/allocation",
    "Email: письмо подписчикам с темой Weekly+Market+Digest и >=3 тикерами в теле",
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def record(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAILED
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        d = (str(detail)[:300] + "...") if len(str(detail)) > 300 else str(detail)
        print(f"  [FAIL] {name}: {d}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED.append(name)


def count_tickers(text):
    up = (text or "").upper()
    return sum(1 for t in TICKERS if t in up)


# ---------------------------------------------------------------------------
# Check 1: Teamly digest page
# ---------------------------------------------------------------------------
def check_teamly():
    print("\n=== Check 1: Teamly — Рыночный новостной дайджест ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Ищем страницу-дайджест: RU 'дайджест/новост/рынок' ИЛИ EN 'market news digest'.
    cur.execute("""
        SELECT title, body FROM teamly.pages
        WHERE title ILIKE '%дайджест%'
           OR (title ILIKE '%новост%' AND title ILIKE '%рынок%')
           OR (title ILIKE '%market%' AND title ILIKE '%news%')
           OR (title ILIKE '%market%' AND title ILIKE '%digest%')
    """)
    rows = cur.fetchall()

    record("Teamly: страница рыночного дайджеста найдена", len(rows) >= 1,
           f"найдено страниц: {len(rows)}")

    body = "\n\n".join((b or "") for _t, b in rows)

    # CRITICAL: дайджест ссылается минимум на 3 из 5 тикеров.
    n_tick = count_tickers(body)
    record("Teamly: дайджест создан и ссылается на >=3 тикера MOEX",
           len(rows) >= 1 and n_tick >= 3,
           f"найдено тикеров: {n_tick}/5")

    # NON-critical структурный: >=5 записей с метками полей.
    bl = body.lower()
    label_hits = sum(1 for lbl in ("title", "symbol", "publisher", "published_date", "summary")
                     if lbl in bl)
    # Считаем число записей по числу вхождений метки Symbol (или тикеров).
    symbol_occurrences = max(bl.count("symbol"), count_distinct_ticker_lines(body))
    record("Teamly: дайджест содержит метки полей Title/Symbol/Publisher/Published_Date/Summary",
           label_hits >= 4, f"меток найдено: {label_hits}/5")
    record("Teamly: дайджест содержит не менее 5 новостных записей",
           symbol_occurrences >= 5, f"записей (по меткам/тикерам): {symbol_occurrences}")

    cur.close()
    conn.close()


def count_distinct_ticker_lines(body):
    """Оценка числа записей: считаем строки, содержащие любой тикер."""
    cnt = 0
    for line in (body or "").splitlines():
        up = line.upper()
        if any(t in up for t in TICKERS):
            cnt += 1
    return cnt


# ---------------------------------------------------------------------------
# Check 2: Forms survey
# ---------------------------------------------------------------------------
def check_forms():
    print("\n=== Check 2: Forms — Опрос настроений инвесторов ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    record("Forms: хотя бы одна форма существует", len(forms) > 0,
           "форм не найдено")

    found_id = None
    for fid, title in forms:
        tl = (title or "").lower()
        if ("настроен" in tl or "инвестор" in tl or "sentiment" in tl
                or "опрос" in tl or "market" in tl):
            found_id = fid
            break
    if found_id is None and forms:
        found_id = forms[0][0]

    record("Forms: найдена форма опроса настроений инвесторов", found_id is not None,
           f"формы: {[(str(r[0])[:20], r[1]) for r in forms]}")

    q_count = 0
    questions = []
    configs = []
    if found_id:
        cur.execute(
            "SELECT title, config FROM gform.questions WHERE form_id = %s ORDER BY position",
            (found_id,),
        )
        qrows = cur.fetchall()
        q_count = len(qrows)
        questions = [(r[0] or "") for r in qrows]
        configs = [r[1] for r in qrows]

    ql = [q.lower() for q in questions]

    # Семантическое покрытие 4 тем (RU + EN ключевые слова).
    has_outlook = any(("перспектив" in q and "рынк" in q) or "outlook" in q
                      or ("оцен" in q and "рынк" in q) for q in ql)
    has_sector = any("сектор" in q or "sector" in q for q in ql)
    has_concern = any("беспоко" in q or "опасен" in q or "concern" in q for q in ql)
    has_alloc = any(("долю акц" in q or "доля акц" in q or "увеличит" in q
                     or "allocation" in q or "equity" in q) for q in ql)
    covered = sum([has_outlook, has_sector, has_concern, has_alloc])

    # CRITICAL: ровно 4 вопроса И покрыты все 4 темы.
    record("Forms: ровно 4 вопроса, покрывающих outlook/sector/concern/allocation",
           q_count == 4 and covered == 4,
           f"вопросов={q_count}, покрыто тем={covered}/4 "
           f"(outlook={has_outlook}, sector={has_sector}, concern={has_concern}, alloc={has_alloc})")

    # NON-critical: у каждого из 4 вопросов есть варианты ответа.
    opt_ok = 0
    for cfg in configs:
        opts = []
        if isinstance(cfg, dict):
            opts = cfg.get("options", []) or []
        elif isinstance(cfg, str):
            try:
                opts = (json.loads(cfg) or {}).get("options", []) or []
            except Exception:
                opts = []
        if len(opts) >= 3:
            opt_ok += 1
    record("Forms: у каждого из 4 вопросов присутствуют варианты ответа (>=3)",
           q_count == 4 and opt_ok == 4, f"вопросов с вариантами: {opt_ok}/{q_count}")

    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Check 3: Email
# ---------------------------------------------------------------------------
def check_email():
    print("\n=== Check 3: Email подписчикам ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, to_addr::text, from_addr::text, COALESCE(body_text, '')
        FROM email.messages
        WHERE to_addr::text ILIKE '%subscribers@newsletter.example.com%'
           OR to_addr::text ILIKE '%subscribers%'
    """)
    rows = cur.fetchall()

    record("Email: письмо подписчикам найдено", len(rows) > 0,
           "нет писем на subscribers@newsletter.example.com")

    # Выбираем письмо с корректной темой (Weekly+Market+Digest), если есть.
    best = None
    for subject, to_addr, from_addr, body in rows:
        s = (subject or "").lower()
        if "weekly" in s and "market" in s and "digest" in s:
            best = (subject, to_addr, from_addr, body)
            break
    if best is None and rows:
        best = rows[0]

    subj = (best[0] if best else "") or ""
    to_addr = (best[1] if best else "") or ""
    body = (best[3] if best else "") or ""
    sl = subj.lower()

    subject_ok = ("weekly" in sl and "market" in sl and "digest" in sl)
    to_ok = "subscribers" in to_addr.lower()
    n_tick = count_tickers(body)

    # CRITICAL: тема содержит все 3 маркера, адресат подписчики, в теле >=3 тикеров.
    record("Email: письмо подписчикам с темой Weekly+Market+Digest и >=3 тикерами в теле",
           bool(best) and subject_ok and to_ok and n_tick >= 3,
           f"subject_ok={subject_ok}, to_ok={to_ok}, тикеров в теле={n_tick}/5; "
           f"subject={subj!r}")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=== Evaluation: market-news-digest-moex-teamly-forms-email ===")

    check_teamly()
    check_forms()
    check_email()

    total = PASS_COUNT + FAIL_COUNT
    pct = 100.0 * PASS_COUNT / total if total else 0.0
    print(f"\n=== SUMMARY: {PASS_COUNT}/{total} проверок пройдено ({pct:.1f}%) ===")

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"pass": PASS_COUNT, "fail": FAIL_COUNT, "pct": pct}, f)

    if CRITICAL_FAILED:
        print(f"CRITICAL FAIL: {CRITICAL_FAILED}")
        print("FAIL")
        sys.exit(1)
    if pct < 70.0:
        print("FAIL (accuracy < 70%)")
        sys.exit(1)
    print("PASS")
    sys.exit(0)


if __name__ == "__main__":
    main()
