"""
Evaluation для moex-dividend-analysis-pdf-email (RU / moex-finance).

Агент строит PDF Dividend_Analysis.pdf и письмо по данным MCP `moex-finance`
(схема moex.stock_info). Эталонные значения НЕ захардкожены: они читаются
вживую из moex.stock_info на момент проверки для четырёх выбранных тикеров
SBER.ME, GAZP.ME, LKOH.ME, MGNT.ME.

Метрики из сида moex.stock_info:
  trailingAnnualDividendYield  -> "текущая дивидендная доходность" (доля; *100 = %)
  trailingAnnualDividendRate   -> годовая сумма дивиденда на акцию (RUB)
  fiveYearAvgDividendYield     -> среднегодовая дивидендная доходность за 5 лет (%)
  previousClose                -> последняя цена закрытия (RUB)
  sector / longName            -> сектор / название компании

Ранжирование в отчёте и письме — по fiveYearAvgDividendYield (убывание),
т.к. trailingAnnualDividendYield у большинства тикеров = 0 (вырожденный ранг).

CRITICAL_CHECKS (семантика): любой их провал => общий FAIL независимо от
accuracy. Иначе порог: accuracy >= 70%. Структурные проверки
(наличие файла, размер, наличие секций) — не критические.
"""

import argparse
import json
import os
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

CHOSEN_SYMBOLS = ["SBER.ME", "GAZP.ME", "LKOH.ME", "MGNT.ME"]

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Семантические критические проверки (формулируются динамически ниже).
CRITICAL_CHECKS = {
    "PDF: дивидендные доходности (5Y avg) всех 4 тикеров корректны",
    "PDF: годовые суммы дивидендов (trailingAnnualDividendRate) всех 4 тикеров корректны",
    "PDF: цены закрытия (previousClose, RUB) всех 4 тикеров корректны",
    "PDF: ранжирование по 5Y avg yield и тикер-лидер корректны",
    "PDF: средняя 5Y дивдоходность и суммарная годовая сумма дивидендов корректны",
    "Email: письмо отправлено с корректными адресами, темой и телом",
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        tag = " (CRITICAL)" if name in CRITICAL_CHECKS else ""
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL]{tag} {name}{msg}")


def num_close(a, b, rel_tol=0.05, abs_tol=0.05):
    try:
        return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)
    except (TypeError, ValueError):
        return False


def load_moex(symbols):
    """Читаем живые данные из moex.stock_info для выбранных тикеров."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    out = {}
    for sym in symbols:
        cur.execute("SELECT data FROM moex.stock_info WHERE symbol = %s", (sym,))
        row = cur.fetchone()
        if not row:
            continue
        data = row[0]
        if isinstance(data, str):
            data = json.loads(data)
        # previousClose: берём так же, как MCP get_stock_info (pg_adapter info()),
        # т.е. второй-по-свежести close из moex.stock_prices, а не JSONB-поле,
        # иначе эталон расходится с единственным доступным агенту источником.
        cur.execute(
            "SELECT close FROM moex.stock_prices WHERE symbol = %s "
            "ORDER BY date DESC LIMIT 2",
            (sym,),
        )
        price_rows = cur.fetchall()
        if len(price_rows) > 1:
            prev_close = price_rows[1][0]
        elif price_rows:
            prev_close = price_rows[0][0]
        else:
            prev_close = data.get("previousClose")
        out[sym] = {
            "longName": data.get("longName"),
            "sector": data.get("sector"),
            "previousClose": float(prev_close) if prev_close is not None else None,
            "trYield": data.get("trailingAnnualDividendYield"),
            "trRate": data.get("trailingAnnualDividendRate"),
            "y5": data.get("fiveYearAvgDividendYield"),
        }
    cur.close()
    conn.close()
    return out


def text_has_number(text, value, tol_abs=0.05, tol_rel=0.05):
    """Проверяем, встречается ли в тексте число, близкое к value
    (учитывая форматирование с 2 знаками)."""
    import re
    if value is None:
        return False
    target = float(value)
    # Числа с возможными разделителями тысяч (запятая или пробел) и десятичной точкой.
    # Пробел как разделитель тысяч учитываем только между группами из 3 цифр.
    candidates = set()
    for m in re.finditer(r"-?\d[\d.,]*", text):
        candidates.add(m.group(0))
    # Дополнительно ловим "1 234,56" / "1 234.56" (пробел-разделитель тысяч).
    for m in re.finditer(r"-?\d{1,3}(?: \d{3})+(?:[.,]\d+)?", text):
        candidates.add(m.group(0))
    for tok in candidates:
        cand = tok.replace(" ", "")
        # Если есть и запятая, и точка: запятая — тысячи. Иначе запятая — десятичная.
        if "," in cand and "." in cand:
            cand = cand.replace(",", "")
        else:
            cand = cand.replace(",", ".")
        cand = cand.strip(".")
        try:
            num = float(cand)
        except ValueError:
            continue
        if abs(num - target) <= max(tol_abs, abs(target) * tol_rel):
            return True
    return False


def extract_pdf_text(pdf_path):
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            return "".join((p.extract_text() or "") for p in reader.pages)
    except ImportError:
        pass
    import pdfplumber
    with pdfplumber.open(pdf_path) as p:
        return "".join((pg.extract_text() or "") for pg in p.pages)


def check_pdf(agent_workspace, moex):
    print("\n=== Checking Dividend_Analysis.pdf ===")
    pdf_path = os.path.join(agent_workspace, "Dividend_Analysis.pdf")

    if not os.path.isfile(pdf_path):
        record("PDF file exists", False, f"Not found: {pdf_path}")
        # Критические проверки, зависящие от PDF, помечаем как провал.
        for c in CRITICAL_CHECKS:
            if c.startswith("PDF:"):
                record(c, False, "no PDF")
        return
    record("PDF file exists", True)
    record("PDF file size reasonable", os.path.getsize(pdf_path) > 1024,
           f"Size: {os.path.getsize(pdf_path)} bytes")

    try:
        text = extract_pdf_text(pdf_path)
    except Exception as e:
        record("PDF text extractable", False, str(e))
        for c in CRITICAL_CHECKS:
            if c.startswith("PDF:"):
                record(c, False, "cannot read PDF text")
        return

    text_lower = text.lower()

    # --- Структурные (не критические) ---
    record("PDF contains Dividend Analysis title",
           "dividend" in text_lower and "analysis" in text_lower)
    symbols_found = sum(1 for s in CHOSEN_SYMBOLS
                        if s in text or s.split(".")[0] in text)
    record("PDF lists all 4 stock symbols", symbols_found == 4,
           f"Found {symbols_found}/4 symbols")
    record("PDF has recommendation section", "recommendation" in text_lower)
    record("PDF has analysis summary section", "analysis summary" in text_lower)

    # Секторы в moex англоязычные ("Energy", "Financial Services", "Consumer Defensive").
    expected_sectors = {(moex[s]["sector"] or "").lower() for s in CHOSEN_SYMBOLS}
    sectors_found = sum(1 for sec in expected_sectors if sec and sec in text_lower)
    record("PDF mentions expected sectors",
           sectors_found >= max(1, len([s for s in expected_sectors if s]) - 1),
           f"Found {sectors_found}/{len(expected_sectors)} sectors")

    # --- CRITICAL: числовые значения, прочитанные вживую ---

    # 1) 5Y avg yield каждого тикера (в процентах, как в moex).
    y5_ok = all(
        moex[s]["y5"] is not None and text_has_number(text, moex[s]["y5"])
        for s in CHOSEN_SYMBOLS
    )
    record("PDF: дивидендные доходности (5Y avg) всех 4 тикеров корректны", y5_ok,
           {s: moex[s]["y5"] for s in CHOSEN_SYMBOLS})

    # 2) trailingAnnualDividendRate каждого тикера (RUB).
    #    Для нулевых ставок 0.00 тривиально присутствует, ключевой — LKOH=541.
    rate_ok = all(
        moex[s]["trRate"] is not None and text_has_number(text, moex[s]["trRate"])
        for s in CHOSEN_SYMBOLS
    )
    record("PDF: годовые суммы дивидендов (trailingAnnualDividendRate) всех 4 тикеров корректны",
           rate_ok, {s: moex[s]["trRate"] for s in CHOSEN_SYMBOLS})

    # 3) previousClose каждого тикера (RUB).
    price_ok = all(
        moex[s]["previousClose"] is not None and text_has_number(text, moex[s]["previousClose"])
        for s in CHOSEN_SYMBOLS
    )
    record("PDF: цены закрытия (previousClose, RUB) всех 4 тикеров корректны", price_ok,
           {s: moex[s]["previousClose"] for s in CHOSEN_SYMBOLS})

    # 4) Ранжирование по 5Y avg yield и тикер-лидер.
    ranked = sorted(CHOSEN_SYMBOLS, key=lambda s: moex[s]["y5"], reverse=True)
    top_sym = ranked[0]
    top_base = top_sym.split(".")[0]
    # лидер назван в тексте рядом с "highest"/"наибольш" или как пункт "1."
    leader_named = (top_base.lower() in text_lower) and (
        "highest" in text_lower or "наибольш" in text_lower or "1." in text
    )
    # порядок ранжирования: первое вхождение тикеров в Analysis Summary убывает по yield.
    summary_idx = text_lower.find("analysis summary")
    summary_text = text[summary_idx:] if summary_idx >= 0 else text
    positions = []
    order_ok = True
    for s in ranked:
        base = s.split(".")[0]
        pos = summary_text.find(base)
        if pos < 0:
            order_ok = False
            break
        positions.append(pos)
    if order_ok:
        order_ok = positions == sorted(positions)
    record("PDF: ранжирование по 5Y avg yield и тикер-лидер корректны",
           leader_named and order_ok,
           f"top={top_sym}, ranked={ranked}, leader_named={leader_named}, order_ok={order_ok}")

    # 5) Средняя 5Y дивдоходность и суммарная годовая сумма дивидендов.
    avg_y5 = sum(moex[s]["y5"] for s in CHOSEN_SYMBOLS) / len(CHOSEN_SYMBOLS)
    total_rate = sum(moex[s]["trRate"] for s in CHOSEN_SYMBOLS)
    avg_ok = text_has_number(text, avg_y5, tol_abs=0.05)
    total_ok = text_has_number(text, total_rate, tol_abs=0.5)
    record("PDF: средняя 5Y дивдоходность и суммарная годовая сумма дивидендов корректны",
           avg_ok and total_ok,
           f"avg5Y={avg_y5:.2f} (ok={avg_ok}), total_rate={total_rate:.2f} (ok={total_ok})")


def check_email(moex):
    print("\n=== Checking Email ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
        emails = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Email: письмо отправлено с корректными адресами, темой и телом", False, str(e))
        return

    ranked = sorted(CHOSEN_SYMBOLS, key=lambda s: moex[s]["y5"], reverse=True)
    avg_y5 = sum(moex[s]["y5"] for s in CHOSEN_SYMBOLS) / len(CHOSEN_SYMBOLS)

    target = None
    for subject, from_addr, to_addr, body_text in emails:
        if "dividend analysis summary" in (subject or "").lower():
            target = (subject, from_addr, to_addr, body_text or "")
            break

    if target is None:
        record("Email: письмо отправлено с корректными адресами, темой и телом", False,
               f"Нет письма с темой 'Dividend Analysis Summary'. Всего писем: {len(emails)}")
        return

    subject, from_addr, to_addr, body = target
    from_ok = "analyst@investteam.com" in str(from_addr or "").lower()
    to_ok = "team@investteam.com" in str(to_addr or "").lower()
    subj_ok = "dividend analysis summary - march 2026" in (subject or "").lower()

    # Тело: перечислены все 4 тикера, присутствует средняя 5Y доходность.
    symbols_in_body = sum(1 for s in CHOSEN_SYMBOLS
                          if s in body or s.split(".")[0] in body)
    body_syms_ok = symbols_in_body == 4
    avg_in_body = text_has_number(body, avg_y5, tol_abs=0.05)

    # Порядок сортировки по 5Y доходности (убывание) в теле письма.
    positions = []
    order_ok = True
    for s in ranked:
        base = s.split(".")[0]
        pos = body.find(base)
        if pos < 0:
            order_ok = False
            break
        positions.append(pos)
    if order_ok:
        order_ok = positions == sorted(positions)

    all_ok = from_ok and to_ok and subj_ok and body_syms_ok and avg_in_body and order_ok
    record("Email: письмо отправлено с корректными адресами, темой и телом", all_ok,
           f"from={from_ok}, to={to_ok}, subj={subj_ok}, syms={symbols_in_body}/4, "
           f"avg_in_body={avg_in_body}, order_ok={order_ok}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("MOEX DIVIDEND ANALYSIS PDF EMAIL - EVALUATION")
    print("=" * 70)

    try:
        moex = load_moex(CHOSEN_SYMBOLS)
    except Exception as e:
        print(f"FATAL: cannot read moex.stock_info: {e}")
        sys.exit(1)

    missing = [s for s in CHOSEN_SYMBOLS if s not in moex]
    if missing:
        print(f"FATAL: missing moex tickers in seed: {missing}")
        sys.exit(1)

    check_pdf(args.agent_workspace, moex)
    check_email(moex)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}, Accuracy: {accuracy:.1f}%")
    if critical_failed:
        print(f"  CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"    - {n}")

    success = (not critical_failed) and accuracy >= 70
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({
                "total_passed": PASS_COUNT,
                "total_checks": total,
                "accuracy": accuracy,
                "critical_failed": critical_failed,
                "success": success,
            }, f, indent=2)

    if critical_failed:
        print("Overall: FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 70:
        print("Overall: PASS")
        sys.exit(0)
    print("Overall: FAIL")
    sys.exit(1)


if __name__ == "__main__":
    main()
