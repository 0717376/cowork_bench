"""Evaluation for sf-sales-discount-gsheet-email (ClickHouse / RU).

Структура проверок:
  - НЕСКОЛЬКО CRITICAL (семантических) проверок: ожидаемые значения непусты
    (русифицированный сид STATUS='Доставлен'), каждый сегмент присутствует с
    верными Order_Count/Discounted_Orders, совпадают Discount_Rate_Pct и
    Avg_Discount_Pct, совпадают Total_Revenue/Discounted_Revenue, строки
    отсортированы по Discount_Rate_Pct по убыванию, и отправлено письмо
    finance-team@company.com с темой "Segment Discount Analysis".
    Любой провал CRITICAL => sys.exit(1) ещё ДО порога точности.
  - Остальные структурные проверки идут к порогу accuracy >= 70.

Замечания по локали:
  - Данные SALES_DW русифицированы ЦЕНТРАЛЬНО (db/zzz_clickhouse_after_init.sql):
    STATUS 'Delivered'->'Доставлен', SEGMENT Consumer->'Частные клиенты' и т.д.
    Поэтому ожидаемые значения вычисляются по STATUS='Доставлен', а сегменты
    сравниваются как русифицированные строки из самого сида.
  - RU/EN-ключевые слова в письме ищутся в .lower() ОРИГИНАЛЬНОГО текста.
  - Английские идентификаторы (имена БД/таблиц/колонок, заголовки листа/темы,
    адрес получателя) сохранены как есть.
"""
import argparse
import os
import sys
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILURES = []


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def record_critical(name, passed, detail=""):
    """Критическая проверка: фиксируется отдельно и валит весь evaluator."""
    record("[CRITICAL] " + name, passed, detail)
    if not passed:
        CRITICAL_FAILURES.append(name)


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def rel_close(a, b, rel=0.02, abs_floor=1.0):
    """Относительная близость для денежных величин."""
    try:
        a = float(a); b = float(b)
    except (TypeError, ValueError):
        return False
    return abs(a - b) <= max(abs_floor, abs(b) * rel)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.parse_args()

    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except Exception as e:
        print(f"  [FATAL] Could not connect to PostgreSQL: {e}")
        sys.exit(1)

    # --- Ожидаемые значения из русифицированного сида ClickHouse (sf_data) ---
    expected_map = {}
    expected_order = []  # порядок сегментов по disc_rate DESC
    try:
        cur.execute('''
            SELECT c."SEGMENT", COUNT(o.*) as orders,
                   COUNT(CASE WHEN o."DISCOUNT" > 0 THEN 1 END) as disc_orders,
                   ROUND(100.0 * COUNT(CASE WHEN o."DISCOUNT" > 0 THEN 1 END)/COUNT(*)::numeric, 1) as disc_rate,
                   ROUND(AVG(CASE WHEN o."DISCOUNT" > 0 THEN o."DISCOUNT" END)::numeric * 100, 2) as avg_disc_pct,
                   ROUND(SUM(o."TOTAL_AMOUNT")::numeric, 2) as total_rev,
                   ROUND(SUM(CASE WHEN o."DISCOUNT" > 0 THEN o."TOTAL_AMOUNT" ELSE 0 END)::numeric, 2) as disc_rev,
                   ROUND(100.0 * SUM(CASE WHEN o."DISCOUNT" > 0 THEN o."TOTAL_AMOUNT" ELSE 0 END) / SUM(o."TOTAL_AMOUNT")::numeric, 1) as impact
            FROM sf_data."SALES_DW__PUBLIC__ORDERS" o
            JOIN sf_data."SALES_DW__PUBLIC__CUSTOMERS" c ON o."CUSTOMER_ID" = c."CUSTOMER_ID"
            WHERE o."STATUS" = 'Доставлен'
            GROUP BY c."SEGMENT" ORDER BY disc_rate DESC
        ''')
        expected = cur.fetchall()
        expected_map = {r[0]: r for r in expected}
        expected_order = [r[0] for r in expected]
    except Exception as e:
        print(f"  [FATAL] Could not compute expected values: {e}")
        sys.exit(1)

    # CRITICAL: ожидаемые значения непусты (защита от рассинхрона сид<->eval)
    record_critical(
        "Ожидаемые значения непусты (STATUS='Доставлен' даёт >=1 сегмент)",
        len(expected_map) >= 1,
        f"найдено сегментов: {len(expected_map)}",
    )

    # --- Google Sheet (структура) ---
    print("\n=== Google Sheet ===")
    grid = {}
    data_rows = {}
    sheet_ok = False
    try:
        cur.execute("SELECT id FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%discount%analysis%'")
        sheets = cur.fetchall()
        record("Google-таблица 'Discount Analysis Report' существует", bool(sheets))
        if sheets:
            sheet_id = sheets[0][0]
            cur.execute("SELECT id FROM gsheet.sheets WHERE spreadsheet_id = %s AND LOWER(title) LIKE '%%segment%%'", (sheet_id,))
            seg_sheets = cur.fetchall()
            record("Лист 'Segment Analysis' присутствует", bool(seg_sheets))
            if seg_sheets:
                seg_sheet_id = seg_sheets[0][0]
                cur.execute("SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id = %s ORDER BY row_index, col_index", (seg_sheet_id,))
                for row_idx, col_idx, value in cur.fetchall():
                    grid.setdefault(row_idx, {})[col_idx] = value
                # 0-индексная раскладка: заголовок в row0, данные с row1
                data_rows = {k: v for k, v in grid.items() if k >= 1}
                record(
                    f"Найдено достаточно строк данных ({len(data_rows)} >= {len(expected_map)})",
                    len(data_rows) >= len(expected_map),
                )
                sheet_ok = True
    except Exception as e:
        record("Проверка Google Sheet без ошибок", False, str(e))

    # Сопоставление строки сегмента (русифицированные значения)
    def find_row(seg):
        for _ri, rd in data_rows.items():
            sv = rd.get(0, "")
            if sv and str(sv).strip().lower() == seg.lower():
                return rd
        return None

    # CRITICAL: каждый сегмент присутствует с верными Order_Count и Discounted_Orders
    print("\n=== CRITICAL: сегменты, Order_Count, Discounted_Orders ===")
    all_segments_present = True
    for seg, exp in expected_map.items():
        rd = find_row(seg)
        if rd is None:
            all_segments_present = False
            record_critical(f"Сегмент '{seg}' присутствует в листе", False)
            continue
        record_critical(
            f"{seg}.Order_Count == {exp[1]}",
            num_close(rd.get(1, ""), exp[1], 1),
            f"{rd.get(1)} vs {exp[1]}",
        )
        record_critical(
            f"{seg}.Discounted_Orders == {exp[2]}",
            num_close(rd.get(2, ""), exp[2], 1),
            f"{rd.get(2)} vs {exp[2]}",
        )

    # CRITICAL: Discount_Rate_Pct и Avg_Discount_Pct (основной аналитический результат)
    print("\n=== CRITICAL: Discount_Rate_Pct, Avg_Discount_Pct ===")
    for seg, exp in expected_map.items():
        rd = find_row(seg)
        if rd is None:
            continue
        record_critical(
            f"{seg}.Discount_Rate_Pct ~= {exp[3]}",
            num_close(rd.get(3, ""), exp[3], 0.2),
            f"{rd.get(3)} vs {exp[3]}",
        )
        record_critical(
            f"{seg}.Avg_Discount_Pct ~= {exp[4]}",
            num_close(rd.get(4, ""), exp[4], 0.2),
            f"{rd.get(4)} vs {exp[4]}",
        )

    # CRITICAL: Total_Revenue и Discounted_Revenue (денежные величины, относит. допуск)
    print("\n=== CRITICAL: Total_Revenue, Discounted_Revenue ===")
    for seg, exp in expected_map.items():
        rd = find_row(seg)
        if rd is None:
            continue
        record_critical(
            f"{seg}.Total_Revenue ~= {exp[5]}",
            rel_close(rd.get(5, ""), exp[5], rel=0.02, abs_floor=1.0),
            f"{rd.get(5)} vs {exp[5]}",
        )
        record_critical(
            f"{seg}.Discounted_Revenue ~= {exp[6]}",
            rel_close(rd.get(6, ""), exp[6], rel=0.02, abs_floor=1.0),
            f"{rd.get(6)} vs {exp[6]}",
        )

    # NON-critical: Revenue_Impact_Pct (структурно проверяемая колонка)
    print("\n=== Revenue_Impact_Pct ===")
    for seg, exp in expected_map.items():
        rd = find_row(seg)
        if rd is None:
            continue
        record(
            f"{seg}.Revenue_Impact_Pct ~= {exp[7]}",
            num_close(rd.get(7, ""), exp[7], 0.3),
            f"{rd.get(7)} vs {exp[7]}",
        )

    # CRITICAL: сортировка строк по Discount_Rate_Pct по убыванию
    print("\n=== CRITICAL: сортировка по Discount_Rate_Pct (убывание) ===")
    if sheet_ok and len(expected_order) >= 2:
        sheet_seq = []
        for ri in sorted(data_rows.keys()):
            sv = data_rows[ri].get(0, "")
            if sv and str(sv).strip():
                sheet_seq.append(str(sv).strip().lower())
        expected_seq = [s.lower() for s in expected_order]
        # сравниваем порядок только по присутствующим в листе сегментам
        filtered = [s for s in sheet_seq if s in expected_seq]
        record_critical(
            "Строки отсортированы по Discount_Rate_Pct по убыванию",
            filtered == expected_seq,
            f"в листе {filtered} vs ожидалось {expected_seq}",
        )
    else:
        record_critical(
            "Строки отсортированы по Discount_Rate_Pct по убыванию",
            False,
            "недостаточно данных для проверки порядка",
        )

    # CRITICAL: письмо finance-team@company.com с темой 'Segment Discount Analysis'
    print("\n=== CRITICAL: письмо финансовой команде ===")
    try:
        cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
        email_rows = cur.fetchall()
    except Exception:
        conn.rollback()
        email_rows = []
    target_email = None
    for subj, to_addr, body in email_rows:
        s = (subj or "").lower()
        t = ",".join(to_addr).lower() if isinstance(to_addr, list) else (to_addr or "").lower()
        if "segment discount analysis" in s and "finance-team@company.com" in t:
            target_email = (subj, to_addr, body)
            break
    record_critical(
        "Отправлено письмо finance-team@company.com с темой 'Segment Discount Analysis'",
        target_email is not None,
        f"writes={len(email_rows)}",
    )

    # NON-critical: тело письма содержит осмысленное резюме (RU/EN ключевые слова)
    print("\n=== Тело письма ===")
    if target_email:
        body = (target_email[2] or "").lower()
        record(
            "Тело письма упоминает скидки/сегмент (RU/EN)",
            any(kw in body for kw in ["скидк", "сегмент", "discount", "segment"]),
            body[:120],
        )
        record(
            "Тело письма упоминает долю/процент или конкретный сегмент",
            any(kw in body for kw in ["процент", "доля", "rate", "%",
                                       "частные", "корпоратив", "государствен",
                                       "малый", "средний", "бизнес"]),
            body[:120],
        )

    cur.close()
    conn.close()

    # --- Итог: CRITICAL-гейт перед порогом точности ---
    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    if CRITICAL_FAILURES:
        print(f"\nCRITICAL FAILURE ({len(CRITICAL_FAILURES)}): {CRITICAL_FAILURES}")
        print("  Overall: FAIL")
        sys.exit(1)

    if accuracy >= 70:
        print("  Overall: PASS")
        sys.exit(0)
    else:
        print("  Overall: FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
