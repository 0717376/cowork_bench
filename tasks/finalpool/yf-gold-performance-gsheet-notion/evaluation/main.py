"""Evaluation для yf-gold-performance-gsheet-notion (RU-стек: moex-finance / teamly).

Источник данных — схема moex.* (российский аналог yahoo-finance).
Золото: GLDRUB_TOM. Акции: OZON.ME, YNDX.ME, PHOR.ME, VTBR.ME, ROSN.ME.

Check 1: Google Sheet «Gold vs Stocks Analysis» с листом «Returns Comparison».
Check 2: страница teamly «Gold vs Stocks Performance».

Все эталонные значения (цены/доходности/shortName) вычисляются из живой БД moex,
поэтому eval автоматически следует за реальными засеянными числами.

CRITICAL_CHECKS гейтят итог: при провале любого критического чека → sys.exit(1)
независимо от accuracy.
"""

import argparse
import json
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname="cowork_gym", user="eigent", password="camel")

# Золото + 5 акций (тикеры MOEX, английские идентификаторы сохраняются).
GOLD = "GLDRUB_TOM"
STOCKS = ["OZON.ME", "YNDX.ME", "PHOR.ME", "VTBR.ME", "ROSN.ME"]
SYMBOLS = STOCKS + [GOLD]

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Содержательные проверки — провал любой рушит итог независимо от accuracy.
CRITICAL_CHECKS = {
    "Таблица с 'Gold' в названии существует",
    "Лист 'Returns Comparison' существует",
    f"Первая строка данных содержит лидера года ({None})",  # имя подставится ниже
    "Страница teamly 'Gold vs Stocks Performance' найдена (не сидовая)",
    "Сводка teamly: верное число акций, обогнавших золото по году",
}
# Заглушку для лидера заменим динамически после расчёта GT (см. main()).


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        detail_str = f": {detail[:200]}" if detail else ""
        flag = " [CRITICAL]" if critical else ""
        print(f"  [FAIL]{flag} {name}{detail_str}")


def num_close(a, b, tol=2.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_data():
    """Вычислить эталонные доходности из БД moex (mirror логики задачи)."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    placeholders = ",".join(["%s"] * len(SYMBOLS))

    # Последние цены (по последней дате OZON.ME — все 6 активов имеют общую дату).
    cur.execute(f"""SELECT symbol, close FROM moex.stock_prices
        WHERE date = (SELECT MAX(date) FROM moex.stock_prices WHERE symbol='OZON.ME')
        AND symbol IN ({placeholders})""", SYMBOLS)
    latest = dict(cur.fetchall())

    # 6 месяцев назад (~126 торговых дней; rn=127).
    cur.execute(f"""
        WITH ranked AS (SELECT symbol, date, close,
            ROW_NUMBER() OVER(PARTITION BY symbol ORDER BY date DESC) as rn
        FROM moex.stock_prices WHERE symbol IN ({placeholders}))
        SELECT symbol, close FROM ranked WHERE rn = 127
    """, SYMBOLS)
    six_m = dict(cur.fetchall())

    # 1 год назад (~252 торговых дня; rn=253).
    cur.execute(f"""
        WITH ranked AS (SELECT symbol, date, close,
            ROW_NUMBER() OVER(PARTITION BY symbol ORDER BY date DESC) as rn
        FROM moex.stock_prices WHERE symbol IN ({placeholders}))
        SELECT symbol, close FROM ranked WHERE rn = 253
    """, SYMBOLS)
    one_y = dict(cur.fetchall())

    # Краткие названия (shortName).
    cur.execute(f"""SELECT symbol, data->>'shortName' FROM moex.stock_info
        WHERE symbol IN ({placeholders})""", SYMBOLS)
    names = dict(cur.fetchall())

    cur.close()
    conn.close()

    gold_6m_ret = round((float(latest[GOLD]) - float(six_m[GOLD])) / float(six_m[GOLD]) * 100, 2)
    gold_1y_ret = round((float(latest[GOLD]) - float(one_y[GOLD])) / float(one_y[GOLD]) * 100, 2)

    results = []
    for s in SYMBOLS:
        l = float(latest[s])
        s6 = float(six_m[s])
        s1 = float(one_y[s])
        ret_6m = round((l - s6) / s6 * 100, 2)
        ret_1y = round((l - s1) / s1 * 100, 2)
        beat_6m = "Yes" if ret_6m > gold_6m_ret and s != GOLD else "No"
        beat_1y = "Yes" if ret_1y > gold_1y_ret and s != GOLD else "No"
        results.append({
            "symbol": s, "name": names.get(s, s),
            "latest": l, "price_6m": s6, "ret_6m": ret_6m,
            "price_1y": s1, "ret_1y": ret_1y,
            "beat_6m": beat_6m, "beat_1y": beat_1y,
        })

    results.sort(key=lambda x: x["ret_1y"], reverse=True)
    return results, gold_6m_ret, gold_1y_ret


def check_gsheet(expected_data, gold_6m, gold_1y):
    """Проверка Google Sheet с таблицей доходностей."""
    print("\n=== Проверка Google Sheet ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%gold%'")
    spreadsheets = cur.fetchall()
    check("Таблица с 'Gold' в названии существует", len(spreadsheets) > 0,
          "Не найдена таблица с 'Gold' в названии", critical=True)
    if not spreadsheets:
        cur.close()
        conn.close()
        return

    ss_id = spreadsheets[0][0]
    print(f"  Найдена таблица: '{spreadsheets[0][1]}' (id={ss_id})")

    cur.execute("""
        SELECT id FROM gsheet.sheets
        WHERE spreadsheet_id = %s AND LOWER(title) LIKE '%%return%%'
    """, (ss_id,))
    sheets = cur.fetchall()
    check("Лист 'Returns Comparison' существует", len(sheets) > 0, critical=True)
    if not sheets:
        cur.close()
        conn.close()
        return

    sheet_id = sheets[0][0]

    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (ss_id, sheet_id))
    cells = cur.fetchall()

    grid = {}
    for row_idx, col_idx, value in cells:
        grid[(row_idx, col_idx)] = value

    max_row = max((r for r, c in grid.keys()), default=0)
    check("Минимум 6 строк данных (6 активов)", max_row >= 6, f"max_row={max_row}")

    all_values = [str(v).strip().lower() for v in grid.values() if v]
    all_values_raw = [str(v).strip() for v in grid.values() if v]

    # Все тикеры присутствуют.
    for item in expected_data:
        sym = item["symbol"]
        found = sym.lower() in all_values or sym in all_values_raw
        check(f"Тикер {sym} присутствует в таблице", found)

    # Значения годовой доходности (с допуском).
    for item in expected_data:
        sym = item["symbol"]
        ret_1y_found = any(num_close(v, item["ret_1y"], 3.0) for v in all_values_raw
                           if v.replace('.', '', 1).replace('-', '', 1).isdigit())
        check(f"{sym}: годовая доходность ~{item['ret_1y']}% присутствует", ret_1y_found)

    # Значения 6-месячной доходности (с допуском) — выборочно по золоту и лидеру.
    for item in expected_data:
        sym = item["symbol"]
        ret_6m_found = any(num_close(v, item["ret_6m"], 3.0) for v in all_values_raw
                           if v.replace('.', '', 1).replace('-', '', 1).isdigit())
        check(f"{sym}: 6-месячная доходность ~{item['ret_6m']}% присутствует", ret_6m_found)

    # Сортировка: первая строка данных = лидер по годовой доходности.
    row1_vals = [grid.get((1, c), "") for c in range(10)]
    top_symbol = expected_data[0]["symbol"]
    check(f"Первая строка данных содержит лидера года ({top_symbol})",
          any(top_symbol in str(v) for v in row1_vals),
          f"Строка 1: {row1_vals[:5]}", critical=True)

    cur.close()
    conn.close()


def check_teamly(expected_data, gold_6m, gold_1y, beat_6m_count, beat_1y_count, top_symbol_name):
    """Проверка страницы teamly 'Gold vs Stocks Performance'.

    Сидовые страницы имеют id <= 3 — они не должны удовлетворять проверке.
    """
    print("\n=== Проверка teamly ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is None:
        check("Страница teamly 'Gold vs Stocks Performance' найдена (не сидовая)",
              False, "схема teamly не найдена", critical=True)
        cur.close()
        conn.close()
        return

    cur.execute("SELECT id, title, COALESCE(body, '') FROM teamly.pages WHERE id > 3")
    pages = cur.fetchall()

    page = None
    for pid, title, body in pages:
        tl = (title or "").lower()
        if "gold" in tl and ("stock" in tl or "performance" in tl):
            page = (pid, title, body)
            break
    check("Страница teamly 'Gold vs Stocks Performance' найдена (не сидовая)",
          page is not None,
          f"новые страницы: {[(p[0], p[1]) for p in pages]}", critical=True)

    if page is None:
        cur.close()
        conn.close()
        return

    text = ((page[1] or "") + " " + (page[2] or "")).lower()

    check("Сводка teamly упоминает золото",
          "gold" in text or "золот" in text, "нет упоминания золота")
    check("Сводка teamly упоминает доходность/показатели",
          "%" in text or "доходност" in text or "return" in text or "perform" in text,
          "нет упоминания доходности")

    # Содержательно: верное число акций, обогнавших золото по году.
    # beat_1y_count = 1 (OZON.ME) при текущем сиде. Проверяем наличие числа в тексте.
    digits_words = {0: "ноль", 1: "одн", 2: "дв", 3: "три", 4: "четыр", 5: "пят"}
    n1 = beat_1y_count
    has_1y_count = str(n1) in text or digits_words.get(n1, "###") in text
    check("Сводка teamly: верное число акций, обогнавших золото по году",
          has_1y_count, f"ожидалось {n1}", critical=True)

    # Лидер года упомянут (тикер или имя).
    top_sym = expected_data[0]["symbol"]
    nm = (top_symbol_name or "").lower()
    has_leader = top_sym.lower() in text or (nm and nm.split()[0] in text)
    check("Сводка teamly упоминает лидера по годовой доходности",
          has_leader, f"ожидался {top_sym} / {top_symbol_name}")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("MOEX GOLD PERFORMANCE GSHEET TEAMLY - EVALUATION")
    print("=" * 70)

    expected_data, gold_6m, gold_1y = get_expected_data()
    top_symbol = expected_data[0]["symbol"]
    top_symbol_name = expected_data[0]["name"]
    beat_6m_count = sum(1 for x in expected_data if x["beat_6m"] == "Yes")
    beat_1y_count = sum(1 for x in expected_data if x["beat_1y"] == "Yes")

    # Подставить динамическое имя критического чека по лидеру года.
    CRITICAL_CHECKS.discard(f"Первая строка данных содержит лидера года ({None})")
    CRITICAL_CHECKS.add(f"Первая строка данных содержит лидера года ({top_symbol})")

    print(f"\n[GT] Золото 6M={gold_6m}%  1Y={gold_1y}%")
    print(f"[GT] Обогнали золото: 6M={beat_6m_count}, 1Y={beat_1y_count}")
    print(f"[GT] Лидер года: {top_symbol} ({top_symbol_name}) {expected_data[0]['ret_1y']}%")

    check_gsheet(expected_data, gold_6m, gold_1y)
    check_teamly(expected_data, gold_6m, gold_1y,
                 beat_6m_count, beat_1y_count, top_symbol_name)

    critical_fails = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print(f"\n=== ИТОГ ===")
    print(f"  Пройдено: {PASS_COUNT}")
    print(f"  Провалено: {FAIL_COUNT}")
    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total if total else 0.0
    print(f"  Accuracy: {accuracy:.2%}")

    if critical_fails:
        print(f"  КРИТИЧЕСКИЙ ПРОВАЛ: {critical_fails}")
        print("  Overall: FAIL")
        sys.exit(1)

    if accuracy >= 0.70:
        print("  Overall: PASS")
        sys.exit(0)
    else:
        print("  Overall: FAIL (accuracy < 70%)")
        sys.exit(1)


if __name__ == "__main__":
    main()
