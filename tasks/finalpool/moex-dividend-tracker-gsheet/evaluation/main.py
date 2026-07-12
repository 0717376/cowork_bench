"""
Evaluation для yf-dividend-tracker-gsheet (RU / moex-finance).

Агент через MCP `moex-finance` (схема moex.*) собирает дивиденды и корпоративные
действия по пяти тикерам Московской биржи (.ME), кладёт их в Google-таблицу
"Dividend Tracker" / лист "Stock Actions" и отправляет письмо с темой
"Dividend Action Summary".

Источник истины для значений — таблица moex.stock_info (читается динамически).
Маппинг полей:
  Dividend.Value            = lastDividendValue (RUB)
  Dividend.Dividend_Yield   = trailingAnnualDividendYield * 100 (в %)
  Dividend.Trailing_Annual_Rate = trailingAnnualDividendRate
  Dividend.Date             = exDividendDate (YYYY-MM-DD, UTC)
  Stock Split.Value         = lastSplitFactor (например "1000:1")
  Stock Split.Date          = lastSplitDate (YYYY-MM-DD, UTC)
"Платит дивиденды сейчас" <=> exDividendDate не пустой.

CRITICAL_CHECKS (семантика): любой их провал => общий FAIL независимо от accuracy.
Структурные проверки (наличие таблицы/листа/строк, ISO-дата) — не критические.
"""
import argparse
import json
import os
import sys
from datetime import datetime, timezone

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# Пять отслеживаемых тикеров MOEX (подмножество глобального сида moex.*)
TICKERS = ["SBER.ME", "GAZP.ME", "LKOH.ME", "MGNT.ME", "TCSG.ME"]
# Базы тикеров без суффикса .ME (для подстрочного поиска в ячейках/письме)
TICKER_BASES = [t.split(".")[0].lower() for t in TICKERS]  # sber gazp lkoh mgnt tcsg

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = False

CRITICAL_CHECKS = {
    "Набор плательщиков дивидендов совпадает с источником moex",
    "Дивидендные значения по тикерам совпадают с источником moex",
    "Строки дробления акций совпадают с источником moex",
    "Таблица 'Dividend Tracker' / лист 'Stock Actions' существует с данными",
    "Письмо 'Dividend Action Summary' отправлено инвестору",
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAILED
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        tag = " (CRITICAL)" if name in CRITICAL_CHECKS else ""
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL]{tag} {name}{msg}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED = True


def str_match(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return str(a).strip().lower() == str(b).strip().lower()


def num_close(a, b, tol=0.5):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def to_float(v):
    try:
        if v is None:
            return None
        return float(str(v).replace(",", ".").strip())
    except (TypeError, ValueError):
        return None


def ts_to_date(ts):
    if ts in (None, "", "null"):
        return None
    try:
        return datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return None


def get_expected_data():
    """Читаем moex.stock_info и строим эталон по 5 тикерам.

    Возвращает (dividend_stocks, expected_actions), где:
      dividend_stocks  — set тикеров (.ME), которые СЕЙЧАС платят дивиденды
                         (exDividendDate не пустой);
      expected_actions — list dict с ключами Symbol, Action_Type, Value,
                         Date, Dividend_Yield, Trailing_Annual_Rate.
    """
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute(
        "SELECT symbol, data FROM moex.stock_info WHERE symbol = ANY(%s) ORDER BY symbol",
        (TICKERS,),
    )
    rows = cur.fetchall()
    cur.close()
    conn.close()

    dividend_stocks = set()
    expected_actions = []

    for symbol, data in rows:
        d = data if isinstance(data, dict) else json.loads(data)

        ex_div = d.get("exDividendDate")
        last_val = d.get("lastDividendValue")
        ta_rate = d.get("trailingAnnualDividendRate")
        ta_yield = d.get("trailingAnnualDividendYield")

        # "платит сейчас" <=> есть экс-дивидендная дата
        if ex_div not in (None, "", "null"):
            dividend_stocks.add(symbol)
            expected_actions.append({
                "Symbol": symbol,
                "Action_Type": "Dividend",
                "Value": to_float(last_val),
                "Date": ts_to_date(ex_div),
                "Dividend_Yield": (to_float(ta_yield) or 0.0) * 100.0,
                "Trailing_Annual_Rate": to_float(ta_rate) or 0.0,
            })

        split_factor = d.get("lastSplitFactor")
        split_date = d.get("lastSplitDate")
        if split_factor not in (None, "", "null") and split_date not in (None, "", "null"):
            expected_actions.append({
                "Symbol": symbol,
                "Action_Type": "Stock Split",
                "Value": str(split_factor),
                "Date": ts_to_date(split_date),
                "Dividend_Yield": 0.0,
                "Trailing_Annual_Rate": 0.0,
            })

    return dividend_stocks, expected_actions


def load_sheet_grid():
    """Возвращает (all_values_text, header_map, data_rows) из таблицы агента.

    header_map: {нормализованное_имя_колонки -> col_index}
    data_rows : list[dict] по найденным колонкам, по строкам данных (без шапки)
    Если таблица/лист/ячейки не найдены — возвращает (None, None, None).
    """
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%dividend%' OR LOWER(title) LIKE '%tracker%'
    """)
    spreadsheets = cur.fetchall()
    if not spreadsheets:
        cur.close(); conn.close()
        return None, None, None

    sp_id = spreadsheets[0][0]
    print(f"  Найдена таблица: {spreadsheets[0][1]} (id={sp_id})")

    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (sp_id,))
    sheets = cur.fetchall()
    if not sheets:
        cur.close(); conn.close()
        return "", {}, []

    # предпочитаем лист "Stock Actions", иначе первый
    sheet_id = sheets[0][0]
    for sid, stitle in sheets:
        if str(stitle).strip().lower() == "stock actions":
            sheet_id = sid
            break

    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (sp_id, sheet_id))
    cells = cur.fetchall()
    cur.close(); conn.close()

    all_values = " ".join(str(c[2]).lower() for c in cells if c[2] is not None)
    if not cells:
        return all_values, {}, []

    # сетка row -> {col -> value}
    grid = {}
    for r, c, v in cells:
        grid.setdefault(r, {})[c] = v

    min_row = min(grid.keys())
    header_cells = grid[min_row]
    header_map = {}
    for c, v in header_cells.items():
        if v is None:
            continue
        header_map[str(v).strip().lower().replace(" ", "_")] = c

    data_rows = []
    for r in sorted(grid.keys()):
        if r == min_row:
            continue
        row_cells = grid[r]
        row = {}
        for name, c in header_map.items():
            row[name] = row_cells.get(c)
        if any(val not in (None, "") for val in row.values()):
            data_rows.append(row)

    return all_values, header_map, data_rows


def match_row(data_rows, symbol, action_type):
    """Ищет строку с данным Symbol и Action_Type (база тикера без .ME допустима)."""
    base = symbol.split(".")[0].lower()
    for row in data_rows:
        rs = str(row.get("symbol", "")).strip().lower()
        ra = str(row.get("action_type", "")).strip().lower()
        sym_ok = rs == symbol.lower() or rs == base or rs.split(".")[0] == base
        if sym_ok and ra == action_type.lower():
            return row
    return None


def check_gsheet(dividend_stocks, expected_actions):
    print("\n=== Проверка Google-таблицы ===")
    all_values, header_map, data_rows = load_sheet_grid()

    exists = all_values is not None
    record("Таблица 'Dividend Tracker' / лист 'Stock Actions' существует с данными",
           exists and data_rows is not None and len(data_rows) > 0,
           "Таблица/лист/строки не найдены")
    if not exists or not data_rows:
        return

    # Структурные (не критические): нужные колонки
    required_cols = ["date", "symbol", "action_type", "value",
                     "dividend_yield", "trailing_annual_rate"]
    for col in required_cols:
        record(f"Колонка {col} присутствует", col in header_map,
               f"Шапка: {list(header_map.keys())}")

    # Структурный: упоминание тикеров (база без .ME)
    for base in TICKER_BASES:
        record(f"Таблица содержит тикер {base.upper()}", base in all_values)

    # CRITICAL: дивидендные значения совпадают с источником
    div_values_ok = True
    div_detail = []
    for act in expected_actions:
        if act["Action_Type"] != "Dividend":
            continue
        row = match_row(data_rows, act["Symbol"], "Dividend")
        if row is None:
            div_values_ok = False
            div_detail.append(f"{act['Symbol']}: нет строки Dividend")
            continue
        # Value = lastDividendValue
        if not num_close(to_float(row.get("value")), act["Value"], tol=0.5):
            div_values_ok = False
            div_detail.append(f"{act['Symbol']} Value {row.get('value')}!={act['Value']}")
        # Trailing_Annual_Rate
        if not num_close(to_float(row.get("trailing_annual_rate")),
                         act["Trailing_Annual_Rate"], tol=1.0):
            div_values_ok = False
            div_detail.append(
                f"{act['Symbol']} TAR {row.get('trailing_annual_rate')}!={act['Trailing_Annual_Rate']}")
        # Dividend_Yield: задокументировано "в процентах" (yield*100), но
        # принимаем и сырую дробную форму (yield), т.к. MCP отдаёт дробь.
        got_yield = to_float(row.get("dividend_yield"))
        yield_pct = act["Dividend_Yield"]            # уже *100
        yield_raw = yield_pct / 100.0
        if not (num_close(got_yield, yield_pct, tol=1.0)
                or num_close(got_yield, yield_raw, tol=0.05)):
            div_values_ok = False
            div_detail.append(
                f"{act['Symbol']} Yield {row.get('dividend_yield')}!={yield_pct:.3f}")
        # Date = exDividendDate
        if not str_match(row.get("date"), act["Date"]):
            div_values_ok = False
            div_detail.append(f"{act['Symbol']} Date {row.get('date')}!={act['Date']}")
    record("Дивидендные значения по тикерам совпадают с источником moex",
           div_values_ok, "; ".join(div_detail))

    # CRITICAL: строки дробления
    split_ok = True
    split_detail = []
    for act in expected_actions:
        if act["Action_Type"] != "Stock Split":
            continue
        row = match_row(data_rows, act["Symbol"], "Stock Split")
        if row is None:
            split_ok = False
            split_detail.append(f"{act['Symbol']}: нет строки Stock Split")
            continue
        if not str_match(row.get("value"), act["Value"]):
            split_ok = False
            split_detail.append(f"{act['Symbol']} factor {row.get('value')}!={act['Value']}")
        if not str_match(row.get("date"), act["Date"]):
            split_ok = False
            split_detail.append(f"{act['Symbol']} splitDate {row.get('date')}!={act['Date']}")
    record("Строки дробления акций совпадают с источником moex",
           split_ok, "; ".join(split_detail))

    # Структурный (не критический): сортировка по Symbol, затем Date
    keys = [(str(r.get("symbol", "")).strip().lower(),
             str(r.get("date", "")).strip()) for r in data_rows]
    record("Строки отсортированы по Symbol, затем Date",
           keys == sorted(keys),
           "Порядок строк не соответствует сортировке Symbol/Date")


def check_dividend_set(dividend_stocks, all_values_sheet):
    """CRITICAL: множество плательщиков в данных соответствует источнику.

    Проверяем, что базы тикеров-плательщиков присутствуют, а тикер-неплательщик
    (без exDividendDate) НЕ помечен как дивидендный (нет строки Dividend для него)."""
    print("\n=== Проверка набора плательщиков дивидендов ===")
    _, _, data_rows = load_sheet_grid()
    if data_rows is None:
        record("Набор плательщиков дивидендов совпадает с источником moex",
               False, "Нет строк таблицы")
        return

    paid_in_sheet = set()
    for row in data_rows:
        if str(row.get("action_type", "")).strip().lower() == "dividend":
            rs = str(row.get("symbol", "")).strip().lower().split(".")[0]
            paid_in_sheet.add(rs)

    expected_bases = {s.split(".")[0].lower() for s in dividend_stocks}
    non_div_bases = {t.split(".")[0].lower() for t in TICKERS} - expected_bases

    missing = expected_bases - paid_in_sheet
    extra = paid_in_sheet & non_div_bases
    ok = not missing and not extra
    record("Набор плательщиков дивидендов совпадает с источником moex", ok,
           f"ожидали={sorted(expected_bases)} в_таблице={sorted(paid_in_sheet)} "
           f"лишние={sorted(extra)} пропущены={sorted(missing)}")


def check_email(dividend_stocks):
    print("\n=== Проверка письма ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    cur.close(); conn.close()

    record("Отправлено хотя бы одно письмо", len(emails) >= 1, f"Найдено {len(emails)}")

    target = None
    for subject, from_addr, to_addr, body_text in emails:
        sl = (subject or "").lower()
        if "dividend" in sl or "action" in sl:
            target = (subject, from_addr, to_addr, body_text)
            break

    record("Письмо 'Dividend Action Summary' отправлено инвестору",
           target is not None and
           "investor@portfolio.example.com" in str(target[2] or "").lower(),
           f"Кому: {target[2] if target else None}")

    if target is None:
        record("Письмо отправлено с portfolio-alerts@finance.example.com", False, "нет письма")
        record("Тело письма упоминает плательщика дивидендов", False, "нет письма")
        return

    subject, from_addr, to_addr, body_text = target
    record("Письмо отправлено с portfolio-alerts@finance.example.com",
           "portfolio-alerts@finance.example.com" in str(from_addr or "").lower(),
           f"От: {from_addr}")

    body = (body_text or "").lower()
    # Тело должно упоминать хотя бы один тикер-плательщик (база без .ME)
    paid_bases = {s.split(".")[0].lower() for s in dividend_stocks}
    record("Тело письма упоминает плательщика дивидендов",
           any(b in body for b in paid_bases),
           f"ни один из {sorted(paid_bases)} не найден в теле")
    record("Тело письма упоминает дивиденды",
           "дивиденд" in body or "dividend" in body,
           "нет упоминания дивидендов")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    dividend_stocks, expected_actions = get_expected_data()
    print(f"[eval] Плательщики дивидендов (источник moex): {sorted(dividend_stocks)}")
    print(f"[eval] Ожидаемых действий: {len(expected_actions)}")

    all_values_sheet, _, _ = load_sheet_grid()
    check_gsheet(dividend_stocks, expected_actions)
    check_dividend_set(dividend_stocks, all_values_sheet)
    check_email(dividend_stocks)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    print(f"\nИтого: {PASS_COUNT}/{total} проверок пройдено ({accuracy:.1f}%)")

    if CRITICAL_FAILED:
        print("=== RESULT: FAIL (провалена критическая проверка) ===")
        sys.exit(1)
    if accuracy >= 70:
        print("=== RESULT: PASS ===")
        sys.exit(0)
    print("=== RESULT: FAIL (accuracy < 70%) ===")
    sys.exit(1)


if __name__ == "__main__":
    main()
