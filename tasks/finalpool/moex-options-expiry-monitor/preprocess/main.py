"""Предобработка для задачи мониторинга экспирации опционов (MOEX).

Действия:
  1. Очищает события Google Calendar (идемпотентно).
  2. Засевает детерминированные опционные цепочки в схему moex.options
     для пяти эмитентов Московской биржи (источник данных, который читает агент).
  3. Запускает локальный mock-портал мониторинга на порту 30227.

Таблица moex.options изначально пуста, а оригинальная задача обращалась к
живым данным опционов. Поэтому здесь мы детерминированно заполняем цепочки
(несколько дат экспирации на тикер, calls+puts, со страйками и implied
volatility), чтобы аналитика в Excel/календаре была воспроизводимой.
Это ИСТОЧНИК данных, а не ответ задачи — агент сам считает агрегаты.
"""
import argparse
import asyncio
import json
import os
import shutil

import psycopg2

PORT = 30227
DB = dict(
    host=os.environ.get("PGHOST", "localhost"),
    port=int(os.environ.get("PGPORT", "5432")),
    dbname="cowork_gym",
    user="eigent",
    password="camel",
)

# Фиксированная «текущая дата» = 2026-03-07, порог near-expiry = 2026-03-14 (7 дней).
# Текущие цены берём из moex.stock_info: SBER 133.3, GAZP 198.0, LKOH 3911.0,
# MGNT 4439.0, MTSS 275.05.
#
# Структура: тикер -> цена -> {дата_экспирации -> {'calls': [IV...], 'puts': [IV...]}}.
# IV задаются как доля (impliedVolatility), Avg_IV% = round(mean*100, 1).
SEED_PLAN = {
    "SBER.ME": {
        "price": 133.3,
        "exps": {
            "2026-03-10": {"calls": [0.55, 0.58, 0.61], "puts": [0.50, 0.52, 0.54]},
            "2026-03-13": {"calls": [0.30, 0.32], "puts": [0.28, 0.31]},
            "2026-04-17": {"calls": [0.22, 0.24, 0.26], "puts": [0.20, 0.23]},
        },
    },
    "GAZP.ME": {
        "price": 198.0,
        "exps": {
            "2026-03-12": {"calls": [0.45, 0.47], "puts": [0.41, 0.43]},
            "2026-05-15": {"calls": [0.33, 0.35, 0.30], "puts": [0.29, 0.31]},
        },
    },
    "LKOH.ME": {
        "price": 3911.0,
        "exps": {
            "2026-03-14": {"calls": [0.38, 0.36], "puts": [0.35, 0.37, 0.39]},
            "2026-06-19": {"calls": [0.28, 0.30], "puts": [0.26, 0.27]},
        },
    },
    "MGNT.ME": {
        "price": 4439.0,
        "exps": {
            "2026-03-11": {"calls": [0.62, 0.65, 0.60], "puts": [0.58, 0.61]},
            "2026-04-17": {"calls": [0.42, 0.44], "puts": [0.40, 0.45, 0.43]},
        },
    },
    "MTSS.ME": {
        "price": 275.05,
        "exps": {
            "2026-03-09": {"calls": [0.33, 0.35], "puts": [0.31, 0.34, 0.30]},
            "2026-07-17": {"calls": [0.25, 0.27, 0.29], "puts": [0.24, 0.26]},
        },
    },
}


def _build_contracts(symbol, exp, option_type, ivs, price):
    """Сформировать список contract-record в формате, ожидаемом адаптером
    (pandas DataFrame со столбцами strike + impliedVolatility и пр.)."""
    contracts = []
    n = len(ivs)
    # Страйки вокруг текущей цены, детерминированный шаг ~5%.
    step = round(price * 0.05, 2)
    base = round(price - step * (n // 2), 2)
    exp_compact = exp.replace("-", "")[2:]  # YYMMDD
    letter = "C" if option_type == "calls" else "P"
    for i, iv in enumerate(ivs):
        strike = round(base + step * i, 2)
        contracts.append(
            {
                "index": i,
                "contractSymbol": f"{symbol.split('.')[0]}{exp_compact}{letter}{int(strike*1000):08d}",
                "strike": strike,
                "lastPrice": round(abs(price - strike) * 0.1 + 1.0, 2),
                "bid": round(abs(price - strike) * 0.1, 2),
                "ask": round(abs(price - strike) * 0.1 + 0.5, 2),
                "change": 0.0,
                "percentChange": 0.0,
                "volume": 100 + i * 10,
                "openInterest": 50 + i * 5,
                "impliedVolatility": iv,
                "inTheMoney": (strike < price) if option_type == "calls" else (strike > price),
                "contractSize": "REGULAR",
                "currency": "RUB",
                "lastTradeDate": "2026-03-05T15:00:00+00:00",
            }
        )
    return contracts


def seed_options():
    conn = psycopg2.connect(**DB)
    conn.autocommit = True
    cur = conn.cursor()
    # Идемпотентно: удалить прежние строки для наших тикеров, затем вставить заново.
    syms = tuple(SEED_PLAN.keys())
    cur.execute("DELETE FROM moex.options WHERE symbol IN %s", (syms,))
    count = 0
    for symbol, info in SEED_PLAN.items():
        price = info["price"]
        for exp, types in info["exps"].items():
            for option_type in ("calls", "puts"):
                contracts = _build_contracts(symbol, exp, option_type, types[option_type], price)
                cur.execute(
                    "INSERT INTO moex.options(symbol, expiration_date, option_type, data) "
                    "VALUES (%s, %s, %s, %s)",
                    (symbol, exp, option_type, json.dumps(contracts)),
                )
                count += 1
    cur.close()
    conn.close()
    print(f"[preprocess] Seeded {count} option chain groups into moex.options.")


def clear_gcal():
    conn = psycopg2.connect(**DB)
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute("DELETE FROM gcal.events")
    cur.close()
    conn.close()
    print("[preprocess] Cleared gcal events.")


async def run_command(cmd):
    proc = await asyncio.create_subprocess_shell(
        cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    await proc.wait()


async def setup_mock_server():
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    mock_dir = os.path.join(task_root, "files", "mock_pages")
    tmp_dir = os.path.join(task_root, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    await run_command(f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null")
    await asyncio.sleep(0.5)
    log_path = os.path.join(tmp_dir, "server.log")
    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {PORT} --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{PORT}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_gcal()
    seed_options()
    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
