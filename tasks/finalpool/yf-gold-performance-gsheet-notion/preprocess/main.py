"""Preprocess для yf-gold-performance-gsheet-notion (RU-стек: moex-finance / teamly).

moex-finance доступен только на чтение. Этот скрипт очищает записываемые схемы:
- gsheet (полностью);
- teamly: удаляет пользовательские страницы (id > 3, сидовые страницы сохраняются)
  и гарантирует наличие пространства FINANCE, куда агент поместит итоговую страницу.
"""

import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_gsheet(cur):
    print("[preprocess] Очистка данных Google Sheets...")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.permissions")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.spreadsheets")
    print("[preprocess] Данные Google Sheets очищены.")


def setup_teamly(cur):
    print("[preprocess] Подготовка teamly...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: схема teamly не найдена; пропуск. "
              "Примените db/zzz_teamly_after_init.sql.")
        return
    # Сидовые страницы имеют id 1..3 — сохраняем их. Удаляем пользовательские.
    cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    # Гарантируем наличие пространства FINANCE для итоговой страницы.
    cur.execute(
        """INSERT INTO teamly.spaces (key, name, description)
           VALUES ('FINANCE', 'Финансовая аналитика',
                   'Рыночная аналитика и отчёты по инструментам MOEX')
           ON CONFLICT (key) DO NOTHING"""
    )
    print("[preprocess] teamly подготовлен (пользовательские страницы очищены, "
          "пространство FINANCE на месте).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = True
    cur = conn.cursor()

    try:
        clear_gsheet(cur)
        setup_teamly(cur)
        print("[preprocess] Готово. Записываемые схемы очищены.")
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
