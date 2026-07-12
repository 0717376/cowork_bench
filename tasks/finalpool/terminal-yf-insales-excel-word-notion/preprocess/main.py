"""Preprocess для terminal-yf-insales-excel-word-notion (RU: moex + teamly).

Котировки moex.* и товары wc.* засеяны глобально и доступны только для чтения —
их не трогаем. Здесь только:
  - чистим рабочие страницы teamly (id > 3, созданные агентом/прошлым прогоном);
  - инжектим RU-шумовые страницы teamly (несвязанные с рыночным дашбордом);
  - удаляем артефакты из рабочего каталога агента.
Идемпотентно: повторный запуск не плодит дубликаты и не пре-сеет ответ.
"""
import argparse
import glob as globmod
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}

# Шумовые страницы Teamly (RU): не относятся к рыночному дашборду.
NOISE_PAGES = [
    ("Старый каталог товаров", "Устаревший список товаров. Не использовать для анализа."),
    ("Архив товара 1", "Архивная карточка снятого с продажи товара."),
    ("Архив товара 2", "Архивная карточка снятого с продажи товара."),
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # --- Teamly: убираем страницы агента/прошлого прогона (id > 3) ---
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        print("[preprocess] Очищены рабочие страницы teamly (id>3).")

        # --- Инжект RU-шумового пространства и страниц ---
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('OLDCAT', 'Старый каталог', 'Устаревшие материалы по товарам.')
            ON CONFLICT (key) DO NOTHING
        """)
        cur.execute("SELECT id FROM teamly.spaces WHERE key='OLDCAT'")
        space_id = cur.fetchone()[0]
        for title, body in NOISE_PAGES:
            cur.execute("SELECT 1 FROM teamly.pages WHERE space_id=%s AND title=%s",
                        (space_id, title))
            if cur.fetchone() is None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s,%s,%s,%s)",
                    (space_id, title, body, "admin"))
        print("[preprocess] Инжектированы RU-шумовые страницы teamly.")

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Commodity_Impact_Analysis.xlsx", "Pricing_Strategy_Memo.docx", "correlation_analysis.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
