"""Preprocess for terminal-sf-hr-diversity-gform-excel-notion.
ClickHouse (sf_data) HR data is read-only. Clear forms and teamly, inject noise data."""
import argparse
import glob
import os
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_gform(cur):
    print("[preprocess] Clearing Forms data...")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")


def clear_teamly(cur):
    """Drop user-created teamly pages (seed pages have id <= 3) and ensure a space.
    Teamly has no database object; the agent creates one page per department."""
    print("[preprocess] Clearing Teamly data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('HRANALYTICS', 'Кадровая аналитика',
                    'Пространство команды People Analytics для оценки разнообразия персонала.')
            ON CONFLICT (key) DO NOTHING
        """)


def inject_noise(cur):
    print("[preprocess] Injecting noise data...")
    # Noise form (RU title, pure distractor — not the diversity survey).
    noise_form_id = str(uuid.uuid4())
    cur.execute("""
        INSERT INTO gform.forms (id, title, description)
        VALUES (%s, 'Опрос удовлетворённости сотрудников Q4', 'Ежеквартальная проверка удовлетворённости')
    """, (noise_form_id,))
    for i, (title, qtype) in enumerate([
        ("Насколько вы довольны своей ролью?", "MULTIPLE_CHOICE"),
        ("Оцените баланс между работой и личной жизнью", "MULTIPLE_CHOICE"),
        ("Дополнительные комментарии?", "TEXT"),
    ]):
        cur.execute("""
            INSERT INTO gform.questions (form_id, title, question_type, required, position)
            VALUES (%s, %s, %s, false, %s)
        """, (noise_form_id, title, qtype, i))

    # Noise teamly pages (RU titles) in the HRANALYTICS space — leftover content
    # the agent must ignore; must NOT satisfy the dashboard page checks.
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("SELECT id FROM teamly.spaces WHERE key = 'HRANALYTICS'")
        row = cur.fetchone()
        if row is None:
            cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
            row = cur.fetchone()
        space_id = row[0] if row else None
        if space_id is not None:
            for title, body in [
                ("Архив протоколов совещаний", "Старые заметки со встреч команды. Не относится к текущей задаче."),
                ("Трекер проектов Q4", "Редизайн сайта, миграция API, конвейер данных. Не относится к разнообразию."),
            ]:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) "
                    "VALUES (%s, %s, %s, %s)",
                    (space_id, title, body, "team"),
                )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gform(cur)
        clear_teamly(cur)
        inject_noise(cur)
        conn.commit()
        print("[preprocess] DB setup done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Diversity_Metrics_Report.xlsx", "diversity_*.py", "diversity_*.json", "dept_*.json"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
