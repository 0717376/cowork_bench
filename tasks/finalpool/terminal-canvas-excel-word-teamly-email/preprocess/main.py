"""Preprocess for terminal-canvas-excel-word-teamly-email.

- Ensures a Teamly space exists for the agent to create the "Student Risk
  Tracker" page in, and idempotently clears any prior tracker pages.
- Injects a couple of RU noise pages so the workspace isn't empty.
- Clears email writable schema and injects RU noise emails.
- Canvas is read-only (single source of truth recomputed live in the eval).

We intentionally do NOT pre-create the tracker page, the report files nor the
email — the agent must produce them itself so the evaluation actually tests it.
"""
import argparse
import os
import uuid
import glob as globmod

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}


def setup_teamly(cur):
    """Ensure a Teamly space exists, clear prior tracker pages, inject noise."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the tracker page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('RETENTION', 'Учебный отдел',
                'Аналитика академической успеваемости и профилактика отчисления студентов.')
        ON CONFLICT (key) DO NOTHING
    """)
    cur.execute("SELECT id FROM teamly.spaces WHERE key = 'RETENTION'")
    space_id = cur.fetchone()[0]

    # Idempotency: drop any tracker pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%student risk tracker%'
            OR title ILIKE '%risk tracker%'
            OR title ILIKE '%трекер%риск%'
    """)

    # Inject a couple of RU noise pages (must NOT resemble the answer page).
    noise_pages = [
        ("Архив курсов", "Архивные материалы по прошедшим семестрам. Историческая справка."),
        ("Регламент учебного отдела", "Общие правила и контакты сотрудников учебного отдела вуза."),
    ]
    for title, body in noise_pages:
        cur.execute("SELECT 1 FROM teamly.pages WHERE space_id=%s AND title=%s",
                    (space_id, title))
        if cur.fetchone() is None:
            cur.execute("""
                INSERT INTO teamly.pages (space_id, title, body, author)
                VALUES (%s, %s, %s, %s)
            """, (space_id, title, body, "Учебный отдел"))
    print("[preprocess] Teamly ready: 'RETENTION' space ensured, prior tracker pages cleared, noise injected.")


def setup_email(cur):
    """Clear email writable schema and inject RU noise emails."""
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    print("[preprocess] Cleared email data.")

    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if row:
        folder_id = row[0]
        cur.execute("SELECT COALESCE(MAX(id), 0) FROM email.messages")
        max_id = cur.fetchone()[0]
        for i in range(2):
            max_id += 1
            cur.execute("""
                INSERT INTO email.messages (id, folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                VALUES (%s, %s, %s, %s, %s, %s::jsonb, NOW(), %s, false)
            """, (max_id, folder_id, f"noise-{uuid.uuid4()}@example.com",
                  f"Еженедельная рассылка №{i+1}", "newsletter@university.edu",
                  '["advisor@university.edu"]',
                  f"Это шумовое содержимое письма №{i+1}."))
        print("[preprocess] Injected noise email data.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        setup_teamly(cur)
        setup_email(cur)
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean workspace deliverables left from prior runs.
    if args.agent_workspace:
        for pattern in ["Student_Risk_Analysis.xlsx", "Intervention_Plan.docx", "risk_scorer.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
