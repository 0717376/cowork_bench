"""
Preprocess script for canvas-assignment-effectiveness-ppt-teamly-email task.

Canvas is read-only (live English data). This script:
- Clears prior Teamly tracker pages and email data (idempotency).
- Ensures a Teamly space exists for the agent to create the tracker page in.
- Injects RU noise (decoy Teamly pages + decoy emails) — NOT the answer.

We intentionally do NOT pre-create the tracker page, the .xlsx, or the .pptx —
the agent must produce them itself so the evaluation actually tests the agent.
"""
import os
import argparse
import json
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def setup_teamly(cur):
    """Ensure an assessment space exists and clear prior tracker pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; "
              "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return

    # Dedicated space for the agent to drop the tracker page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('ASSESS', 'Учебно-методическая работа',
                'Анализ эффективности заданий, трекеры доработки, отчёты комиссии.')
        ON CONFLICT (key) DO NOTHING
    """)

    # Idempotency: drop any tracker pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%improvement%tracker%'
            OR title ILIKE '%assignment%tracker%'
            OR title ILIKE '%трекер%доработ%'
    """)

    # Inject RU noise pages (decoys, not the answer) into the ASSESS space.
    cur.execute("SELECT id FROM teamly.spaces WHERE key='ASSESS'")
    space_id = cur.fetchone()[0]
    noise_pages = [
        ("Протокол заседания кафедры (март)",
         "# Протокол\n\nОбсудили нагрузку преподавателей и расписание на весенний семестр."),
        ("Заявка на канцелярию",
         "# Заявка\n\nНеобходимо дозаказать бумагу для принтера и картриджи."),
    ]
    for title, body in noise_pages:
        cur.execute("""
            INSERT INTO teamly.pages (space_id, title, body, author)
            VALUES (%s, %s, %s, %s)
        """, (space_id, title, body, "Канцелярия"))
    print("[preprocess] Teamly ready: 'ASSESS' space ensured, prior tracker pages cleared, RU noise injected.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # ── Teamly (ex-Notion) ───────────────────────────────────────────────
        setup_teamly(cur)

        # ── Email ────────────────────────────────────────────────────────────
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        inbox_id = row[0] if row else 3073

        noise_emails = [
            ("Еженедельная сводка кафедры", "manager@university.edu",
             json.dumps(["staff@university.edu"]),
             "Еженедельная сводка статуса работ методического отдела."),
            ("Заявка на канцелярию", "admin@university.edu",
             json.dumps(["procurement@university.edu"]),
             "Нужно дозаказать бумагу для принтера и картриджи."),
        ]
        for subj, from_addr, to_addr, body in noise_emails:
            cur.execute("""
                INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, body_text)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (inbox_id, str(uuid.uuid4()), subj, from_addr, to_addr, body))
        print("[preprocess] Injected RU noise emails.")

        conn.commit()
        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
