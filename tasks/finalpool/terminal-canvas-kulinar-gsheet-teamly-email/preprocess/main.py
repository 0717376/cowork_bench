"""
Preprocess for terminal-canvas-kulinar-gsheet-teamly-email task.

Clears gsheet, teamly (Wellness Pilot pages), email schemas. Injects noise data.
Ensures a Teamly space exists for the agent to create the program page in.
Canvas and Kulinar are read-only (globally seeded).
"""
import argparse
import os
import uuid
from datetime import datetime, timedelta

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_schemas(conn):
    with conn.cursor() as cur:
        # Clear gsheet
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.spreadsheets")
        # Clear email
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
    conn.commit()
    print("[preprocess] Cleared gsheet, email schemas.")


def setup_teamly(conn):
    """Ensure a Teamly space exists for the Wellness Pilot page; clear prior runs."""
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
            return
        # Dedicated space for the agent to drop the program page into.
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('WELLNESS', 'Студенческое благополучие',
                    'Программы питания и вовлечённости студентов отдела по работе со студентами.')
            ON CONFLICT (key) DO NOTHING
        """)
        # Idempotency: drop any program pages left from previous runs.
        cur.execute("""
            DELETE FROM teamly.pages
             WHERE title ILIKE '%wellness pilot%'
                OR title ILIKE '%пилот%благополуч%'
        """)
    conn.commit()
    print("[preprocess] Teamly ready: 'WELLNESS' space ensured, prior program pages cleared.")


def inject_noise_gsheet(conn):
    """Inject a noise spreadsheet the agent should ignore."""
    with conn.cursor() as cur:
        ss_id = str(uuid.uuid4())
        cur.execute(
            "INSERT INTO gsheet.spreadsheets (id, title) VALUES (%s, %s)",
            (ss_id, "Budget Tracking Q1 2026"),
        )
        cur.execute(
            "INSERT INTO gsheet.sheets (spreadsheet_id, title, index, row_count, column_count) "
            "VALUES (%s, %s, 0, 10, 4) RETURNING id",
            (ss_id, "Q1_Budget"),
        )
        sheet_id = cur.fetchone()[0]
        for i, (label, val) in enumerate(
            [("Department", "Marketing"), ("Budget", "45000"), ("Spent", "31200"), ("Remaining", "13800")]
        ):
            cur.execute(
                "INSERT INTO gsheet.cells (spreadsheet_id, sheet_id, row_index, col_index, value) "
                "VALUES (%s, %s, 0, %s, %s)",
                (ss_id, sheet_id, i, label),
            )
            cur.execute(
                "INSERT INTO gsheet.cells (spreadsheet_id, sheet_id, row_index, col_index, value) "
                "VALUES (%s, %s, 1, %s, %s)",
                (ss_id, sheet_id, i, val),
            )
    conn.commit()
    print("[preprocess] Injected noise gsheet data.")


def inject_noise_teamly(conn):
    """Inject a noise Teamly page the agent should ignore."""
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            return
        cur.execute("SELECT id FROM teamly.spaces WHERE key = 'WELLNESS'")
        row = cur.fetchone()
        if not row:
            return
        space_id = row[0]
        cur.execute(
            "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
            (
                space_id,
                "Проекты департамента",
                "# Проекты департамента\n\nВнутренний трекер проектов. Текущий статус: Открыт.\n"
                "Редизайн сайта — в работе. Не относится к пилоту по питанию.",
                "admin@university.edu",
            ),
        )
    conn.commit()
    print("[preprocess] Injected noise teamly data.")


def inject_noise_email(conn, launch):
    """Inject noise emails the agent should ignore."""

    def dt(days, hours, minutes=0):
        return (launch + timedelta(days=days, hours=hours, minutes=minutes)).strftime("%Y-%m-%d %H:%M:%S")

    with conn.cursor() as cur:
        cur.execute(
            """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
            VALUES
            ((SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1),
             'noise-msg-001', 'Совещание по бюджету Q1', 'finance@university.edu',
             '["admin@university.edu"]', %s, 'Просьба ознакомиться с приложённым бюджетным отчётом Q1 до совещания на следующей неделе.', true),
            ((SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1),
             'noise-msg-002', 'Обновление по парковке кампуса', 'facilities@university.edu',
             '["all_staff@university.edu"]', %s, 'С 15 марта парковка B закрывается на ремонт покрытия. Пользуйтесь парковкой C.', true),
            ((SELECT id FROM email.folders WHERE name = 'Sent' LIMIT 1),
             'noise-msg-003', 'Re: Планирование выездного семинара факультета', 'dean@university.edu',
             '["faculty_committee@university.edu"]', %s, 'Семинар подтверждён на 20 апреля. Просьба прислать заявки на сессии до 25 марта.', true)
            """,
            (dt(-6, -1), dt(-4, 4, 30), dt(-2, 1)),
        )
    conn.commit()
    print("[preprocess] Injected noise email data.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    launch = datetime.strptime(args.launch_time or "2026-03-07 10:00:00", "%Y-%m-%d %H:%M:%S")

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_schemas(conn)
        setup_teamly(conn)
        inject_noise_gsheet(conn)
        inject_noise_teamly(conn)
        inject_noise_email(conn, launch)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
