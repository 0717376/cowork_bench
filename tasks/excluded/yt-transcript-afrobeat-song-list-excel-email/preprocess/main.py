"""
Препроцессинг для задачи yt-transcript-afrobeat-song-list-excel-email.

Очищает почту и страницы Teamly, созданные агентом.
Внедряет письмо от music@label.com с просьбой проанализировать Afrobeat-трек-лист.

Требования:
  - PostgreSQL cowork_gym, запущенный на localhost:5432
"""
import argparse
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


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
        # Сидовые страницы teamly (id <= 63 из db-инициализации). Всё, что создал
        # агент в прошлых прогонах, удаляем.
        cur.execute("DELETE FROM teamly.page_labels WHERE page_id > 63")
        cur.execute("DELETE FROM teamly.pages WHERE id > 63")
    conn.commit()
    print("[preprocess] Очищены email и созданные агентом страницы Teamly.")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        conn.commit()
        return cur.fetchone()[0]


def inject_email(conn, folder_id):
    with conn.cursor() as cur:
        msg_id = str(uuid.uuid4())
        cur.execute("""
            INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
            VALUES (%s, %s, %s, %s, %s, NOW(), %s)
        """, (
            folder_id,
            msg_id,
            "Request: Afrobeat Mix Tracklist Analysis",
            "music@label.com",
            '["curator@musicteam.com"]',
            "Hi,\n\nWe need a full tracklist analysis of the Afrobeat Mix 2024 video "
            "(video ID: 7ZQzGq32kAY). Please identify all songs and artists featured, "
            "organize them into a spreadsheet, write up your curator notes, publish the "
            "tracklist to our team wiki, and email me the final results.\n\n"
            "We need the Excel file Afrobeat_Tracklist.xlsx with a Tracklist sheet and "
            "an Artist_Summary sheet, plus a Word document Curator_Notes.docx.\n\n"
            "Thanks,\nMusic Label Team"
        ))
    conn.commit()
    print("[preprocess] Внедрено письмо-запрос от music@label.com.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        folder_id = ensure_email_folder(conn)
        inject_email(conn, folder_id)
    finally:
        conn.close()

    print("\n[preprocess] Препроцессинг успешно завершён!")


if __name__ == "__main__":
    main()
