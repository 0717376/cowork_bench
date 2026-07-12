import argparse
import json
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Clear email schema
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")

        # Ensure folders exist
        cur.execute("DELETE FROM email.folders")
        cur.execute(
            "INSERT INTO email.folders (id, name) VALUES (1, 'INBOX'), (2, 'Sent'), (3, 'Drafts')"
        )

        # Inject noise emails
        noise_emails = [
            {
                "folder_id": 1,
                "message_id": str(uuid.uuid4()),
                "subject": "Еженедельное обновление по складу магазина",
                "from_addr": "inventory@university.edu",
                "to_addr": json.dumps(["bookstore_manager@university.edu"]),
                "date": "2025-11-01 09:00:00",
                "body_text": "Отчёт по складу за эту неделю показывает нормальный уровень запасов по всем отделам.",
                "is_read": True,
            },
            {
                "folder_id": 1,
                "message_id": str(uuid.uuid4()),
                "subject": "Кампусное мероприятие: техническая ярмарка в следующем месяце",
                "from_addr": "events@university.edu",
                "to_addr": json.dumps(["all_staff@university.edu"]),
                "date": "2025-10-28 14:30:00",
                "body_text": "Приглашаем на ежегодную техническую ярмарку со студенческими проектами и новой электроникой.",
                "is_read": True,
            },
            {
                "folder_id": 1,
                "message_id": str(uuid.uuid4()),
                "subject": "Re: Заказ учебников на весенний семестр",
                "from_addr": "academic_affairs@university.edu",
                "to_addr": json.dumps(["bookstore_manager@university.edu"]),
                "date": "2025-10-20 11:15:00",
                "body_text": "Пожалуйста, подтвердите список заказа учебников для курсов предстоящего весеннего семестра.",
                "is_read": False,
            },
            {
                "folder_id": 2,
                "message_id": str(uuid.uuid4()),
                "subject": "Запрос цен поставщику: партия электроники",
                "from_addr": "bookstore_manager@university.edu",
                "to_addr": json.dumps(["supplier@techvendor.com"]),
                "date": "2025-10-15 16:00:00",
                "body_text": "Запрашиваем обновлённые цены на нашу следующую партию электроники для склада.",
                "is_read": True,
            },
            {
                "folder_id": 1,
                "message_id": str(uuid.uuid4()),
                "subject": "Предложение по программе студенческих скидок",
                "from_addr": "marketing@university.edu",
                "to_addr": json.dumps(["bookstore_manager@university.edu"]),
                "date": "2025-10-10 08:45:00",
                "body_text": "Во вложении черновик предложения по программе студенческих скидок. Пожалуйста, ознакомьтесь.",
                "is_read": True,
            },
        ]

        for e in noise_emails:
            cur.execute(
                """INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                VALUES (%(folder_id)s, %(message_id)s, %(subject)s, %(from_addr)s,
                        %(to_addr)s::jsonb, %(date)s, %(body_text)s, %(is_read)s)""",
                e,
            )

        conn.commit()
        print("Preprocess completed: email schema cleared and noise emails injected.")
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
