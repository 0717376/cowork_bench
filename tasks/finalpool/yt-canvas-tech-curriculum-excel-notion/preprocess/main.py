"""
Preprocess for yt-canvas-tech-curriculum-excel-notion task.

Clears the teamly deliverable page so the agent starts fresh.
Canvas data (canvas schema) is READ-ONLY. YouTube video/title content
(youtube.videos) is READ-ONLY; only the youtube.channels.video_count metadata
is repaired here (it ships seeded to 0, which breaks channels_listVideos /
navigateList pagination so the documented "list all channel videos" workflow
sees only page 1). video_count is set deterministically to the real per-channel
COUNT(*) FROM youtube.videos so the intended full-channel pagination workflow
becomes reachable. No video rows, titles or view counts are altered.

Prerequisites:
  - PostgreSQL cowork_gym database running on localhost:5432
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


def clear_tables(conn):
    with conn.cursor() as cur:
        # Remove any prior 'Tech Course Video Resources' deliverable pages so the
        # agent starts fresh. Other teamly pages/spaces are left intact.
        cur.execute("""
            DELETE FROM teamly.pages
            WHERE title ILIKE '%tech course video%'
               OR title ILIKE '%course video resource%'
        """)
    conn.commit()
    print("[preprocess] Cleared prior teamly deliverable pages.")


def repair_channel_video_count(conn):
    """Set youtube.channels.video_count to the real per-channel video row count.

    The base seed ships video_count=0 for every channel, which makes
    channels_listVideos compute totalPages=ceil(0/20)=0 and hasNextPage=false
    (navigateList likewise). The intended "list all channel videos" workflow then
    only ever sees the first page (20 newest videos), so the top-by-views
    keyword-matching videos required by the groundtruth are unreachable. This is
    pure metadata repaired deterministically from the (read-only) video rows.
    """
    with conn.cursor() as cur:
        cur.execute("""
            UPDATE youtube.channels c
            SET video_count = sub.cnt
            FROM (
                SELECT channel_id, COUNT(*) AS cnt
                FROM youtube.videos
                GROUP BY channel_id
            ) sub
            WHERE c.channel_id = sub.channel_id
        """)
        updated = cur.rowcount
    conn.commit()
    print(f"[preprocess] Repaired youtube.channels.video_count for {updated} channels.")


def verify_readonly_data(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) FROM youtube.videos WHERE channel_title = 'Fireship'")
        yt_count = cur.fetchone()[0]
        cur.execute("""
            SELECT COUNT(*) FROM canvas.courses
            WHERE name ILIKE '%вычислен%' OR name ILIKE '%данны%'
        """)
        course_count = cur.fetchone()[0]
    print(f"[preprocess] Fireship videos: {yt_count} (read-only)")
    print(f"[preprocess] Matching courses: {course_count} (read-only)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        repair_channel_video_count(conn)
        verify_readonly_data(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
