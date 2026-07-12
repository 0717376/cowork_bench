"""Generate groundtruth Excel for sf-hr-mentorship-excel-gcal-email (ClickHouse fork).

Mentor/mentee selection is queried LIVE from the russified sf_data DWH so the
groundtruth names/departments are Russian and stay in sync with the central map
(db/zzz_clickhouse_after_init.sql). Requires the DB to be up. The evaluation also
recomputes the expected sets live, so this xlsx is informational/reference.
"""
import os
import openpyxl
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname="cowork_gym", user="eigent", password="camel")

MENTOR_SQL = '''
    SELECT "EMPLOYEE_NAME", "DEPARTMENT", "PERFORMANCE_RATING", "YEARS_EXPERIENCE"
    FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
    WHERE "YEARS_EXPERIENCE" >= 10 AND "PERFORMANCE_RATING" >= 4
    ORDER BY "PERFORMANCE_RATING" DESC, "YEARS_EXPERIENCE" DESC
    LIMIT 10
'''

MENTEE_SQL = '''
    SELECT "EMPLOYEE_NAME", "DEPARTMENT", "PERFORMANCE_RATING", "YEARS_EXPERIENCE"
    FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
    WHERE "YEARS_EXPERIENCE" <= 2 AND "PERFORMANCE_RATING" >= 3
    ORDER BY "PERFORMANCE_RATING" DESC
    LIMIT 10
'''

conn = psycopg2.connect(**DB)
cur = conn.cursor()
cur.execute(MENTOR_SQL)
MENTORS = cur.fetchall()
cur.execute(MENTEE_SQL)
MENTEES = cur.fetchall()
cur.close()
conn.close()

out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "groundtruth_workspace")
os.makedirs(out_dir, exist_ok=True)

wb = openpyxl.Workbook()

ws_pairs = wb.active
ws_pairs.title = "Pairs"
ws_pairs.append(["Mentor_Name", "Mentor_Department", "Mentor_Rating",
                 "Mentee_Name", "Mentee_Department", "Mentee_Experience"])
for i in range(10):
    m = MENTORS[i]
    me = MENTEES[i]
    # m = (name, dept, rating, years); mentee experience = years
    ws_pairs.append([m[0], m[1], m[2], me[0], me[1], me[3]])

ws_summary = wb.create_sheet("Program_Summary")
ws_summary.append(["Metric", "Value"])
total_pairs = 10
avg_mentor_rating = round(sum(m[2] for m in MENTORS) / len(MENTORS), 2)
avg_mentee_exp = round(sum(m[3] for m in MENTEES) / len(MENTEES), 2)
ws_summary.append(["Total_Pairs", total_pairs])
ws_summary.append(["Avg_Mentor_Rating", avg_mentor_rating])
ws_summary.append(["Avg_Mentee_Experience", avg_mentee_exp])

wb.save(os.path.join(out_dir, "Mentorship_Pairs.xlsx"))
print("Groundtruth Excel created.")
print(f"  Total_Pairs: {total_pairs}")
print(f"  Avg_Mentor_Rating: {avg_mentor_rating}")
print(f"  Avg_Mentee_Experience: {avg_mentee_exp}")
print(f"  Mentors: {[m[0] for m in MENTORS]}")
