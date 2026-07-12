"""
Evaluation для задачи scholarly-reading-group-gcal-gsheet-word.
Проверяет: GSheet (расписание семинара), GCal (события), Word (документ).

Модель оценки:
  - CRITICAL_CHECKS — семантические проверки сути задачи. Любой провал
    критической проверки => общий FAIL независимо от accuracy.
  - Иначе PASS требует accuracy >= 70%.

Прозаический текст (вводный абзац, краткие резюме, обозначение недель) агент
пишет по-русски, поэтому совпадения принимаются и на русском, и на английском.
Идентификаторы (названия статей, arxiv ID, имена авторов, имя файла,
названия событий, заголовки столбцов) остаются английскими — eval грепает их
как английские подстроки.
"""
import argparse
import json
import os
import sys

import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Критические проверки: любой провал => общий FAIL независимо от accuracy.
CRITICAL_CHECKS = {
    # Шумовая RLHF-статья ДОЛЖНА быть отброшена — ядро отборочного суждения.
    "GSheet НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)",
    "Word НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)",
    # Все 3 релевантные статьи присутствуют в GSheet.
    "GSheet содержит запись 'Attention Is All You Need'",
    "GSheet содержит запись BERT",
    "GSheet содержит запись few-shot/GPT",
    # Конкретный результат планирования.
    "GCal: событие 16 марта 2026",
    "GCal: событие 23 марта 2026",
    "GCal: событие 30 марта 2026",
    "GCal: события запланированы на 15:00 (UTC)",
    # Word отображает по одной отдельной статье на неделю.
    "Word: каждая из 3 недель сопоставлена своей отдельной статье",
    # Word упоминает >= 2 статей с подтверждением авторами + вводный абзац.
    "Word упоминает минимум 2 из 3 статей о трансформерах",
    "Word содержит вводный абзац с критериями отбора",
}

# Маркеры шумовой RLHF-статьи, которой быть НЕ должно.
NOISE_MARKERS = ["ouyang", "2203.02155", "instructions with human feedback"]


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        target_ss = None
        for sid, title in spreadsheets:
            if title and ("transformer" in title.lower() or "reading group" in title.lower()):
                target_ss = sid
                break

        record("GSheet 'Transformer Reading Group Schedule' exists",
               target_ss is not None,
               f"Found sheets: {[t for _, t in spreadsheets]}")

        if target_ss is None:
            # Критические content-проверки не могут пройти без таблицы — отметим провал.
            record("GSheet содержит запись 'Attention Is All You Need'", False, "no spreadsheet")
            record("GSheet содержит запись BERT", False, "no spreadsheet")
            record("GSheet содержит запись few-shot/GPT", False, "no spreadsheet")
            record("GSheet НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)", False, "no spreadsheet")
            conn.close()
            return

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (target_ss,))
        sheets = cur.fetchall()
        if not sheets:
            record("GSheet has at least one sheet", False)
            record("GSheet содержит запись 'Attention Is All You Need'", False, "no sheet")
            record("GSheet содержит запись BERT", False, "no sheet")
            record("GSheet содержит запись few-shot/GPT", False, "no sheet")
            record("GSheet НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)", False, "no sheet")
            conn.close()
            return

        sheet_id = sheets[0][0]
        cur.execute("""
            SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 0
        """, (target_ss, sheet_id))
        data_rows = cur.fetchone()[0]
        record("GSheet has at least 3 paper rows (one per week)", data_rows >= 3,
               f"Found {data_rows} data rows")

        # Содержимое статей в ячейках.
        cur.execute("""
            SELECT LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
        """, (target_ss, sheet_id))
        cell_values = [row[0] for row in cur.fetchall() if row[0]]
        all_text = " ".join(cell_values)

        # Ужесточено: требуем точное название, а не голую подстроку 'attention'.
        has_attention = "attention is all you need" in all_text
        has_bert = "bert" in all_text
        has_gpt = ("few-shot" in all_text or "few shot" in all_text
                   or "language models are" in all_text)
        record("GSheet содержит запись 'Attention Is All You Need'", has_attention)
        record("GSheet содержит запись BERT", has_bert)
        record("GSheet содержит запись few-shot/GPT", has_gpt)

        # Критично: шумовая RLHF-статья НЕ должна попасть в таблицу.
        noise_hit = next((m for m in NOISE_MARKERS if m in all_text), None)
        record("GSheet НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)",
               noise_hit is None, f"Найден шумовой маркер: {noise_hit}")

        conn.close()
    except Exception as e:
        record("GSheet connection", False, str(e))


def check_gcal():
    print("\n=== Checking Google Calendar ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("""
            SELECT summary, start_datetime FROM gcal.events
            WHERE LOWER(summary) LIKE '%reading group%' OR LOWER(summary) LIKE '%transformer%'
            ORDER BY start_datetime
        """)
        events = cur.fetchall()
        record("GCal has at least 3 Reading Group events", len(events) >= 3,
               f"Found {len(events)} events: {[(e[0], str(e[1])[:10]) for e in events]}")

        dates = [str(e[1])[:10] for e in events]
        has_march16 = any("2026-03-16" in d for d in dates)
        has_march23 = any("2026-03-23" in d for d in dates)
        has_march30 = any("2026-03-30" in d for d in dates)
        record("GCal: событие 16 марта 2026", has_march16, f"Dates found: {dates}")
        record("GCal: событие 23 марта 2026", has_march23, f"Dates found: {dates}")
        record("GCal: событие 30 марта 2026", has_march30, f"Dates found: {dates}")

        cur.execute("""
            SELECT EXTRACT(HOUR FROM start_datetime AT TIME ZONE 'UTC') as utc_hour
            FROM gcal.events
            WHERE LOWER(summary) LIKE '%reading group%' OR LOWER(summary) LIKE '%transformer%'
        """)
        hours = [int(row[0]) for row in cur.fetchall()]
        has_correct_time = any(h == 15 for h in hours)
        record("GCal: события запланированы на 15:00 (UTC)", has_correct_time,
               f"UTC hours found: {hours}")

        conn.close()
    except Exception as e:
        record("GCal connection", False, str(e))
        record("GCal: событие 16 марта 2026", False, str(e))
        record("GCal: событие 23 марта 2026", False, str(e))
        record("GCal: событие 30 марта 2026", False, str(e))
        record("GCal: события запланированы на 15:00 (UTC)", False, str(e))


def check_word(agent_workspace):
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "Transformer_Reading_List.docx")
    if not os.path.isfile(doc_path):
        record("Word file Transformer_Reading_List.docx exists", False, f"Not found at: {doc_path}")
        # Зависимые критические проверки не могут пройти без файла.
        record("Word: каждая из 3 недель сопоставлена своей отдельной статье", False, "no file")
        record("Word упоминает минимум 2 из 3 статей о трансформерах", False, "no file")
        record("Word содержит вводный абзац с критериями отбора", False, "no file")
        record("Word НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)", False, "no file")
        return
    record("Word file Transformer_Reading_List.docx exists", True)

    try:
        doc = Document(doc_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        record("Word: каждая из 3 недель сопоставлена своей отдельной статье", False, "unreadable")
        record("Word упоминает минимум 2 из 3 статей о трансформерах", False, "unreadable")
        record("Word содержит вводный абзац с критериями отбора", False, "unreadable")
        record("Word НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)", False, "unreadable")
        return
    record("Word file readable", True)

    paragraphs = [p.text for p in doc.paragraphs]
    full_text = "\n".join(paragraphs).lower()

    # Заголовок: 'transformer' + контекст семинара (англ. или рус.).
    has_heading = "transformer" in full_text and (
        "reading group" in full_text or "reading list" in full_text
        or "architecture" in full_text or "архитектур" in full_text
        or "семинар" in full_text)
    record("Word has 'Transformer' heading with reading group context", has_heading)

    has_content = len(full_text) > 300
    record("Word has substantial content", has_content, f"Text length: {len(full_text)}")

    # Вводный абзац с критериями отбора (англ. или рус.).
    has_intro = ("criteria" in full_text or "criterion" in full_text
                 or "selected" in full_text or "selection" in full_text
                 or "критери" in full_text or "отбор" in full_text
                 or "выбран" in full_text or "выбор" in full_text)
    record("Word содержит вводный абзац с критериями отбора", has_intro)

    has_attention = "attention is all you need" in full_text or (
        "attention" in full_text and "vaswani" in full_text)
    has_bert = "bert" in full_text and ("devlin" in full_text or "bidirectional" in full_text)
    has_fewshot = ("few-shot" in full_text or "few shot" in full_text
                   or "gpt-3" in full_text or "language models are" in full_text)

    papers_mentioned = sum([has_attention, has_bert, has_fewshot])
    record("Word mentions 'Attention Is All You Need'", has_attention)
    record("Word mentions BERT", has_bert)
    record("Word mentions few-shot learners (GPT-3)", has_fewshot)
    record("Word упоминает минимум 2 из 3 статей о трансформерах", papers_mentioned >= 2,
           f"Found {papers_mentioned}/3 papers")

    # Критично: каждая из 3 недель сопоставлена своей ОТДЕЛЬНОЙ статье.
    # Сопоставляем неделю -> статью по абзацу/строке (англ. 'week N' или рус. 'неделя N').
    week_tokens = {
        1: ("week 1", "week1", "неделя 1", "неделя1"),
        2: ("week 2", "week2", "неделя 2", "неделя2"),
        3: ("week 3", "week3", "неделя 3", "неделя3"),
    }
    paper_signatures = {
        "attention": lambda t: "attention is all you need" in t or "vaswani" in t,
        "bert": lambda t: "bert" in t or "devlin" in t,
        "fewshot": lambda t: ("few-shot" in t or "few shot" in t
                              or "gpt-3" in t or "language models are" in t or "brown" in t),
    }
    week_to_paper = {}
    for para in paragraphs:
        pt = para.lower()
        for wk, toks in week_tokens.items():
            if any(tok in pt for tok in toks):
                for pname, matcher in paper_signatures.items():
                    if matcher(pt):
                        week_to_paper.setdefault(wk, pname)
    # Если в одном абзаце нет и недели и статьи (статья может быть отдельной строкой),
    # сделаем fallback: проверим, что присутствуют все три обозначения недель И
    # что им соответствуют 3 различные статьи по порядку появления.
    distinct_papers = set(week_to_paper.values())
    all_weeks_present = all(any(tok in full_text for tok in toks)
                            for toks in week_tokens.values())
    week_mapping_ok = len(week_to_paper) >= 3 and len(distinct_papers) >= 3
    if not week_mapping_ok and all_weeks_present and papers_mentioned >= 3:
        # Все недели и все 3 статьи присутствуют, но не в одном абзаце — принимаем.
        week_mapping_ok = True
    record("Word: каждая из 3 недель сопоставлена своей отдельной статье",
           week_mapping_ok,
           f"week->paper: {week_to_paper}, weeks_present={all_weeks_present}")

    # Критично: шумовая RLHF-статья НЕ должна попасть в документ.
    noise_hit = next((m for m in NOISE_MARKERS if m in full_text), None)
    record("Word НЕ содержит шумовую RLHF-статью (Ouyang/2203.02155)",
           noise_hit is None, f"Найден шумовой маркер: {noise_hit}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_gcal()
    check_word(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print(f"\n=== Results: {PASS_COUNT}/{total} passed ({accuracy:.1f}%) ===")
    if critical_failed:
        print(f"CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"  - {n}")

    if args.res_log_file:
        try:
            with open(args.res_log_file, "w") as f:
                json.dump({
                    "total_passed": PASS_COUNT, "total_checks": total,
                    "accuracy": accuracy, "critical_failed": critical_failed,
                }, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    success = (not critical_failed) and accuracy >= 70
    if success:
        print("All checks passed (no critical failures, accuracy >= 70%)!")
        sys.exit(0)
    else:
        print("Evaluation FAILED (critical failure or accuracy < 70%).")
        sys.exit(1)


if __name__ == "__main__":
    main()
