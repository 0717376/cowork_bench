#!/usr/bin/env python3
"""Скрипт предобработки: подготовка рабочего каталога агента.

Копирует ТОЛЬКО исходные данные, которые агент потребляет
(support_data_export.csv, survey_template.xlsx, config.json, README.md).
НЕ копирует итоговые файлы-ответы (satisfaction_analysis.xlsx, action_plans.xlsx,
satisfaction_report.docx, executive_summary.docx) — их создаёт агент.
Идемпотентно: очищает прежние итоговые файлы, если они остались от прошлого запуска.
"""

from argparse import ArgumentParser
import shutil
from pathlib import Path

# Исходные файлы, которые получает агент (НЕ ответы).
SOURCE_FILES = [
    "support_data_export.csv",
    "survey_template.xlsx",
    "config.json",
    "README.md",
]

# Итоговые файлы-ответы, которые должен произвести агент сам.
ANSWER_FILES = [
    "satisfaction_analysis.xlsx",
    "action_plans.xlsx",
    "satisfaction_report.docx",
    "executive_summary.docx",
]


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if not args.agent_workspace:
        print("Preprocess: no agent_workspace provided, nothing to do")
        return

    agent_ws = Path(args.agent_workspace)
    agent_ws.mkdir(parents=True, exist_ok=True)

    initial_ws = Path(__file__).resolve().parent.parent / "initial_workspace"

    # Идемпотентная очистка прежних ответов.
    for fname in ANSWER_FILES:
        target = agent_ws / fname
        if target.exists():
            target.unlink()

    # Копируем исходные данные.
    for fname in SOURCE_FILES:
        src = initial_ws / fname
        if src.exists():
            shutil.copy2(src, agent_ws / fname)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
