#!/usr/bin/env python3
"""Скрипт предобработки для подготовки задачи budget-variance-analysis-quarterly.

Источник данных ClickHouse в этой задаче является фантомной зависимостью: все
исходные данные лежат в локальных файлах (approved_budget.xlsx, cost_center_mapping.csv),
поэтому инъекция в БД не требуется. Предобработка лишь идемпотентно удаляет
выходные файлы предыдущих запусков агента, НЕ пред-создавая никаких результатов.
"""

from argparse import ArgumentParser
from pathlib import Path

# Выходные артефакты, которые должен создать сам агент. Удаляем их перед запуском,
# чтобы прогон был чистым (без пред-сидинга ответов).
AGENT_OUTPUTS = [
    "variance_analysis.xlsx",
    "variance_tracking.xlsx",
    "budget_forecast.xlsx",
    "dept_variance_reports.docx",
    "executive_presentation.pptx",
]


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        # Идемпотентная очистка прежних результатов агента.
        for name in AGENT_OUTPUTS:
            f = agent_ws / name
            if f.exists():
                try:
                    f.unlink()
                except OSError:
                    pass

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
