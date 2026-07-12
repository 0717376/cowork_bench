#!/usr/bin/env python3
"""Скаффолд для русификации задачи Cowork-Bench.

Копирует структуру tasks/finalpool/<FROM>/ в tasks/finalpool/<TO>/, заменяет
agent_system_prompt.md на русский шаблон, чистит task_config.json от заведомо
неработающих/чужих MCP (rail_12306 → rzd, notion удаляется), добавляет
заглушку строки в tasks_review.csv и печатает TODO-чеклист.

Содержимое task.md, evaluation/main.py, preprocess/main.py — НЕ переводит.
Это делает пользователь руками, исходя из выбранного маршрута/города/дат.

Usage:
    scripts/scaffold_ru_task.py --from 12306-foo-bar --to rzd-foo-bar
    scripts/scaffold_ru_task.py --from 12306-foo-bar --to rzd-foo-bar --no-csv
"""
from __future__ import annotations

import argparse
import csv
import json
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
TASKS_DIR = REPO_ROOT / "tasks" / "finalpool"
CSV_PATH = REPO_ROOT / "tasks_review.csv"

RU_SYSTEM_PROMPT = """Доступная рабочая директория: !!<<<<||||workspace_dir||||>>>>!!
При обработке задач, если нужно читать/записывать локальные файлы и пользователь указывает относительный путь, можно соединить его с указанной выше рабочей директорией, чтобы получить полный путь.
Если вы считаете, что задача выполнена, просто ответьте, не вызывая никаких инструментов — это сразу завершит задачу, и продолжить работу будет нельзя.
"""

MCP_REPLACEMENTS = {"rail_12306": "rzd"}
MCP_DROP = {"notion"}  # placeholder-токен — см. docs/ENVIRONMENT_QUIRKS.md


def fail(msg: str) -> None:
    print(f"[scaffold] FAIL — {msg}", file=sys.stderr)
    sys.exit(1)


def adjust_mcps(mcps: list[str]) -> tuple[list[str], list[str]]:
    """Применяет правила MCP_REPLACEMENTS / MCP_DROP. Возвращает (новый_список, заметки)."""
    out, notes = [], []
    for m in mcps:
        if m in MCP_DROP:
            notes.append(f"removed `{m}` (см. ENVIRONMENT_QUIRKS.md)")
            continue
        if m in MCP_REPLACEMENTS:
            new = MCP_REPLACEMENTS[m]
            out.append(new)
            notes.append(f"`{m}` → `{new}`")
        else:
            out.append(m)
    return out, notes


def copy_task(src: Path, dst: Path) -> None:
    shutil.copytree(src, dst)
    # Удалим pycache, который мог приехать из исходника.
    for cache in dst.rglob("__pycache__"):
        shutil.rmtree(cache, ignore_errors=True)


def patch_system_prompt(task_dir: Path) -> None:
    p = task_dir / "docs" / "agent_system_prompt.md"
    if p.exists():
        p.write_text(RU_SYSTEM_PROMPT, encoding="utf-8")


def patch_task_config(task_dir: Path) -> list[str]:
    cfg_path = task_dir / "task_config.json"
    if not cfg_path.exists():
        return ["task_config.json отсутствует — создайте вручную"]
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    orig_mcps = cfg.get("needed_mcp_servers", [])
    new_mcps, notes = adjust_mcps(orig_mcps)
    cfg["needed_mcp_servers"] = new_mcps
    cfg_path.write_text(json.dumps(cfg, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return notes


def append_csv_row(name: str, mcps: list[str]) -> None:
    if not CSV_PATH.exists():
        print(f"[scaffold] CSV не найден: {CSV_PATH}, пропущено")
        return
    # Шапка известна — 20 колонок (см. docs/ENVIRONMENT_QUIRKS.md).
    new_row = [
        name, "rzd", ";".join(mcps),
        "", "", "", "", "", "False", "", "True",     # task_md_words..has_preprocess
        "", "", "", "no", "", "", "P2",              # last_result_*..priority
        "scaffold", "Scaffolded from original; русифицировать task.md/evaluation/preprocess",
    ]
    with CSV_PATH.open("r", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    header_cols = len(rows[0])
    if len(new_row) != header_cols:
        fail(f"new row has {len(new_row)} cols, header has {header_cols}")
    # Найдём последнюю rzd-строку и вставим сразу после неё.
    insert_at = len(rows)
    for i, row in enumerate(rows):
        if row and row[1] == "rzd":
            insert_at = i + 1
    rows.insert(insert_at, new_row)
    with CSV_PATH.open("w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows(rows)
    print(f"[scaffold] CSV: вставлена строка #{insert_at + 1} в группу rzd")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--from", dest="src", required=True, help="имя исходной задачи (в tasks/finalpool/)")
    parser.add_argument("--to", dest="dst", required=True, help="имя новой задачи (в tasks/finalpool/)")
    parser.add_argument("--no-csv", action="store_true", help="не трогать tasks_review.csv")
    args = parser.parse_args()

    src = TASKS_DIR / args.src
    dst = TASKS_DIR / args.dst
    if not src.is_dir():
        fail(f"источник не найден: {src}")
    if dst.exists():
        fail(f"цель уже существует: {dst}")

    copy_task(src, dst)
    print(f"[scaffold] скопировано: {src.name} → {dst.name}")

    patch_system_prompt(dst)
    print("[scaffold] agent_system_prompt.md заменён на русский шаблон")

    mcp_notes = patch_task_config(dst)
    cfg = json.loads((dst / "task_config.json").read_text(encoding="utf-8"))
    mcps = cfg.get("needed_mcp_servers", [])
    print(f"[scaffold] task_config.json mcps: {mcps}")
    for n in mcp_notes:
        print(f"    {n}")

    if not args.no_csv:
        append_csv_row(dst.name, mcps)

    print()
    print("[scaffold] ✓ скелет готов. TODO руками:")
    print(f"    1. {dst}/docs/task.md — перевести на русский, поменять города/маршруты/даты")
    print(f"    2. {dst}/initial_workspace/* — перевести конфиги/гайды")
    print(f"    3. {dst}/preprocess/main.py — адаптировать инжект (русские summary/from_addr)")
    print(f"    4. {dst}/evaluation/main.py — обновить проверки (см. ENVIRONMENT_QUIRKS.md о normalize)")
    print(f"    5. если правили db/ — запустить scripts/test_db_migration.sh")
    print(f"    6. 1+5 прогонов стабильности через scripts/run_containerized.sh")
    print(f"    7. дополнить notes/metrics в tasks_review.csv (строка уже на месте)")


if __name__ == "__main__":
    main()
