#!/usr/bin/env python3
"""Dump a flat checklist of all tasks for top-down review.

Output: tasks_review.md (checkbox list grouped by category, with markers).
Markers: ⚠️ = evaluator has unused check_* function; 中 = Chinese chars in eval.
Also annotates last-run result (P/EF/AF) when available.
"""
import csv
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tasks_review.csv"
OUT = ROOT / "tasks_review.md"


def short_result(r):
    if not r["last_result_status"]:
        return ""
    if r["last_result_pass"] == "True":
        return "P"
    if r["last_result_status"] == "success":
        return "EF"
    return "AF"


def main():
    rows = list(csv.DictReader(CSV.open()))
    by_cat = defaultdict(list)
    for r in rows:
        by_cat[r["category"]].append(r)

    lines = ["# Tasks review checklist", "",
             "Legend: ⚠️ = unused check_* function in evaluator; 中 = chinese in eval; "
             "result: P=pass, EF=eval_fail, AF=agent_fail.", ""]

    for cat in sorted(by_cat.keys()):
        tasks = sorted(by_cat[cat], key=lambda r: r["task"])
        lines.append(f"## {cat} ({len(tasks)})")
        lines.append("")
        for r in tasks:
            markers = []
            if r["eval_unused_funcs"]:
                markers.append(f"⚠️ {r['eval_unused_funcs']}")
            if r["eval_has_chinese"] == "True":
                markers.append("中")
            result = short_result(r)
            if result:
                markers.append(f"[{result}]")
            marker_str = " " + " ".join(markers) if markers else ""
            lines.append(f"- [ ] `{r['task']}`{marker_str}")
        lines.append("")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {OUT.relative_to(ROOT)}: {len(rows)} tasks across {len(by_cat)} categories")


if __name__ == "__main__":
    main()
