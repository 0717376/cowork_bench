"""Deterministic recorder for russification batches — NO LLM touches the CSV.

Takes a JSON report (list of per-task objects) produced by wf_russify.js and:
  1. renames task dirs whose name carried a pre-swap token (rename_to),
  2. rewrites the matching tasks_review.csv row (new name, current needed_mcps
     read from the post-edit task_config.json, verdict-derived status, notes).

Static-only run (no container): last_result_pass stays blank; status reflects
the static adversarial verdict (ru-static-ok / needs-fix / blocked-on-fork).

Usage: python3 scripts/ru_record.py <report.json>
Report entry shape: {task, rename_to|null, verdict in {ru_ok,needs_fix,blocked},
                     eval_bug:bool, note:str}
"""
import csv
import json
import os
import subprocess
import sys

PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FP = os.path.join(PROJECT, "tasks", "finalpool")
CSV = os.path.join(PROJECT, "tasks_review.csv")

STATUS = {"ru_ok": "ru-static-ok", "needs_fix": "needs-fix", "blocked": "blocked-on-fork"}
LAST = {"ru_ok": "static-ok", "needs_fix": "needs-fix", "blocked": "blocked-on-fork"}


def mcps(name):
    p = os.path.join(FP, name, "task_config.json")
    if not os.path.isfile(p):
        return ""
    try:
        return ";".join(json.load(open(p)).get("needed_mcp_servers", []) or [])
    except Exception:
        return ""


def main(report_path):
    report = json.load(open(report_path))
    rows = list(csv.DictReader(open(CSV)))
    hdr = list(rows[0].keys())
    idx = {r["task"]: r for r in rows}
    renamed = updated = 0
    missing = []
    for e in report:
        if not e:
            continue
        old = e["task"]
        new = e.get("rename_to") or old
        verdict = e.get("verdict", "needs_fix")
        if new != old:
            src, dst = os.path.join(FP, old), os.path.join(FP, new)
            if os.path.isdir(src) and not os.path.isdir(dst):
                subprocess.run(["mv", src, dst], check=True)
                renamed += 1
            elif not os.path.isdir(dst):
                new = old  # source missing and dst absent -> keep old name
        r = idx.get(old)
        if not r:
            missing.append(old)
            continue
        r["task"] = new
        r["needed_mcps"] = mcps(new)
        r["last_result_pass"] = ""
        r["last_result_status"] = LAST.get(verdict, verdict)
        r["eval_bug"] = "yes" if e.get("eval_bug") else "no"
        r["status"] = STATUS.get(verdict, "needs-fix")
        r["notes"] = (e.get("note") or "")[:500]
        if not r.get("task_quality"):
            r["task_quality"] = "good"
        if not r.get("priority"):
            r["priority"] = "P1"
        if not r.get("ru_effort_min"):
            r["ru_effort_min"] = "30"
        updated += 1
    with open(CSV, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=hdr)
        w.writeheader()
        w.writerows(rows)
    bad = [r["task"] for r in rows if len(r) != len(hdr)]
    print(json.dumps({"renamed": renamed, "updated": updated, "missing": missing,
                      "rows": len(rows), "shape_ok": not bad}, ensure_ascii=False))


if __name__ == "__main__":
    main(sys.argv[1])
