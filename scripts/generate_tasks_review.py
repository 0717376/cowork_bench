#!/usr/bin/env python3
"""Generate tasks_review.csv: one row per task in tasks/finalpool/.

Auto-fills columns derivable from the filesystem (category, needed MCPs,
evaluator structure heuristics, groundtruth files, latest run result).
Manual columns are left blank for human review.
"""
import csv
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
TASKS_DIR = ROOT / "tasks" / "finalpool"
BENCH_LOGS_DIR = ROOT / "benchmark_logs"
OUT = ROOT / "tasks_review.csv"

CHINESE_RE = re.compile(r"[一-鿿]")
DEF_RE = re.compile(r"^\s*def\s+(\w+)\s*\(", re.MULTILINE)

# Categories actually present in tasks/finalpool/ (by docs/README and inspection).
# Order matters — first match wins for compound prefixes like "12306-wc-...".
KNOWN_CATEGORIES = [
    "12306", "rzd", "wc", "sf", "yf", "canvas", "arxiv", "kulinar", "yt",
    "academic", "scholarly", "fetch", "playwright", "pw", "terminal",
    "notion", "ppt", "gsheet", "gform", "gcal", "memory", "research",
    "train", "ecommerce",
]
# Helpers — not real check_* functions; ignore as "unused" signal
HELPER_FN_NAMES = {
    "get_conn", "safe_float", "str_contains", "num_close", "nums_close",
    "number_close_match", "load_json", "load_workbook", "normalize",
    "to_float", "to_int", "approx", "close_enough",
}


def category_of(name: str) -> str:
    for cat in KNOWN_CATEGORIES:
        if name == cat or name.startswith(cat + "-"):
            return cat
    return "other"


def latest_run_results() -> dict:
    """Map task -> (status, eval_pass, duration_s) from most-recent summary.csv."""
    if not BENCH_LOGS_DIR.exists():
        return {}
    runs = sorted(BENCH_LOGS_DIR.glob("fully_parallel_*"), reverse=True)
    for run in runs:
        s = run / "summary.csv"
        if s.exists():
            out = {}
            for row in csv.DictReader(s.open()):
                out[row["task"]] = row
            return out
    return {}


def analyze_eval(eval_file: Path) -> dict:
    """Heuristics on evaluation/main.py: defined fns vs called, chinese chars."""
    if not eval_file.exists():
        return dict(eval_loc=0, eval_funcs_defined=0, eval_funcs_called=0,
                    eval_unused_funcs="", eval_has_chinese=False)
    text = eval_file.read_text(encoding="utf-8", errors="replace")
    funcs = DEF_RE.findall(text)
    # exclude trivial wrappers / main itself
    check_funcs = [f for f in funcs if f not in ("main", "check", "num_close",
                                                  "str_match", "get_sheet")]
    called = set()
    for fn in check_funcs:
        # Count calls inside `def main():` or top-level
        # crude: any occurrence of "{fn}(" outside its own def line
        for m in re.finditer(rf"\b{re.escape(fn)}\s*\(", text):
            # skip definition line
            line_start = text.rfind("\n", 0, m.start()) + 1
            line = text[line_start:m.end()]
            if not line.lstrip().startswith("def "):
                called.add(fn)
                break
    unused_all = sorted(set(check_funcs) - called)
    # Filter to real check_* functions (and similar) — helpers are noise.
    unused_real = [u for u in unused_all
                   if u not in HELPER_FN_NAMES and not u.startswith("_")]
    return dict(
        eval_loc=text.count("\n") + 1,
        eval_funcs_defined=len(check_funcs),
        eval_funcs_called=len(called),
        eval_unused_funcs=";".join(unused_real),
        eval_has_chinese=bool(CHINESE_RE.search(text)),
    )


def gt_files(task: Path) -> str:
    gt = task / "groundtruth_workspace"
    if not gt.exists():
        return ""
    return ";".join(sorted(p.name for p in gt.iterdir() if p.is_file()))


def needed_mcps(task: Path) -> str:
    cfg = task / "task_config.json"
    if not cfg.exists():
        return ""
    try:
        data = json.loads(cfg.read_text())
        return ";".join(data.get("needed_mcp_servers", []))
    except Exception:
        return ""


def task_word_count(task: Path) -> int:
    md = task / "docs" / "task.md"
    if not md.exists():
        return 0
    return len(md.read_text(encoding="utf-8", errors="replace").split())


def has_preprocess(task: Path) -> bool:
    p = task / "preprocess" / "main.py"
    if not p.exists():
        return False
    # skip empty stubs
    return p.read_text(encoding="utf-8", errors="replace").strip() != ""


MANUAL_COLS = ("eval_bug", "task_quality", "ru_effort_min",
               "priority", "status", "notes")


def load_existing_manual() -> dict:
    """Read prior CSV (if exists) and return {task: {manual_col: value}} for
    any non-empty manual annotations. Lets re-runs preserve hand-edits."""
    if not OUT.exists():
        return {}
    out = {}
    for row in csv.DictReader(OUT.open()):
        kept = {c: row.get(c, "") for c in MANUAL_COLS if row.get(c, "")}
        if kept:
            out[row["task"]] = kept
    return out


def main():
    if not TASKS_DIR.exists():
        sys.exit(f"Not found: {TASKS_DIR}")

    last = latest_run_results()
    prior = load_existing_manual()

    rows = []
    for task in sorted(TASKS_DIR.iterdir()):
        if not task.is_dir():
            continue
        name = task.name
        if name.startswith(".") or name in ("__pycache__",):
            continue
        eval_info = analyze_eval(task / "evaluation" / "main.py")
        run = last.get(name, {})
        rows.append(dict(
            task=name,
            category=category_of(name),
            needed_mcps=needed_mcps(task),
            task_md_words=task_word_count(task),
            eval_loc=eval_info["eval_loc"],
            eval_funcs_defined=eval_info["eval_funcs_defined"],
            eval_funcs_called=eval_info["eval_funcs_called"],
            eval_unused_funcs=eval_info["eval_unused_funcs"],
            eval_has_chinese=eval_info["eval_has_chinese"],
            gt_files=gt_files(task),
            has_preprocess=has_preprocess(task),
            last_result_status=run.get("status", ""),
            last_result_pass=run.get("eval_pass", ""),
            last_result_duration_s=run.get("duration_s", ""),
            eval_bug=prior.get(name, {}).get("eval_bug", ""),
            task_quality=prior.get(name, {}).get("task_quality", ""),
            ru_effort_min=prior.get(name, {}).get("ru_effort_min", ""),
            priority=prior.get(name, {}).get("priority", ""),
            status=prior.get(name, {}).get("status", "pending"),
            notes=prior.get(name, {}).get("notes", ""),
        ))

    fields = list(rows[0].keys())
    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

    susp = sum(1 for r in rows if r["eval_funcs_defined"] > r["eval_funcs_called"])
    cn = sum(1 for r in rows if r["eval_has_chinese"])
    print(f"Wrote {OUT.relative_to(ROOT)}: {len(rows)} rows")
    print(f"  Categories: {len(set(r['category'] for r in rows))}")
    print(f"  Suspicious evaluators (defined > called): {susp}")
    print(f"  Evaluators with Chinese chars (literal-string risk): {cn}")


if __name__ == "__main__":
    main()
