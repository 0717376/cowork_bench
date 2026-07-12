"""
Evaluation for the weekly office lunch menu task (kulinar, russified).

Structural (NON-critical) checks:
  - Weekly_Menu.docx exists and is readable
  - Has at least 5 weekday sections (Mon-Fri, RU+EN)
  - Document has sufficient content (>= 500 chars)
  - memory.json updated with entities

Semantic CRITICAL checks (any failure => overall FAIL, before accuracy gate):
  - The five weekday dishes parsed from the docx all exist in the kulinar
    source data AND each has difficulty <= 3.
  - The five dishes span at least 4 distinct kulinar categories AND include
    at least one soup (суп) and at least one vegetable (овощное) dish.
  - No category is repeated on consecutive days (Mon->Fri).
  - filtered_recipes.json exists, every entry has difficulty <= 3, and all
    five selected dishes are present in it (proves the filter pipeline ran).
  - memory.json has a dietary_preferences entity capturing the rules AND a
    selections entity whose dish names match the five docx dishes.

Pass: no critical failure AND accuracy >= 70%.
"""
import json
import os
import re
import sys
from argparse import ArgumentParser
from datetime import datetime

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Critical (semantic) checks: any failure => overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "All five dishes exist in kulinar and difficulty <= 3",
    ">= 4 distinct categories incl. at least one soup and one vegetable",
    "No category repeated on consecutive days",
    "filtered_recipes.json valid (all diff<=3, contains the five dishes)",
    "memory.json: dietary_preferences + matching selections",
}


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        print(f"  [FAIL] {name}: {detail}")


def _norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower())


def load_kulinar_recipes():
    """Возвращает список рецептов kulinar или None, если база недоступна.

    Ищем all_recipes.json: сначала переменная окружения, затем поднимаемся
    вверх по дереву от файла eval до local_servers/kulinar-mcp/...
    """
    candidates = []
    env = os.environ.get("KULINAR_RECIPES_JSON")
    if env:
        candidates.append(env)

    here = os.path.abspath(__file__)
    cur = os.path.dirname(here)
    rel = os.path.join("local_servers", "kulinar-mcp", "src", "data", "all_recipes.json")
    for _ in range(12):
        candidates.append(os.path.join(cur, rel))
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent

    for path in candidates:
        try:
            if path and os.path.exists(path):
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list) and data:
                    return data
        except Exception:
            continue
    return None


KULINAR = load_kulinar_recipes()
KULINAR_BY_NAME = {_norm_name(r["name"]): r for r in KULINAR} if KULINAR else {}

# Day labels (RU primary; EN accepted for back-compat).
DAY_LABELS = [
    ("monday", ["понедельник", "monday"]),
    ("tuesday", ["вторник", "tuesday"]),
    ("wednesday", ["среда", "wednesday"]),
    ("thursday", ["четверг", "thursday"]),
    ("friday", ["пятница", "friday"]),
]


def parse_docx_dishes(paragraphs):
    """Парсит названия выбранных блюд из секций по дням.

    Стратегия: разбиваем параграфы на блоки по заголовкам дней; внутри блока
    ищем строку вида 'Блюдо: <name>' / 'Dish: <name>', иначе сопоставляем
    любой kulinar-рецепт, чьё имя встречается в тексте блока.
    Возвращает список (day_key, dish_name_or_None) по порядку дней.
    """
    texts = [p for p in paragraphs]
    # locate index of each day's heading
    day_idx = []
    for i, t in enumerate(texts):
        low = t.strip().lower()
        for key, labels in DAY_LABELS:
            if any(low == lab or low.startswith(lab) for lab in labels):
                day_idx.append((i, key))
                break

    result = {}
    for j, (i, key) in enumerate(day_idx):
        end = day_idx[j + 1][0] if j + 1 < len(day_idx) else len(texts)
        block = texts[i:end]
        block_text = "\n".join(block)
        dish = None
        # explicit "Блюдо:" / "Dish:" line
        for line in block:
            m = re.match(r"\s*(?:блюдо|dish)\s*[:\-]\s*(.+)", line.strip(), re.IGNORECASE)
            if m:
                dish = m.group(1).strip()
                break
        # fallback: match any kulinar recipe name appearing in the block
        if dish is None and KULINAR_BY_NAME:
            low_block = block_text.lower()
            best = None
            for nm in KULINAR_BY_NAME:
                if nm in low_block:
                    if best is None or len(nm) > len(best):
                        best = nm
            if best is not None:
                dish = KULINAR_BY_NAME[best]["name"]
        if dish is not None and key not in result:
            result[key] = dish
    # ordered by weekday sequence
    return [(key, result.get(key)) for key, _ in DAY_LABELS]


def is_vegetable(recipe):
    tags = [str(t).lower() for t in recipe.get("tags", [])]
    if "овощное" in tags:
        return True
    # category-based fallback: salads/sides that are plant-based
    if recipe.get("category") in ("салат", "гарнир") and (
        "постное" in tags or "летнее" in tags
    ):
        return True
    return False


def main():
    global PASS_COUNT, FAIL_COUNT

    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    agent_ws = args.agent_workspace

    # ---- Structural: Weekly_Menu.docx exists and is readable ----
    docx_path = os.path.join(agent_ws, "Weekly_Menu.docx")
    if not os.path.exists(docx_path):
        check("Weekly_Menu.docx exists", False, f"Not found at {docx_path}")
        print(f"\nResults: {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} passed, {FAIL_COUNT} failed")
        sys.exit(1)

    try:
        from docx import Document
        doc = Document(docx_path)
        all_paragraphs = [p.text for p in doc.paragraphs]
        full_text = "\n".join(all_paragraphs)
        check("Weekly_Menu.docx exists and is readable", True)
    except Exception as e:
        check("Weekly_Menu.docx exists and is readable", False, str(e))
        print(f"\nResults: {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} passed, {FAIL_COUNT} failed")
        sys.exit(1)

    low_text = full_text.lower()

    # ---- Structural: at least 5 weekday sections (RU+EN) ----
    found_days = []
    for key, labels in DAY_LABELS:
        if any(lab in low_text for lab in labels):
            found_days.append(key)
    check("Has at least 5 weekday sections (Mon-Fri)",
          len(found_days) >= 5,
          f"Found {len(found_days)} weekday(s): {found_days}")

    # ---- Structural: sufficient content ----
    check("Document has sufficient content (>= 500 chars)",
          len(full_text.strip()) >= 500,
          f"Only {len(full_text.strip())} characters")

    # ---- Parse selected dishes from the docx ----
    if not KULINAR:
        check("Kulinar source data available", False,
              "all_recipes.json not found; cannot run semantic checks")
    else:
        check("Kulinar source data available", True)

    parsed = parse_docx_dishes(all_paragraphs)
    selected = [(d, nm) for d, nm in parsed if nm]
    resolved = []  # list of (day, recipe-dict) for dishes found in kulinar
    unknown = []
    for day, nm in selected:
        rec = KULINAR_BY_NAME.get(_norm_name(nm)) if KULINAR_BY_NAME else None
        if rec:
            resolved.append((day, rec))
        else:
            unknown.append((day, nm))

    # ---- CRITICAL 1: all five dishes exist in kulinar and difficulty <= 3 ----
    have_five = len(selected) >= 5
    all_known = have_five and not unknown
    all_easy = all_known and all(r.get("difficulty", 99) <= 3 for _, r in resolved)
    check("All five dishes exist in kulinar and difficulty <= 3",
          all_known and all_easy,
          f"selected={[(d, n) for d, n in selected]}, unknown={unknown}, "
          f"difficulties={[(d, r.get('difficulty')) for d, r in resolved]}")

    # ---- CRITICAL 2: >=4 categories incl. soup + vegetable ----
    cats = [r.get("category") for _, r in resolved]
    distinct_cats = set(cats)
    has_soup = any(r.get("category") == "суп" for _, r in resolved)
    has_veg = any(is_vegetable(r) for _, r in resolved)
    check(">= 4 distinct categories incl. at least one soup and one vegetable",
          all_known and len(distinct_cats) >= 4 and has_soup and has_veg,
          f"categories={cats}, distinct={len(distinct_cats)}, "
          f"soup={has_soup}, vegetable={has_veg}")

    # ---- CRITICAL 3: no category repeated on consecutive days ----
    # Build sequence in weekday order using the resolved recipes.
    by_day = {day: r for day, r in resolved}
    seq = [by_day[k].get("category") for k, _ in DAY_LABELS if k in by_day]
    no_consec = all_known and all(
        seq[i] != seq[i + 1] for i in range(len(seq) - 1)
    )
    check("No category repeated on consecutive days",
          no_consec,
          f"category sequence={seq}")

    # ---- CRITICAL 4: filtered_recipes.json valid ----
    filt_path = os.path.join(agent_ws, "filtered_recipes.json")
    filt_ok = False
    filt_detail = ""
    if not os.path.exists(filt_path):
        filt_detail = f"not found at {filt_path}"
    else:
        try:
            with open(filt_path, encoding="utf-8") as f:
                filt = json.load(f)
            if not isinstance(filt, list) or not filt:
                filt_detail = "filtered_recipes.json is empty or not a list"
            else:
                all_diff_ok = all(
                    isinstance(e, dict) and e.get("difficulty", 99) <= 3 for e in filt
                )
                filt_names = {_norm_name(e.get("name", "")) for e in filt if isinstance(e, dict)}
                sel_names = {_norm_name(n) for _, n in selected}
                contains_all = bool(sel_names) and sel_names.issubset(filt_names)
                filt_ok = all_diff_ok and contains_all
                filt_detail = (f"all_diff<=3={all_diff_ok}, "
                               f"missing={sorted(sel_names - filt_names)}")
        except Exception as e:
            filt_detail = f"error reading filtered_recipes.json: {e}"
    check("filtered_recipes.json valid (all diff<=3, contains the five dishes)",
          all_known and filt_ok, filt_detail)

    # ---- CRITICAL 5: memory.json dietary_preferences + matching selections ----
    memory_path = os.path.join(agent_ws, "memory", "memory.json")
    mem_struct_ok = False
    mem_detail = ""
    entities = []
    if not os.path.exists(memory_path):
        mem_detail = f"memory.json not found at {memory_path}"
    else:
        try:
            with open(memory_path, encoding="utf-8") as f:
                raw = f.read()
            try:
                mem_data = json.loads(raw)
                entities = mem_data.get("entities", [])
            except json.JSONDecodeError:
                # The issued memory MCP persists its graph as JSONL (one JSON
                # object per line, type=="entity"/"relation"), not a single
                # pretty-printed object. Reconstruct entities from those lines.
                entities = []
                for line in raw.splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    obj = json.loads(line)
                    if isinstance(obj, dict) and obj.get("type") == "entity":
                        entities.append(obj)

            def ent_text(e):
                obs = e.get("observations", []) or []
                return (" ".join([str(e.get("name", ""))] + [str(o) for o in obs])).lower()

            blob = " ".join(ent_text(e) for e in entities)
            # dietary_preferences captured: look for the four rules' keywords
            has_pref = (
                "слож" in blob or "difficulty" in blob
            ) and (
                "суп" in blob or "soup" in blob
            ) and (
                "категор" in blob or "categor" in blob
            )
            # selections recorded: all five docx dish names appear in memory
            sel_names = {_norm_name(n) for _, n in selected}
            mem_low = blob
            sel_in_mem = bool(sel_names) and all(n in mem_low for n in sel_names)
            mem_struct_ok = has_pref and sel_in_mem
            mem_detail = f"has_pref={has_pref}, selections_in_memory={sel_in_mem}"
        except Exception as e:
            mem_detail = f"error reading memory.json: {e}"
    check("memory.json: dietary_preferences + matching selections",
          all_known and mem_struct_ok, mem_detail)

    # ---- Structural: memory non-empty ----
    check("Memory file has been updated with entities",
          len(entities) > 0,
          "No entities found in memory.json")

    # ---- Summary + gating ----
    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total > 0 else 0
    print(f"\nResults: {PASS_COUNT}/{total} passed ({accuracy:.1f}%), {FAIL_COUNT} failed")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print(f"CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"  - {n}")

    if args.res_log_file:
        result = {
            "total_passed": PASS_COUNT,
            "total_checks": total,
            "accuracy": accuracy,
            "critical_failed": critical_failed,
            "timestamp": datetime.now().isoformat(),
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if critical_failed:
        print("FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    print("FAIL")
    sys.exit(1)


if __name__ == "__main__":
    main()
