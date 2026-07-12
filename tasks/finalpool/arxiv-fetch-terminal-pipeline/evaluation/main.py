"""
Evaluation script for arxiv-fetch-terminal-pipeline task.

Checks:
1. Memory knowledge graph has entities for papers and a research session
2. paper_data.json was created with correct paper data
3. analysis_results.json exists with correct statistics

Critical checks (see CRITICAL_CHECKS): any failure there => overall FAIL
regardless of accuracy. Pass threshold otherwise: accuracy >= 80%.

Usage:
    python -m evaluation.main --agent_workspace <path> --groundtruth_workspace <path>
"""
import argparse
import json
import os
import sys

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Critical checks: any failure => overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "paper_data.json exists",
    "analysis_results.json exists",
    "Total papers is 4",
    "At least 3 of 4 expected paper IDs in results",
}


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def float_close(a, b, tol=50.0):
    """Compare two numeric values with tolerance."""
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _load_json(path):
    try:
        with open(path, "r") as f:
            return json.load(f), None
    except Exception as e:
        return None, str(e)


EXPECTED_PAPER_IDS = {"2107.03374", "2002.08155", "2203.07814", "2305.06161"}
EXPECTED_TOTAL_CITATIONS = 9000


def check_memory(agent_workspace):
    """Check that memory.json has entities for papers and a research session."""
    print("\n=== Checking Memory ===")

    memory_path = os.path.join(agent_workspace, "memory", "memory.json")
    if not os.path.isfile(memory_path):
        check("memory.json exists", False, f"Not found: {memory_path}")
        return

    check("memory.json exists", True)

    with open(memory_path, "r") as f:
        content = f.read().strip()

    if not content or content == "{}":
        check("Memory has content", False, "memory.json is empty")
        return

    check("Memory has content", True)

    # Memory format from @modelcontextprotocol/server-memory is a knowledge graph.
    # The memory file may be JSONL (one object per line) or a single JSON object.
    memory_data = None
    try:
        memory_data = json.loads(content)
    except json.JSONDecodeError:
        # try JSONL
        ents = []
        for line in content.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(obj, dict):
                ents.append(obj)
        memory_data = {"entities": [e for e in ents if e.get("type") != "relation"]}

    check("Memory is valid JSON", memory_data is not None, "Cannot parse memory.json")
    if memory_data is None:
        return

    # Collect entities (support several shapes).
    entities = []
    if isinstance(memory_data, list):
        entities = memory_data
    elif isinstance(memory_data, dict):
        entities = memory_data.get("entities", [])

    entity_text = ""
    for ent in entities:
        if isinstance(ent, dict):
            entity_text += json.dumps(ent, ensure_ascii=False).lower() + " "

    # Check for paper entities (at least 3 of the markers below).
    paper_keywords = ["codex", "codebert", "alphacode", "starcoder",
                      "2107.03374", "2002.08155", "2203.07814", "2305.06161",
                      "code generation", "evaluating large"]
    paper_entity_count = sum(1 for kw in paper_keywords if kw in entity_text)
    check("Memory has paper entities (at least 3 keywords found)",
          paper_entity_count >= 3,
          f"Found {paper_entity_count} paper-related keywords in memory entities")

    # Check for a research session entity (RU/EN tolerant).
    has_session = any(m in entity_text for m in
                      ["research_session", "research session", "session",
                       "сесси", "code generation", "исследован"])
    check("Memory has research session entity",
          has_session,
          "No research_session entity found")


def check_paper_data(agent_workspace):
    """Check paper_data.json was created with required fields."""
    print("\n=== Checking paper_data.json ===")

    paper_data_path = os.path.join(agent_workspace, "paper_data.json")
    if not os.path.isfile(paper_data_path):
        check("paper_data.json exists", False, f"Not found: {paper_data_path}")
        return

    check("paper_data.json exists", True)

    papers, err = _load_json(paper_data_path)
    if err is not None:
        check("paper_data.json is valid JSON", False, err)
        return
    check("paper_data.json is valid JSON", True)

    is_list = isinstance(papers, list)
    check("paper_data.json contains a list", is_list, f"Type: {type(papers)}")
    if not is_list:
        return

    check("paper_data.json has 4 papers", len(papers) == 4, f"Found {len(papers)} papers")

    # Each paper must have the required fields.
    required_fields = ["title", "arxiv_id", "authors", "citation_count", "abstract"]
    all_ok = True
    for i, paper in enumerate(papers):
        if not isinstance(paper, dict):
            all_ok = False
            continue
        missing = [f for f in required_fields if f not in paper]
        if missing:
            all_ok = False
            check(f"Paper {i+1} missing fields", False, f"Missing: {missing}")
    check("All papers have required fields", all_ok and len(papers) > 0,
          "Some papers are missing required fields")


def check_analysis_results(agent_workspace, groundtruth_workspace):
    """Check analysis_results.json against groundtruth expectations."""
    print("\n=== Checking Analysis Results ===")

    agent_file = os.path.join(agent_workspace, "analysis_results.json")

    if not os.path.isfile(agent_file):
        check("analysis_results.json exists", False, f"Not found: {agent_file}")
        return

    check("analysis_results.json exists", True)

    agent_results, err = _load_json(agent_file)
    if err is not None:
        check("analysis_results.json is valid JSON", False, err)
        return
    check("analysis_results.json is valid JSON", True)

    # Total papers
    check("Total papers is 4",
          agent_results.get("total_papers") == 4,
          f"Got {agent_results.get('total_papers')}")

    # Total citations (with tolerance)
    total_cit = agent_results.get("total_citations", 0)
    check("Total citations close to expected",
          float_close(total_cit, EXPECTED_TOTAL_CITATIONS, tol=500),
          f"Got {total_cit}, expected ~{EXPECTED_TOTAL_CITATIONS}")

    # Most cited paper
    most_cited = (agent_results.get("most_cited_paper", "") or "").lower()
    check("Most cited paper is Codex/Evaluating LLMs",
          "evaluating" in most_cited or "codex" in most_cited or "code" in most_cited,
          f"Got '{most_cited}'")

    # Paper IDs present
    paper_ids = set(agent_results.get("paper_arxiv_ids", []))
    overlap = paper_ids & EXPECTED_PAPER_IDS
    check("At least 3 of 4 expected paper IDs in results",
          len(overlap) >= 3,
          f"Found {len(overlap)} matching IDs: {overlap}")

    # Paper titles present
    titles = agent_results.get("paper_titles", [])
    check("At least 3 paper titles in results",
          len(titles) >= 3,
          f"Found {len(titles)} titles")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    check_memory(args.agent_workspace)
    check_paper_data(args.agent_workspace)
    check_analysis_results(args.agent_workspace, gt_dir)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    print(f"\n=== Results: {PASS_COUNT}/{total} passed ({accuracy:.1f}%) ===")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print(f"CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"  - {n}")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "critical_failed": critical_failed,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if critical_failed:
        print("FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 80:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
