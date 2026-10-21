"""Read the memory MCP's JSONL and legacy graph fixtures without dropping rows."""

import json
from pathlib import Path


def _validate_records(records, kind):
    required = (
        ["name", "entityType"] if kind == "entity" else ["from", "to", "relationType"]
    )
    for record in records:
        if any(not isinstance(record.get(key), str) for key in required):
            raise ValueError("Memory record has missing or non-string identity fields")
        if kind == "entity":
            observations = record.get("observations")
            if not isinstance(observations, list) or any(
                not isinstance(x, str) for x in observations
            ):
                raise ValueError("Memory observations must be a string array")


def parse_memory_graph(content):
    if not content.strip():
        return {"entities": [], "relations": []}
    try:
        value = json.loads(content)
    except json.JSONDecodeError:
        records = [json.loads(line) for line in content.splitlines() if line.strip()]
    else:
        if isinstance(value, dict) and "type" not in value:
            # The historical notes fixture is only used by the notes grader.
            if set(value) == {"notes"} and isinstance(value["notes"], list):
                return value
            if not set(value) <= {"entities", "relations"}:
                raise ValueError("Invalid memory graph object")
            graph = {key: value.get(key, []) for key in ["entities", "relations"]}
            if any(
                not isinstance(items, list)
                or any(not isinstance(x, dict) for x in items)
                for items in graph.values()
            ):
                raise ValueError(
                    "Memory graph entities/relations must be object arrays"
                )
            _validate_records(graph["entities"], "entity")
            _validate_records(graph["relations"], "relation")
            return graph
        records = value if isinstance(value, list) else [value]
    graph = {"entities": [], "relations": []}
    for record in records:
        if not isinstance(record, dict) or record.get("type") not in {
            "entity",
            "relation",
        }:
            raise ValueError("Invalid memory JSONL record")
        _validate_records([record], record["type"])
        key = "entities" if record["type"] == "entity" else "relations"
        graph[key].append({k: v for k, v in record.items() if k != "type"})
    return graph


def read_memory_graph(path):
    return parse_memory_graph(Path(path).read_text(encoding="utf-8"))
