"""Canonical memory JSONL must be readable by every memory grader."""

import contextlib
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from utils.memory_graph import parse_memory_graph

ROOT = Path(__file__).resolve().parents[2]


def jsonl(graph):
    return "\n".join(
        json.dumps(dict(type=kind, **r), ensure_ascii=False)
        for key, kind in [("entities", "entity"), ("relations", "relation")]
        for r in graph.get(key, [])
    )


class MemoryGraphContract(unittest.TestCase):
    def test_legacy_graphs_and_canonical_output_are_equivalent(self):
        for p in (ROOT / "tasks/finalpool").rglob("memory.json"):
            d = json.loads(p.read_text())
            if isinstance(d, dict) and set(d) <= {"entities", "relations"}:
                with self.subTest(path=str(p)):
                    graph = parse_memory_graph(p.read_text())
                    self.assertEqual(parse_memory_graph(jsonl(graph)), graph)

    def test_bad_rows_cannot_be_silently_ignored(self):
        valid = '{"type":"entity","name":"x","entityType":"note","observations":[]}'
        for data in [
            valid + "\nnot json",
            valid + '\n{"type":"bad"}',
            '{"entities":1}',
            "[]\n{}",
        ]:
            with self.subTest(data=data), self.assertRaises(ValueError):
                parse_memory_graph(data)
        self.assertEqual(parse_memory_graph(valid)["entities"][0]["name"], "x")

    def test_actual_grader_memory_checks_accept_equivalent_jsonl(self):
        for task in [
            "arxiv-fetch-terminal-pipeline",
            "memory-multi-source-research",
            "kulinar-teamly-recipe-kb",
            "arxiv-lit-review-gsheet",
        ]:
            taskroot = ROOT / "tasks/finalpool" / task
            graph = json.loads(
                (taskroot / "groundtruth_workspace/memory/memory.json").read_text()
            )
            spec = importlib.util.spec_from_file_location(
                task.replace("-", "_"), taskroot / "evaluation/main.py"
            )
            m = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(m)
            with tempfile.TemporaryDirectory() as temp:
                p = Path(temp) / "memory/memory.json"
                p.parent.mkdir()
                outcomes = []
                for content in [json.dumps(graph), jsonl(graph)]:
                    checks = []

                    def record(name, passed, *extra):
                        checks.append((name, bool(passed)))

                    if hasattr(m, "check"):
                        m.check = record
                    if hasattr(m, "record"):
                        m.record = record
                    p.write_text(content)
                    with contextlib.redirect_stdout(io.StringIO()):
                        m.check_memory(temp)
                    outcomes.append(checks)
                self.assertTrue(outcomes[0], task)
                self.assertEqual(outcomes[0], outcomes[1], task)
                self.assertFalse(
                    any(
                        not passed
                        for name, passed in outcomes[1]
                        if "JSON" in name or "readable" in name
                    ),
                    task,
                )


if __name__ == "__main__":
    unittest.main()
