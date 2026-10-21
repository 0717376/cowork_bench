"""Policy acceptance does not mean every component check passed."""
import asyncio
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import AsyncMock, patch
from utils.evaluation import evaluator


class EvaluationPolicy(unittest.TestCase):
    def test_partial_credit_acceptance_is_described_as_policy_acceptance(self):
        with tempfile.TemporaryDirectory() as temp:
            cfg = SimpleNamespace(task_dir="synthetic", log_file="unused",
                agent_workspace=temp, launch_time="2026-01-01 00:00:00",
                evaluation=SimpleNamespace(groundtruth_workspace=temp,
                                           evaluation_command="synthetic-grader"))
            with patch.object(evaluator.TaskConfig, "from_dict", return_value=cfg), \
                 patch.object(evaluator, "check_http_fixtures", create=True, return_value=[]), \
                 patch.object(evaluator, "run_command", new_callable=AsyncMock,
                              return_value=("[FAIL] optional formatting\nOverall 80%; PASS\n", "", 0)):
                result = asyncio.run(evaluator.TaskEvaluator.evaluate_one(
                    {"config": {"task_root": temp}, "status": "success"}))
        self.assertIs(result["pass"], True)
        self.assertNotIn("All evaluation checks passed", result["details"])
        self.assertIn("policy", result["details"])


if __name__ == "__main__":
    unittest.main()
