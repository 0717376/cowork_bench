"""Cowork-Bench entry point for the Strands runner.

Mirrors main.py's CLI/env contract so it can be slotted into run_containerized.sh
by swapping the script name. Model config (LLM_BASE_URL / LLM_API_KEY / LLM_MODEL)
is read by strands_runner.model.make_model — see strands_runner/model.py.
"""
import argparse
import asyncio
import os

from utils.data_structures.task_config import TaskConfig
from utils.evaluation.evaluator import TaskEvaluator
from utils.general.helper import print_color, read_json

from strands_runner.agent import StrandsTaskAgent
from strands_runner.model import make_model


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--eval_config", default="scripts/eval_config_strands.json")
    parser.add_argument("--task_dir", default="wc-coupon-campaign-gcal-gform")
    parser.add_argument("--model_name", default=None)
    parser.add_argument("--max_steps", type=int, default=None)
    parser.add_argument("--debug", action="store_true")
    # Phase split for groundtruth isolation:
    #   agent — preprocess + run agent, save logs; NO eval (groundtruth absent from this container).
    #   eval  — only evaluate the persisted log against groundtruth (separate container with groundtruth).
    #   all   — legacy single-process behaviour (agent + eval together).
    parser.add_argument("--phase", choices=["all", "agent", "eval"], default="all")
    args = parser.parse_args()

    cfg = read_json(args.eval_config)
    agent_cfg = cfg.get("agent", {})
    model_name = (
        args.model_name
        or os.environ.get("LLM_MODEL")
        or os.environ.get("MODEL_NAME")
        or agent_cfg.get("model_name")
    )
    if not model_name:
        raise SystemExit("LLM_MODEL (or MODEL_NAME / --model_name) must be set")
    provider = "strands"  # static — used only for naming agent_short_name / log dirs
    max_steps = (
        args.max_steps
        or cfg.get("global_task_config", {}).get("max_steps_under_single_turn_mode", 100)
    )
    dump_path = cfg.get("dump_path", "./dumps/")

    global_task_config = {"dump_path": dump_path,
                          "max_steps_under_single_turn_mode": max_steps}

    task_config = TaskConfig.build(
        args.task_dir,
        agent_short_name=f"{provider}/{model_name}",
        global_task_config=global_task_config,
        single_turn_mode=True,
        cn_mode=False,
    )

    print_color(f"====== {args.task_dir} | {provider}/{model_name} | steps={max_steps} ======", "yellow")
    print_color(f"workspace : {task_config.agent_workspace}", "cyan")
    print_color(f"log       : {task_config.log_file}", "cyan")

    if args.phase in ("all", "agent"):
        model = make_model(model_name)
        agent = StrandsTaskAgent(
            task_config=task_config,
            model=model,
            max_steps=max_steps,
            debug=args.debug,
        )
        status = await agent.run()
        print_color(f"\n====== Status: {status.value} ======", "yellow")
        if args.phase == "agent":
            # Eval happens in a separate container that has groundtruth.
            return 0

    print_color("\n====== Evaluating ======", "yellow")
    eval_res = await TaskEvaluator.evaluate_from_log_file(task_config.log_file)
    print(f"Pass:    {eval_res.get('pass', False)}")
    print(f"Details: {eval_res.get('details', 'N/A')}")
    return 0 if eval_res.get("pass", False) else 1


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
