"""StrandsTaskAgent — Strands-based runner with the same contract as
utils/roles/task_agent.TaskAgent (init(task_config, model, max_steps),
async run() -> TaskStatus). The log format on disk matches CAMEL's exactly
so utils/evaluation/evaluator.py works unchanged.
"""
from __future__ import annotations

import contextlib
import json
import os
import re
import shutil
import subprocess
import traceback
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Optional

from strands import Agent, tool
from strands.agent.conversation_manager.summarizing_conversation_manager import (
    DEFAULT_SUMMARIZATION_PROMPT,
    SummarizingConversationManager,
)

from utils.data_structures.task_config import TaskConfig
from utils.general.helper import copy_folder_contents, print_color

from .local_tools import build_local_tools
from .mcp_clients import build_mcp_clients


@tool
def _summary_noop() -> str:
    """Заглушка: даёт summary-агенту непустой tools (некоторые шлюзы отвергают tools:[])."""
    return "ok"


class TaskStatus(Enum):
    SUCCESS = "success"
    FAILED = "failed"
    MAX_TURNS_REACHED = "max_turns_reached"
    INTERRUPTED = "interrupted"


# Phase 2: Russian global system prompt is loaded from prompts/global_system_bench.md
# and concatenated with the task-specific system prompt:
#     <now_line> + <global_system_bench> + <task.system_prompts.agent>
# This mirrors how strands/agent/core.py assembles its prompts (see core.py:299-310).
_GLOBAL_SYS_PROMPT_PATH = Path(__file__).parent / "prompts" / "global_system_bench.md"

# Fallback used if global_system_bench.md is missing (should not happen in production).
_FALLBACK_SYS_PROMPT_TMPL = (
    "You are a helpful AI assistant. Your workspace directory is: {workspace}\n"
    "Complete the user's task using the provided tools, then end your turn."
)


def _load_global_system_bench() -> str:
    try:
        return _GLOBAL_SYS_PROMPT_PATH.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return ""


# Lines mentioning claim_done in task system prompts get scrubbed — we use
# end_turn as the completion signal, not a sentinel tool. Pattern catches
# both "claim_done" and "local-claim_done" with surrounding context.
_CLAIM_DONE_LINE_RE = re.compile(r"^.*claim[_-]?done.*$\n?", flags=re.IGNORECASE | re.MULTILINE)


def _serialize_messages(messages) -> list:
    """Strands Messages → JSON-friendly list. Strands content is a list of
    blocks ({text}, {toolUse}, {toolResult}); we serialize as-is so traj.json
    captures the full multi-modal trajectory."""
    out = []
    for m in messages or []:
        try:
            out.append({"role": getattr(m, "role", None) or m.get("role"),
                        "content": getattr(m, "content", None) or m.get("content")})
        except Exception:
            out.append(str(m))
    # default=str handles dataclasses, datetime, bytes-fallback inside content blocks
    return json.loads(json.dumps(out, default=str, ensure_ascii=False))


def _extract_tool_calls(messages) -> list:
    """Flatten toolUse blocks across the trajectory — convenient for analysis,
    matches the shape evaluator's traj.json tool_calls field used to take."""
    calls = []
    for m in messages or []:
        content = getattr(m, "content", None) or (m.get("content") if isinstance(m, dict) else None) or []
        for blk in content:
            if not isinstance(blk, dict):
                continue
            if "toolUse" in blk:
                calls.append(blk["toolUse"])
    return json.loads(json.dumps(calls, default=str, ensure_ascii=False))


class StrandsTaskAgent:
    def __init__(self, task_config: TaskConfig, model, max_steps: int = 100, debug: bool = False):
        self.task_config = task_config
        self.model = model
        self.max_steps = max_steps  # Strands has no built-in turn cap; kept for API parity / future hook
        self.debug = debug
        self._workspace: Optional[str] = None

    async def _setup_workspace(self) -> str:
        workspace = os.path.abspath(self.task_config.agent_workspace)
        # Drop stale artifacts from previous runs (fixed dump path is reused).
        if os.path.isdir(workspace):
            shutil.rmtree(workspace, ignore_errors=True)
        os.makedirs(workspace, exist_ok=True)
        init = self.task_config.initialization
        if init and init.workspace and os.path.exists(str(init.workspace)):
            await copy_folder_contents(str(init.workspace), workspace)
        for srv, d in [("arxiv_local", "arxiv_local_storage"),
                       ("memory", "memory"),
                       ("playwright_with_chunk", ".playwright_output")]:
            if srv in self.task_config.needed_mcp_servers:
                os.makedirs(os.path.join(workspace, d), exist_ok=True)
        return workspace

    def _run_preprocess(self):
        init = self.task_config.initialization
        if not (init and init.process_command):
            return
        cmd = init.process_command
        cmd += f" --agent_workspace {self.task_config.agent_workspace}"
        lt = self.task_config.launch_time or ""
        lt_clean = " ".join(lt.split()[:2])
        cmd += f" --launch_time \"{lt_clean}\""
        print_color("[preprocess] running...", "yellow")
        r = subprocess.run(cmd, shell=True, capture_output=not self.debug, text=True)
        if r.returncode != 0:
            print_color(f"[preprocess] failed:\n{r.stderr or ''}", "red")
            raise RuntimeError(f"preprocess failed (exit {r.returncode})")
        print_color("[preprocess] done.", "green")

    def _build_system_prompt(self, workspace: str) -> str:
        # Phase 2 — Russian. Layered: now_line + global_system_bench + task prompt.
        # Order matches strands/agent/core.py:299-310 (date → global → task).
        now_line = f"Текущее время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.\n"

        global_bench = _load_global_system_bench()
        if not global_bench:
            global_bench = _FALLBACK_SYS_PROMPT_TMPL.format(workspace=workspace)

        sp = self.task_config.system_prompts
        task_prompt = ""
        if sp and sp.agent:
            # Scrub any claim_done instructions — we use stop_reason=="end_turn"
            # as the completion signal, not a benchmark-specific sentinel tool.
            task_prompt = _CLAIM_DONE_LINE_RE.sub("", sp.agent).strip()

        workspace_line = f"\nРабочая директория агента: {workspace}\n"

        parts = [now_line.rstrip(), global_bench.rstrip(), workspace_line.rstrip()]
        if task_prompt:
            parts.append(task_prompt.rstrip())
        return "\n\n".join(parts)

    def _save_log(self, status: TaskStatus, start_time: datetime,
                  result=None, messages=None):
        log_path = self.task_config.log_file
        if not log_path:
            return
        Path(log_path).parent.mkdir(parents=True, exist_ok=True)

        # traj_log.json: evaluator reads this — needs `config` + `status`.
        record = {
            "config": self.task_config.to_dict(),
            "status": status.value,
            "start_time": start_time.isoformat(),
            "end_time": datetime.now().isoformat(),
        }
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)

        # traj.json: full trajectory for analysis (not read by evaluator).
        traj_path = str(Path(log_path).parent / "traj.json")
        traj = {
            "status": status.value,
            "start_time": start_time.isoformat(),
            "end_time": datetime.now().isoformat(),
            "stop_reason": getattr(result, "stop_reason", None) if result else None,
            "messages": _serialize_messages(messages),
            "tool_calls": _extract_tool_calls(messages),
        }
        with open(traj_path, "w", encoding="utf-8") as f:
            json.dump(traj, f, ensure_ascii=False, indent=2, default=str)

    async def run(self) -> TaskStatus:
        start_time = datetime.now()
        status = TaskStatus.FAILED
        result = None
        agent: Optional[Agent] = None

        try:
            workspace = await self._setup_workspace()
            self._workspace = workspace
            self._run_preprocess()

            task_src_dir = os.path.abspath(os.path.join("tasks/finalpool", self.task_config.task_dir))
            http_mcp_urls = None
            http_mcp_timeout = float(os.environ.get("MCP_HTTP_TIMEOUT", "600"))
            raw_urls = os.environ.get("MCP_HTTP_URLS", "")
            if raw_urls:
                http_mcp_urls = {}
                for item in raw_urls.split(","):
                    if "=" in item:
                        k, v = item.split("=", 1)
                        http_mcp_urls[k.strip()] = v.strip()

            mcp_clients = build_mcp_clients(
                self.task_config.needed_mcp_servers, workspace,
                task_dir=task_src_dir,
                http_mcp_urls=http_mcp_urls,
                http_mcp_timeout=http_mcp_timeout,
            )

            # Strands MCPClient is a context manager — ExitStack composes N of them.
            with contextlib.ExitStack() as stack:
                for c in mcp_clients:
                    stack.enter_context(c)
                mcp_tools = []
                for c in mcp_clients:
                    mcp_tools.extend(c.list_tools_sync())

                local_tools = build_local_tools(
                    workspace, self.task_config.needed_local_tools or [],
                )
                all_tools = mcp_tools + local_tools

                print_color(
                    f"[strands] Total tools: {len(all_tools)} (MCP: {len(mcp_tools)}, local: {len(local_tools)})",
                    "cyan",
                )
                tool_names = []
                for t in all_tools:
                    n = getattr(t, "tool_name", None) or getattr(t, "name", None) or repr(t)
                    tool_names.append(n)
                print_color(f"[strands] Tool names: {tool_names}", "cyan")

                sys_prompt = self._build_system_prompt(workspace)
                summarizer = Agent(
                    model=self.model,
                    system_prompt=DEFAULT_SUMMARIZATION_PROMPT,
                    tools=[_summary_noop],
                )
                agent = Agent(
                    model=self.model,
                    system_prompt=sys_prompt,
                    tools=all_tools,
                    conversation_manager=SummarizingConversationManager(
                        summarization_agent=summarizer,
                        summary_ratio=0.3,
                        preserve_recent_messages=10,
                        proactive_compression={"compression_threshold": 0.7},
                    ),
                )

                task_str = self.task_config.task_str
                print_color(f"\n[task] {task_str[:300]}\n", "yellow")
                result = await agent.invoke_async(task_str)

                stop_reason = getattr(result, "stop_reason", None)
                if self.debug:
                    print_color(f"[strands] stop_reason={stop_reason}", "cyan")

                # Completion signal: model voluntarily ended its turn. The
                # evaluator (eval_res.json) decides whether the produced
                # artifacts actually solve the task — orchestration success
                # and task-correctness are separated on purpose.
                if stop_reason == "end_turn":
                    status = TaskStatus.SUCCESS
                    print_color("[strands] Agent ended turn cleanly.", "green")
                elif stop_reason == "max_tokens":
                    status = TaskStatus.MAX_TURNS_REACHED
                    print_color("[strands] Hit max_tokens.", "yellow")
                else:
                    print_color(f"[strands] Unexpected stop_reason={stop_reason}", "red")

        except KeyboardInterrupt:
            status = TaskStatus.INTERRUPTED
        except Exception as e:
            print_color(f"[strands] Error: {e}", "red")
            if self.debug:
                traceback.print_exc()

        messages = getattr(agent, "messages", None) if agent is not None else None
        self._save_log(status, start_time, result=result, messages=messages)
        return status
