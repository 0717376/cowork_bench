"""Local Strands tools for the Toolathlon runner.

Mirrors utils/aux_tools/* (CAMEL versions) as Strands @tool factories. Each
factory binds runtime state (workspace path, done flag) at construction time
so the tools themselves take only LLM-visible arguments.
"""
from __future__ import annotations

import asyncio
import json
import os
import subprocess
import time
import uuid
from typing import Literal

import httpx
from bs4 import BeautifulSoup
from markdownify import markdownify as _html_to_md
from strands import tool


# ── todowrite ──────────────────────────────────────────────────────
def make_todowrite(agent_workspace: str):
    """Strands-port of strands/agent/tools.py::make_todowrite, minus the
    StateBroadcaster. Persists the plan to <workspace>/.todo.json so it
    survives across calls and shows up in artifacts; mirrors to stderr so
    it's visible in run.log."""
    workspace = os.path.abspath(agent_workspace)
    todo_path = os.path.join(workspace, ".todo.json")

    @tool
    def todowrite(todos: list[dict]) -> str:
        """IMPORTANT: Call this tool FIRST, before any text reply, for EVERY user request — including a single trivial step. Always express the plan through this tool, never as plain text.

        This tool creates and manages a structured task list for the current session. It helps you track progress.

        ALWAYS USE:
        - At the start of every request — capture the work as todos (even one step) BEFORE responding with text
        - After receiving new instructions — add them as todos right away
        - As you work — keep the list current

        Step lifecycle: create the step as "in_progress", do the work, then call this tool again to mark it "completed". Do this even for a one-step plan so progress stays visible.

        Each todo: {"content": "description", "status": "pending|in_progress|completed|cancelled", "priority": "high|medium|low"}

        Always pass the FULL updated list (not just changed items). Mark only ONE task as in_progress at a time.
        """
        os.makedirs(workspace, exist_ok=True)
        try:
            with open(todo_path, "w", encoding="utf-8") as f:
                json.dump(todos, f, ensure_ascii=False, indent=2)
        except OSError as e:
            return f"todowrite: failed to persist plan: {e}"

        # stderr — visible in run.log without polluting MCP stdout channels.
        import sys
        print(f"[todowrite] {len(todos)} items", file=sys.stderr, flush=True)

        active = [t for t in todos if t.get("status") not in ("completed", "cancelled")]
        return f"{len(active)} active todos\n" + json.dumps(todos, ensure_ascii=False, indent=2)

    return todowrite


# ── sleep ──────────────────────────────────────────────────────────
@tool
async def sleep(seconds: float = 1) -> str:
    """Sleep for the given number of seconds.

    Args:
        seconds: Number of seconds to sleep (default 1).
    """
    await asyncio.sleep(seconds)
    return f"Slept {seconds} seconds."


# ── python_execute ─────────────────────────────────────────────────
def make_python_execute(agent_workspace: str):
    workspace = os.path.abspath(agent_workspace)

    @tool
    def python_execute(code: str, filename: str = "", timeout: int = 30) -> str:
        """Execute Python code in the agent workspace and return stdout/stderr.

        Args:
            code: Python source code to execute.
            filename: Optional filename (with .py). A random UUID name is used if omitted.
            timeout: Max execution time in seconds (capped at 120).
        """
        eff_timeout = min(int(timeout), 120)
        fn = filename or f"{uuid.uuid4()}.py"
        if not fn.endswith(".py"):
            fn += ".py"

        tmp_dir = os.path.join(workspace, ".python_tmp")
        os.makedirs(tmp_dir, exist_ok=True)
        file_path = os.path.join(tmp_dir, fn)
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(code)

        cmd = f"uv run --directory {workspace} ./.python_tmp/{fn}"
        start = time.time()
        try:
            r = subprocess.run(
                cmd, shell=True, capture_output=True,
                text=True, encoding="utf-8", timeout=eff_timeout,
            )
        except subprocess.TimeoutExpired:
            return f"=== TIMEOUT ===\nExceeded {eff_timeout}s limit."

        elapsed = time.time() - start
        parts = []
        if r.stdout:
            parts += ["=== STDOUT ===", r.stdout.rstrip()]
        if r.stderr:
            parts += ["=== STDERR ===", r.stderr.rstrip()]
        parts += [
            "=== INFO ===",
            f"Return code: {r.returncode}",
            f"Time: {elapsed:.2f}s / {eff_timeout}s limit",
        ]
        return "\n".join(parts) if parts else "No output."

    return python_execute


# ── overlong-output tools ──────────────────────────────────────────
_OVERLONG_DIR = ".overlong_tool_outputs"
_PAGE_SIZE = 10000


def make_overlong_tools(agent_workspace: str):
    overlong_dir = os.path.join(os.path.abspath(agent_workspace), _OVERLONG_DIR)

    def _dir():
        os.makedirs(overlong_dir, exist_ok=True)
        return overlong_dir

    @tool
    def save_overlong_output(content: str, label: str = "") -> str:
        """Save a large text to disk and return a reference ID.

        Args:
            content: The large text content to store.
            label: Optional human-readable label for this output.
        """
        fid = str(uuid.uuid4())[:8]
        path = os.path.join(_dir(), f"{fid}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(
                {"id": fid, "label": label, "content": content, "saved_at": time.time()},
                f,
            )
        preview = content[:200] + ("..." if len(content) > 200 else "")
        return (
            f"Saved {len(content)} chars as [{fid}] label='{label}'.\n"
            f"Preview: {preview}\n"
            f"Use view_overlong_output(id='{fid}') to read."
        )

    @tool
    def view_overlong_output(id: str, page: int = 0) -> str:
        """View a saved overlong output by ID, paginated.

        Args:
            id: The reference ID returned by save_overlong_output.
            page: Page number (0-indexed, each page ~10000 chars).
        """
        path = os.path.join(_dir(), f"{id}.json")
        if not os.path.exists(path):
            return f"No saved output with id '{id}'."
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        content = data["content"]
        total_pages = max(1, (len(content) + _PAGE_SIZE - 1) // _PAGE_SIZE)
        start = page * _PAGE_SIZE
        chunk = content[start: start + _PAGE_SIZE]
        return (
            f"[{id}] label='{data['label']}' | page {page+1}/{total_pages} | "
            f"chars {start}–{start+len(chunk)} of {len(content)}\n\n{chunk}"
        )

    return save_overlong_output, view_overlong_output


# ── webfetch (ported from strands/agent/tools.py:58-118) ───────────
_WEBFETCH_MAX_RESPONSE_SIZE = 5 * 1024 * 1024
_WEBFETCH_DEFAULT_TIMEOUT = 30
_WEBFETCH_MAX_TIMEOUT = 120
_WEBFETCH_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36"
)
_WEBFETCH_ACCEPT = {
    "markdown": "text/markdown;q=1.0, text/x-markdown;q=0.9, text/plain;q=0.8, text/html;q=0.7, */*;q=0.1",
    "text":     "text/plain;q=1.0, text/markdown;q=0.9, text/html;q=0.8, */*;q=0.1",
    "html":     "text/html;q=1.0, application/xhtml+xml;q=0.9, text/plain;q=0.8, text/markdown;q=0.7, */*;q=0.1",
}
_WEBFETCH_STRIP = ["script", "style", "noscript", "iframe", "object", "embed", "meta", "link"]


def _wf_text(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(_WEBFETCH_STRIP):
        t.decompose()
    return soup.get_text(separator=" ", strip=True)


def _wf_markdown(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(_WEBFETCH_STRIP):
        t.decompose()
    return _html_to_md(str(soup), heading_style="ATX", bullets="-", code_language="").strip()


@tool
def webfetch(
    url: str,
    format: Literal["markdown", "text", "html"] = "markdown",
    timeout: int | None = None,
) -> str:
    """Fetches content from a specified URL and returns it in the requested format.

    Use this tool when you need to retrieve and analyze web content — articles, docs, API responses.

    Usage notes:
    - The URL must be a fully-formed valid http:// or https:// URL
    - Format options: "markdown" (default), "text", or "html"
    - This tool is read-only and does not modify any files
    - Response size limit: 5MB
    """
    if not (url.startswith("http://") or url.startswith("https://")):
        return "Error: URL must start with http:// or https://"

    eff = min(timeout or _WEBFETCH_DEFAULT_TIMEOUT, _WEBFETCH_MAX_TIMEOUT)
    headers = {
        "User-Agent": _WEBFETCH_UA,
        "Accept": _WEBFETCH_ACCEPT.get(format, "*/*"),
        "Accept-Language": "en-US,en;q=0.9",
    }
    try:
        with httpx.Client(follow_redirects=True, timeout=eff) as client:
            r = client.get(url, headers=headers)
            if r.status_code == 403 and r.headers.get("cf-mitigated") == "challenge":
                r = client.get(url, headers={**headers, "User-Agent": "strands-webfetch"})
    except httpx.TimeoutException:
        return f"Error: Request timed out after {eff}s"
    except httpx.RequestError as e:
        return f"Error: Request failed — {e}"

    if r.status_code != 200:
        return f"Error: Request failed with status code {r.status_code}"
    cl = r.headers.get("content-length")
    if cl and int(cl) > _WEBFETCH_MAX_RESPONSE_SIZE:
        return "Error: Response too large (exceeds 5MB limit)"
    if len(r.content) > _WEBFETCH_MAX_RESPONSE_SIZE:
        return "Error: Response too large (exceeds 5MB limit)"

    content_type = r.headers.get("content-type", "").lower()
    mime = content_type.split(";")[0].strip()
    if mime.startswith("image/") and mime != "image/svg+xml":
        return f"[Binary image response: {mime}, {len(r.content)} bytes — not displayed]"

    text = r.text
    is_html = "text/html" in content_type or text[:200].lstrip().lower().startswith(("<html", "<!doctype"))
    if format == "html":
        return text
    if format == "text":
        return _wf_text(text) if is_html else text
    return _wf_markdown(text) if is_html else text


def build_local_tools(workspace: str, needed: list[str]):
    """Return Strands tools for the names in `needed`.

    `claim_done` is intentionally NOT provided: we use `stop_reason=="end_turn"`
    as the completion signal (production Strands behavior) instead of a
    benchmark-specific sentinel tool. Task prompts that reference claim_done
    are scrubbed in StrandsTaskAgent._build_system_prompt.
    """
    needed_set = set(needed or [])
    tools = []
    if "python_execute" in needed_set:
        tools.append(make_python_execute(workspace))
    if "handle_overlong_tool_outputs" in needed_set:
        save_fn, view_fn = make_overlong_tools(workspace)
        tools += [save_fn, view_fn]
    if "sleep" in needed_set:
        tools.append(sleep)
    # webfetch is always-on — cheap, no state, useful for almost every task.
    tools.append(webfetch)
    # todowrite is always-on — required by the Russian global_system_bench
    # prompt ("план фиксируй через todowrite"). Cheap, no state outside workspace.
    tools.append(make_todowrite(workspace))
    return tools
