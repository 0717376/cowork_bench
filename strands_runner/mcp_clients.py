"""Build Strands MCPClient objects from YAML configs (or HTTP URLs).

Mirrors utils/mcp/tool_servers.py but returns strands.tools.mcp.MCPClient
instances. Each yaml file under configs/mcp_servers/ declares a server with
fields: name, params{command,args,env,cwd}, client_session_timeout_seconds.
"""
from __future__ import annotations

import os
from datetime import timedelta
from pathlib import Path
from typing import Dict, List, Optional

import yaml
from mcp import StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.client.streamable_http import streamablehttp_client
from strands.tools.mcp import MCPClient


def _resolve(value, local_servers_path: str, agent_workspace: str, task_dir: str = ""):
    if not isinstance(value, str):
        return value
    return (
        value
        .replace("${local_servers_paths}", local_servers_path)
        .replace("${agent_workspace}", agent_workspace)
        .replace("${task_dir}", task_dir)
    )


def _pg_bridge(env: dict) -> dict:
    """Mirror tool_servers.py: bridge libpq PGHOST/... to PG_HOST/... names
    used by emails-mcp, woocommerce-mcp, yahoo-finance-mcp."""
    out = dict(env)
    for src, dst in [
        ("PGHOST", "PG_HOST"), ("PGPORT", "PG_PORT"), ("PGDATABASE", "PG_DATABASE"),
        ("PGUSER", "PG_USER"), ("PGPASSWORD", "PG_PASSWORD"),
    ]:
        if env.get(src) is not None:
            out[dst] = env[src]
    return out


def build_mcp_clients(
    needed_servers: List[str],
    agent_workspace: str,
    config_dir: str = "configs/mcp_servers",
    task_dir: str = "",
    http_mcp_urls: Optional[Dict[str, str]] = None,
    http_mcp_timeout: float = 600.0,
) -> List[MCPClient]:
    """Build MCPClient list for the requested servers. HTTP MCPs (per
    http_mcp_urls) take priority over yaml/stdio definitions of the same name."""
    local_servers_path = os.environ.get("LOCAL_SERVERS_PATH", os.path.abspath("./local_servers"))
    agent_workspace = os.path.abspath(agent_workspace)
    task_dir = os.path.abspath(task_dir) if task_dir else ""

    clients: List[MCPClient] = []
    remaining = list(needed_servers)

    if http_mcp_urls:
        for name, url in http_mcp_urls.items():
            if name not in remaining:
                continue
            clients.append(_http_client(name, url, http_mcp_timeout))
            remaining.remove(name)

    if not remaining:
        return clients

    config_path = Path(config_dir)
    if not config_path.exists():
        raise FileNotFoundError(f"MCP config dir not found: {config_dir}")

    for cfg_file in sorted(config_path.glob("*.yaml")):
        with open(cfg_file, encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
        if not cfg:
            continue
        name = cfg.get("name", cfg_file.stem)
        if name not in remaining:
            continue

        params = cfg.get("params", {})
        resolve = lambda v: _resolve(v, local_servers_path, agent_workspace, task_dir)
        command = resolve(params.get("command", ""))
        args = [resolve(a) for a in params.get("args", [])]
        env = {k: resolve(v) for k, v in params.get("env", {}).items()}
        cwd = resolve(params.get("cwd", agent_workspace))
        os.makedirs(cwd, exist_ok=True)

        # Runtime env wins over yaml defaults (so MCP_HTTP_URLS/PG* overrides land).
        full_env = _pg_bridge({**env, **os.environ})
        timeout = float(cfg.get("client_session_timeout_seconds", 60))
        clients.append(_stdio_client(name, command, args, full_env, cwd, timeout))

    # Warn for unmatched names so a typo in needed_mcp_servers surfaces early.
    seen = set()
    for f in config_path.glob("*.yaml"):
        try:
            doc = yaml.safe_load(open(f, encoding="utf-8"))
            if doc:
                seen.add(doc.get("name", f.stem))
        except Exception:
            pass
    missing = [s for s in remaining if s not in seen]
    if missing:
        print(f"[mcp_clients] Warning: no yaml config found for: {missing}")

    return clients


_STARTUP_TIMEOUT_FLOOR = 300  # Upstream YAML's `client_session_timeout_seconds`
# is a per-call session timeout in CAMEL, not a startup handshake budget. Under
# parallel runs (N=10+ tasks) ~50-70 MCP processes spawn at once; many python
# MCPs use `uv run` which triggers a fresh `uv sync` (rebuild) on first launch
# inside each container. Under load that easily blows past 60s. 300s = generous
# one-time budget; startup happens once per task so cost is negligible.


def _stdio_client(name: str, command: str, args: list, env: dict, cwd: str, timeout: float) -> MCPClient:
    """Strands MCPClient wrapping mcp.client.stdio with our process params."""
    params = StdioServerParameters(command=command, args=list(args), env=env, cwd=cwd or None)
    return MCPClient(
        lambda: stdio_client(params),
        prefix=name,
        startup_timeout=max(_STARTUP_TIMEOUT_FLOOR, int(timeout)),
    )


def _http_client(name: str, url: str, timeout: float) -> MCPClient:
    """Strands MCPClient wrapping the streamable-http transport."""
    td = timedelta(seconds=timeout)
    return MCPClient(
        lambda: streamablehttp_client(url, timeout=td),
        prefix=name,
        startup_timeout=max(_STARTUP_TIMEOUT_FLOOR, int(timeout)),
    )
