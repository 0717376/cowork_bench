#!/bin/bash
# Hard-stop run_parallel.sh and clean up all spawned agent/pg containers + networks.
#
# Usage:
#   bash scripts/stop_parallel.sh
#
# Why this exists: Ctrl+C in the run_parallel.sh terminal normally triggers
# the EXIT trap which cleans up. But if the trap doesn't fire (SIGKILL,
# disconnected terminal, hung subshell), this script is the manual fallback.
# Safe to run when nothing is active — it just does nothing.

set -u

echo "[stop_parallel] killing run_parallel.sh ..."
pkill -f "run_parallel.sh" 2>/dev/null || true

sleep 2

containers=$(docker ps -a --format '{{.Names}}' | grep -E '^(agent|eval|pg)-' || true)
if [ -n "$containers" ]; then
    n=$(echo "$containers" | wc -l | tr -d ' ')
    echo "[stop_parallel] removing $n containers ..."
    echo "$containers" | xargs -I{} docker rm -fv {} >/dev/null 2>&1 || true
fi

networks=$(docker network ls --format '{{.Name}}' | grep '^net-' || true)
if [ -n "$networks" ]; then
    n=$(echo "$networks" | wc -l | tr -d ' ')
    echo "[stop_parallel] removing $n networks ..."
    echo "$networks" | xargs -I{} docker network rm {} >/dev/null 2>&1 || true
fi

left_procs=$(ps aux | grep -E "run_parallel|run_containerized" | grep -v grep | wc -l | tr -d ' ')
left_conts=$(docker ps -a --format '{{.Names}}' | grep -cE '^(agent|eval|pg)-' || echo 0)

echo "[stop_parallel] done. processes left: $left_procs, containers left: $left_conts"
