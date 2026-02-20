#!/usr/bin/env python3
"""Swarm status monitor - shows running Claude agents and their tasks.

Usage:
    swarm-status           # Show current status once
    swarm-status --watch   # Continuously update (every 2s)
    swarm-status --logs ID # Show logs for specific agent
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# State directory matches tool.py
STATE_DIR = Path.home() / ".local" / "state" / "claude-swarm"
AGENTS_FILE = STATE_DIR / "agents.json"


def load_agents() -> dict:
    """Load agents from state file."""
    if not AGENTS_FILE.exists():
        return {}
    with open(AGENTS_FILE) as f:
        return json.load(f)


def check_process_running(pid: int) -> bool:
    """Check if a process is still running."""
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def format_duration(seconds: float) -> str:
    """Format duration as human-readable string."""
    if seconds < 60:
        return f"{seconds:.0f}s"
    elif seconds < 3600:
        return f"{seconds / 60:.1f}m"
    else:
        return f"{seconds / 3600:.1f}h"


def get_gtd_task_title(task_id: str) -> str:
    """Try to get task title from GTD (best effort)."""
    try:
        import subprocess
        result = subprocess.run(
            ["gtd", "show", task_id, "--format", "json"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            data = json.loads(result.stdout)
            return data.get("title", "")[:50]
    except Exception:
        pass
    return ""


def print_status_table(agents: dict, show_completed: bool = False):
    """Print a formatted status table."""
    # Header
    print("\n" + "=" * 80)
    print("SWARM STATUS".center(80))
    print("=" * 80)

    running = []
    completed = []

    for agent_id, agent in agents.items():
        pid = agent.get("pid", 0)
        is_running = check_process_running(pid)

        started = datetime.fromisoformat(agent["started_at"])
        runtime = (datetime.now() - started).total_seconds()

        status = "RUNNING" if is_running else agent.get("status", "completed").upper()

        entry = {
            "id": agent_id,
            "task": agent.get("task_id", "")[:8],
            "type": agent.get("agent_type", "worker"),
            "model": agent.get("model", "sonnet"),
            "status": status,
            "runtime": format_duration(runtime),
            "pid": pid,
        }

        if is_running:
            running.append(entry)
        else:
            completed.append(entry)

    # Running agents
    if running:
        print(f"\n{'RUNNING AGENTS':^80}")
        print("-" * 80)
        print(f"{'ID':<10} {'Task':<10} {'Type':<12} {'Model':<8} {'Runtime':<10} {'PID':<8}")
        print("-" * 80)
        for a in running:
            print(
                f"{a['id']:<10} {a['task']:<10} {a['type']:<12} {a['model']:<8} "
                f"{a['runtime']:<10} {a['pid']:<8}"
            )
    else:
        print("\n  No running agents")

    # Completed agents (if requested)
    if show_completed and completed:
        print(f"\n{'COMPLETED AGENTS':^80}")
        print("-" * 80)
        print(f"{'ID':<10} {'Task':<10} {'Type':<12} {'Status':<10} {'Runtime':<10}")
        print("-" * 80)
        for a in completed[-10:]:  # Last 10 only
            print(
                f"{a['id']:<10} {a['task']:<10} {a['type']:<12} {a['status']:<10} {a['runtime']:<10}"
            )

    print("\n" + "=" * 80)
    print(f"  Running: {len(running)}  |  Completed: {len(completed)}  |  Total: {len(agents)}")
    print("=" * 80 + "\n")


def print_agent_logs(agent_id: str, tail: int = 50):
    """Print logs for a specific agent."""
    agents = load_agents()

    if agent_id not in agents:
        # Try partial match
        matches = [k for k in agents if k.startswith(agent_id)]
        if len(matches) == 1:
            agent_id = matches[0]
        elif len(matches) > 1:
            print(f"Ambiguous agent ID. Matches: {matches}")
            return
        else:
            print(f"Agent '{agent_id}' not found")
            return

    agent = agents[agent_id]
    stdout_file = Path(agent.get("output_file", ""))
    stderr_file = Path(agent.get("error_file", ""))

    print(f"\n{'=' * 80}")
    print(f"LOGS FOR AGENT: {agent_id}")
    print(f"Task: {agent.get('task_id', 'unknown')}")
    print(f"Type: {agent.get('agent_type', 'unknown')} | Model: {agent.get('model', 'unknown')}")
    print(f"{'=' * 80}\n")

    if stdout_file.exists():
        content = stdout_file.read_text()
        lines = content.splitlines()
        if tail > 0:
            lines = lines[-tail:]
        print("--- STDOUT ---")
        print("\n".join(lines) if lines else "(empty)")
    else:
        print("--- STDOUT ---\n(no output file)")

    print()

    if stderr_file.exists():
        content = stderr_file.read_text()
        lines = content.splitlines()
        if tail > 0:
            lines = lines[-tail:]
        if lines:
            print("--- STDERR ---")
            print("\n".join(lines))


def watch_mode(interval: float = 2.0, show_completed: bool = False):
    """Continuously display status."""
    try:
        while True:
            # Clear screen
            print("\033[2J\033[H", end="")

            agents = load_agents()
            print_status_table(agents, show_completed)
            print(f"  [Updating every {interval}s - Press Ctrl+C to exit]")

            time.sleep(interval)
    except KeyboardInterrupt:
        print("\n\nExiting watch mode.")


def main():
    parser = argparse.ArgumentParser(
        description="Monitor Claude swarm agents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  swarm-status                 Show current status
  swarm-status -w              Watch mode (updates every 2s)
  swarm-status -l abc123       Show logs for agent abc123
  swarm-status -a              Include completed agents
        """,
    )
    parser.add_argument("-w", "--watch", action="store_true", help="Watch mode (continuous updates)")
    parser.add_argument("-i", "--interval", type=float, default=2.0, help="Watch interval in seconds")
    parser.add_argument("-l", "--logs", metavar="ID", help="Show logs for specific agent")
    parser.add_argument("-t", "--tail", type=int, default=50, help="Number of log lines to show")
    parser.add_argument("-a", "--all", action="store_true", help="Include completed agents")

    args = parser.parse_args()

    if args.logs:
        print_agent_logs(args.logs, args.tail)
    elif args.watch:
        watch_mode(args.interval, args.all)
    else:
        agents = load_agents()
        print_status_table(agents, args.all)


if __name__ == "__main__":
    main()
