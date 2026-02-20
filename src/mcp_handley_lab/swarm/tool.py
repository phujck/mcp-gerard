"""Swarm MCP server for multi-agent orchestration."""

import json
import os
import signal
import subprocess
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.swarm.config import (
    ensure_state_dir,
    list_agent_types,
    load_agent_config,
)
from mcp_handley_lab.swarm.models import (
    AgentInfo,
    AgentStatus,
    AgentTypeConfig,
    LogsResult,
    SpawnedAgent,
    SpawnResult,
)

mcp = FastMCP("Swarm")

AGENTS_FILE = "agents.json"


def _agents_file() -> Path:
    return ensure_state_dir() / AGENTS_FILE


def _load_agents() -> dict[str, SpawnedAgent]:
    """Load agents from state file."""
    path = _agents_file()
    if not path.exists():
        return {}
    with open(path) as f:
        data = json.load(f)
    return {k: SpawnedAgent(**v) for k, v in data.items()}


def _save_agents(agents: dict[str, SpawnedAgent]) -> None:
    """Save agents to state file."""
    path = _agents_file()
    with open(path, "w") as f:
        json.dump({k: v.model_dump(mode="json") for k, v in agents.items()}, f, indent=2, default=str)


def _update_agent_status(agent: SpawnedAgent) -> SpawnedAgent:
    """Check if process is still running and update status."""
    if agent.status != AgentStatus.RUNNING:
        return agent

    try:
        os.kill(agent.pid, 0)  # Check if process exists
    except OSError:
        # Process no longer exists
        agent.status = AgentStatus.COMPLETED
        agent.ended_at = datetime.now()
        # Try to get exit code from output
        if agent.output_file.exists():
            content = agent.output_file.read_text()
            if "error" in content.lower() or "failed" in content.lower():
                agent.status = AgentStatus.FAILED

    return agent


def _build_prompt(config: AgentTypeConfig, task_id: str) -> str:
    """Build the full prompt for an agent."""
    parts = []

    if config.prompt_prefix:
        parts.append(config.prompt_prefix)

    parts.append(f"\n## Your Task\nExecute the task defined in GTD card: `{task_id}`\n")
    parts.append("Read this card first to get your full instructions.")

    if config.prompt_suffix:
        parts.append(config.prompt_suffix)

    # Add permission constraints
    perms = config.permissions
    constraints = []
    if not perms.can_edit:
        constraints.append("- You MUST NOT edit any files")
    if not perms.can_create_files:
        constraints.append("- You MUST NOT create new files")
    if not perms.can_run_bash:
        constraints.append("- You MUST NOT run bash commands")
    if perms.file_patterns:
        patterns = ", ".join(perms.file_patterns)
        constraints.append(f"- You may ONLY edit files matching: {patterns}")

    if constraints:
        parts.append("\n## Constraints\n" + "\n".join(constraints))

    return "\n".join(parts)


@mcp.tool(
    description="Spawn a Claude worker agent to execute a GTD task. Returns immediately with agent ID for monitoring."
)
def spawn(
    task_id: str = Field(..., description="GTD task card UUID to execute"),
    agent_type: str = Field(
        default="worker",
        description="Agent type: worker, reviewer, researcher, coder, tester (or custom)",
    ),
    model: str = Field(
        default="",
        description="Override model (sonnet, opus, haiku). Empty uses agent type default.",
    ),
    timeout: int = Field(
        default=0,
        description="Override timeout in seconds. 0 uses agent type default.",
    ),
) -> SpawnResult:
    """Spawn a new agent process."""
    config = load_agent_config(agent_type)
    model = model or config.model
    timeout = timeout or config.timeout

    agent_id = str(uuid4())[:8]
    state_dir = ensure_state_dir()
    output_file = state_dir / f"{agent_id}.stdout"
    error_file = state_dir / f"{agent_id}.stderr"

    prompt = _build_prompt(config, task_id)

    # Build claude command
    cmd = [
        "claude",
        "-p",  # Print mode - non-interactive
        prompt,
        "--model", model,
        "--output-format", "stream-json",
        "--verbose",  # Required for stream-json with --print
    ]

    # Add tool permissions based on agent config
    perms = config.permissions
    allowed = []
    disallowed = []

    # Always allow reading
    allowed.extend(["Read", "Grep", "Glob"])

    # GTD access for task coordination
    if "gtd" in perms.allowed_mcp_servers:
        allowed.append("mcp__gtd__*")

    # Edit permissions
    if perms.can_edit:
        allowed.extend(["Edit", "Write"])
    else:
        disallowed.extend(["Edit", "Write", "NotebookEdit"])

    # Bash permissions
    if perms.can_run_bash:
        allowed.append("Bash")
    else:
        disallowed.append("Bash")

    # File creation
    if not perms.can_create_files:
        # Write is already handled above
        pass

    if allowed:
        cmd.extend(["--allowedTools", ",".join(allowed)])
    if disallowed:
        cmd.extend(["--disallowedTools", ",".join(disallowed)])

    # Spawn process with output capture
    with open(output_file, "w") as stdout_f, open(error_file, "w") as stderr_f:
        process = subprocess.Popen(
            cmd,
            stdout=stdout_f,
            stderr=stderr_f,
            stdin=subprocess.DEVNULL,
            start_new_session=True,  # Detach from parent
        )

    agent = SpawnedAgent(
        agent_id=agent_id,
        task_id=task_id,
        agent_type=agent_type,
        model=model,
        pid=process.pid,
        status=AgentStatus.RUNNING,
        output_file=output_file,
        error_file=error_file,
    )

    agents = _load_agents()
    agents[agent_id] = agent
    _save_agents(agents)

    return SpawnResult(
        agent_id=agent_id,
        task_id=task_id,
        pid=process.pid,
        output_file=str(output_file),
        message=f"Agent '{agent_type}' spawned with model '{model}' (timeout: {timeout}s)",
    )


@mcp.tool(description="List all spawned agents with their current status.")
def list_agents(
    status_filter: str = Field(
        default="",
        description="Filter by status: running, completed, failed, killed. Empty for all.",
    ),
    include_completed: bool = Field(
        default=False,
        description="Include completed/failed agents (default: only running)",
    ),
) -> list[AgentInfo]:
    """List spawned agents."""
    agents = _load_agents()
    result = []

    for agent in agents.values():
        agent = _update_agent_status(agent)

        # Apply filters
        if status_filter and agent.status.value != status_filter:
            continue
        if not include_completed and agent.status != AgentStatus.RUNNING:
            continue

        runtime = (datetime.now() - agent.started_at).total_seconds()
        if agent.ended_at:
            runtime = (agent.ended_at - agent.started_at).total_seconds()

        result.append(
            AgentInfo(
                agent_id=agent.agent_id,
                task_id=agent.task_id,
                task_title="",  # Could fetch from GTD but adds latency
                agent_type=agent.agent_type,
                model=agent.model,
                status=agent.status,
                started_at=agent.started_at.isoformat(),
                runtime_seconds=round(runtime, 1),
                pid=agent.pid,
            )
        )

    # Save updated statuses
    for info in result:
        if info.agent_id in agents:
            agents[info.agent_id].status = info.status
    _save_agents(agents)

    return result


@mcp.tool(description="Terminate a running agent by ID.")
def kill(
    agent_id: str = Field(..., description="Agent ID to terminate"),
    force: bool = Field(default=False, description="Force kill with SIGKILL instead of SIGTERM"),
) -> str:
    """Kill a running agent."""
    agents = _load_agents()

    if agent_id not in agents:
        return f"Agent '{agent_id}' not found"

    agent = agents[agent_id]
    agent = _update_agent_status(agent)

    if agent.status != AgentStatus.RUNNING:
        return f"Agent '{agent_id}' is not running (status: {agent.status.value})"

    sig = signal.SIGKILL if force else signal.SIGTERM
    try:
        os.kill(agent.pid, sig)
        agent.status = AgentStatus.KILLED
        agent.ended_at = datetime.now()
        agents[agent_id] = agent
        _save_agents(agents)
        return f"Agent '{agent_id}' terminated"
    except OSError as e:
        return f"Failed to kill agent '{agent_id}': {e}"


@mcp.tool(description="Get stdout/stderr logs from an agent.")
def logs(
    agent_id: str = Field(..., description="Agent ID to get logs from"),
    tail: int = Field(default=100, description="Number of lines from end (0 for all)"),
    stream: str = Field(default="both", description="Which stream: stdout, stderr, or both"),
) -> LogsResult:
    """Get agent logs."""
    agents = _load_agents()

    if agent_id not in agents:
        raise ValueError(f"Agent '{agent_id}' not found")

    agent = agents[agent_id]
    agent = _update_agent_status(agent)

    def read_tail(path: Path, n: int) -> str:
        if not path.exists():
            return ""
        content = path.read_text()
        if n <= 0:
            return content
        lines = content.splitlines()
        return "\n".join(lines[-n:])

    stdout = ""
    stderr = ""

    if stream in ("stdout", "both"):
        stdout = read_tail(agent.output_file, tail)
    if stream in ("stderr", "both"):
        stderr = read_tail(agent.error_file, tail)

    return LogsResult(
        agent_id=agent_id,
        status=agent.status,
        stdout=stdout,
        stderr=stderr,
        exit_code=agent.exit_code,
    )


@mcp.tool(description="List available agent types and their configurations.")
def agent_types() -> list[dict]:
    """List available agent types."""
    result = []
    for type_name in list_agent_types():
        config = load_agent_config(type_name)
        result.append({
            "name": config.name,
            "description": config.description,
            "model": config.model,
            "timeout": config.timeout,
            "can_edit": config.permissions.can_edit,
            "can_bash": config.permissions.can_run_bash,
        })
    return result


@mcp.tool(description="Clean up completed/failed agent records and their log files.")
def cleanup(
    keep_days: int = Field(default=1, description="Keep agents newer than this many days"),
    dry_run: bool = Field(default=True, description="Preview what would be deleted"),
) -> dict:
    """Clean up old agent records."""
    agents = _load_agents()
    cutoff = datetime.now().timestamp() - (keep_days * 86400)

    to_delete = []
    for agent_id, agent in agents.items():
        if agent.status == AgentStatus.RUNNING:
            continue
        if agent.started_at.timestamp() < cutoff:
            to_delete.append(agent_id)

    if dry_run:
        return {
            "would_delete": to_delete,
            "count": len(to_delete),
            "dry_run": True,
        }

    deleted_files = []
    for agent_id in to_delete:
        agent = agents.pop(agent_id)
        for path in [agent.output_file, agent.error_file]:
            if path.exists():
                path.unlink()
                deleted_files.append(str(path))

    _save_agents(agents)

    return {
        "deleted_agents": to_delete,
        "deleted_files": deleted_files,
        "count": len(to_delete),
        "dry_run": False,
    }
