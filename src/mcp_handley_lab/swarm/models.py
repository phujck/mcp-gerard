"""Data models for swarm agent management."""

from datetime import datetime
from enum import Enum
from pathlib import Path
from pydantic import BaseModel, Field


class AgentStatus(str, Enum):
    """Status of a spawned agent."""

    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    KILLED = "killed"


class AgentPermissions(BaseModel):
    """Permissions for an agent type."""

    can_edit: bool = Field(default=True, description="Can edit files")
    can_create_files: bool = Field(default=True, description="Can create new files")
    can_delete_files: bool = Field(default=False, description="Can delete files")
    can_run_bash: bool = Field(default=True, description="Can run bash commands")
    allowed_mcp_servers: list[str] = Field(
        default_factory=lambda: ["gtd"], description="MCP servers the agent can access"
    )
    file_patterns: list[str] | None = Field(
        default=None, description="Glob patterns for allowed file edits (None = all)"
    )


class AgentTypeConfig(BaseModel):
    """Configuration for an agent type."""

    name: str = Field(..., description="Agent type name")
    description: str = Field(default="", description="What this agent type does")
    model: str = Field(default="sonnet", description="Claude model to use")
    timeout: int = Field(default=300, description="Timeout in seconds")
    permissions: AgentPermissions = Field(default_factory=AgentPermissions)
    prompt_prefix: str = Field(
        default="", description="Prompt prefix injected before task instructions"
    )
    prompt_suffix: str = Field(
        default="", description="Prompt suffix injected after task instructions"
    )


class SpawnedAgent(BaseModel):
    """A spawned agent process."""

    agent_id: str = Field(..., description="Unique agent identifier")
    task_id: str = Field(..., description="GTD task card UUID")
    agent_type: str = Field(..., description="Agent type name")
    model: str = Field(..., description="Claude model being used")
    pid: int = Field(..., description="Process ID")
    status: AgentStatus = Field(default=AgentStatus.RUNNING)
    started_at: datetime = Field(default_factory=datetime.now)
    ended_at: datetime | None = Field(default=None)
    exit_code: int | None = Field(default=None)
    output_file: Path = Field(..., description="Path to stdout capture file")
    error_file: Path = Field(..., description="Path to stderr capture file")


class SpawnResult(BaseModel):
    """Result of spawning an agent."""

    agent_id: str = Field(..., description="Unique agent identifier")
    task_id: str = Field(..., description="GTD task card UUID")
    pid: int = Field(..., description="Process ID")
    output_file: str = Field(..., description="Path to stdout capture")
    message: str = Field(default="Agent spawned successfully")


class AgentInfo(BaseModel):
    """Summary info about an agent for listing."""

    agent_id: str
    task_id: str
    task_title: str = Field(default="")
    agent_type: str
    model: str
    status: AgentStatus
    started_at: str
    runtime_seconds: float
    pid: int


class LogsResult(BaseModel):
    """Result of fetching agent logs."""

    agent_id: str
    status: AgentStatus
    stdout: str = Field(default="")
    stderr: str = Field(default="")
    exit_code: int | None = Field(default=None)
