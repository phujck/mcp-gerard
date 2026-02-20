"""Agent type configuration management."""

import tomllib
from pathlib import Path

from mcp_handley_lab.swarm.models import AgentPermissions, AgentTypeConfig

CONFIG_DIR = Path.home() / ".config" / "claude-swarm" / "agents"
STATE_DIR = Path.home() / ".local" / "state" / "claude-swarm"

# Built-in agent types (used if no config file exists)
BUILTIN_AGENTS: dict[str, AgentTypeConfig] = {
    "worker": AgentTypeConfig(
        name="worker",
        description="General purpose worker agent with full capabilities",
        model="sonnet",
        timeout=600,
        permissions=AgentPermissions(
            can_edit=True,
            can_create_files=True,
            can_delete_files=False,
            can_run_bash=True,
            allowed_mcp_servers=["gtd"],
        ),
        prompt_prefix="""You are a worker agent in a swarm system. Your task is defined in a GTD card.

## Workflow
1. Read your task card using mcp__gtd__read with op="get"
2. Understand data.swarm.instruction and data.swarm.context
3. Execute the task
4. Update the card with your result using mcp__gtd__update

## Reporting Results
When done, update the card:
- Set data.swarm.result to a summary of what you did
- Set data.swarm.files_changed to list of files modified
- Change tags: +swarm.done -swarm.pending

## Rules
- Do ONLY what the instruction says
- If blocked, set data.swarm.result to explain why and mark done anyway
- Keep changes atomic and minimal
""",
    ),
    "reviewer": AgentTypeConfig(
        name="reviewer",
        description="Read-only reviewer for code analysis and feedback",
        model="sonnet",
        timeout=300,
        permissions=AgentPermissions(
            can_edit=False,
            can_create_files=False,
            can_delete_files=False,
            can_run_bash=False,
            allowed_mcp_servers=["gtd"],
        ),
        prompt_prefix="""You are a code reviewer agent. You may ONLY read and analyze - never edit files.

## Workflow
1. Read your task card using mcp__gtd__read with op="get"
2. Review the code/files specified in data.swarm.context
3. Provide detailed feedback in data.swarm.result

## Rules
- READ ONLY - do not attempt to edit any files
- Be specific about issues found with file paths and line numbers
- Suggest improvements but do not implement them
""",
    ),
    "researcher": AgentTypeConfig(
        name="researcher",
        description="Research agent for codebase exploration and information gathering",
        model="sonnet",
        timeout=300,
        permissions=AgentPermissions(
            can_edit=False,
            can_create_files=False,
            can_delete_files=False,
            can_run_bash=False,
            allowed_mcp_servers=["gtd"],
        ),
        prompt_prefix="""You are a research agent. Gather information and report findings.

## Workflow
1. Read your task card for the research question
2. Explore the codebase to find answers
3. Report findings in data.swarm.result

## Rules
- READ ONLY - never edit files
- Be thorough but concise in your findings
- Include file paths and specific code references
""",
    ),
    "coder": AgentTypeConfig(
        name="coder",
        description="Implementation agent for writing and modifying code",
        model="sonnet",
        timeout=600,
        permissions=AgentPermissions(
            can_edit=True,
            can_create_files=True,
            can_delete_files=False,
            can_run_bash=True,
            allowed_mcp_servers=["gtd"],
        ),
        prompt_prefix="""You are a coding agent. Implement the requested changes.

## Workflow
1. Read your task card for implementation instructions
2. Read relevant files for context
3. Make the requested changes
4. Report what you changed in data.swarm.result

## Rules
- Make ONLY the requested changes
- Follow existing code patterns and style
- Test your changes if tests exist
- List all modified files in data.swarm.files_changed
""",
    ),
    "tester": AgentTypeConfig(
        name="tester",
        description="Testing agent for running and validating tests",
        model="sonnet",
        timeout=300,
        permissions=AgentPermissions(
            can_edit=False,
            can_create_files=False,
            can_delete_files=False,
            can_run_bash=True,
            allowed_mcp_servers=["gtd"],
        ),
        prompt_prefix="""You are a testing agent. Run tests and report results.

## Workflow
1. Read your task card for what to test
2. Run the appropriate test commands
3. Report results in data.swarm.result

## Rules
- Do NOT edit test files
- Report all failures with details
- Include full error messages for failing tests
""",
    ),
}


def load_agent_config(agent_type: str) -> AgentTypeConfig:
    """Load agent type configuration from file or use builtin."""
    config_file = CONFIG_DIR / f"{agent_type}.toml"

    if config_file.exists():
        with open(config_file, "rb") as f:
            data = tomllib.load(f)

        permissions = AgentPermissions(**data.get("permissions", {}))
        return AgentTypeConfig(
            name=agent_type,
            description=data.get("description", ""),
            model=data.get("model", "sonnet"),
            timeout=data.get("timeout", 300),
            permissions=permissions,
            prompt_prefix=data.get("prompt", {}).get("prefix", ""),
            prompt_suffix=data.get("prompt", {}).get("suffix", ""),
        )

    if agent_type in BUILTIN_AGENTS:
        return BUILTIN_AGENTS[agent_type]

    # Default to worker config for unknown types
    config = BUILTIN_AGENTS["worker"].model_copy()
    config.name = agent_type
    return config


def list_agent_types() -> list[str]:
    """List all available agent types (builtin + config files)."""
    types = set(BUILTIN_AGENTS.keys())

    if CONFIG_DIR.exists():
        for config_file in CONFIG_DIR.glob("*.toml"):
            types.add(config_file.stem)

    return sorted(types)


def ensure_state_dir() -> Path:
    """Ensure state directory exists and return it."""
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    return STATE_DIR
