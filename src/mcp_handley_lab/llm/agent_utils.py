"""Shared utility functions for agent management - not exposed as MCP tools.

Uses Git-backed memory storage for conversation persistence.
"""

from pathlib import Path

from mcp_handley_lab.llm import memory


def create_agent(
    agent_name: str, system_prompt: str = "", cwd: Path | None = None
) -> str:
    """Create a new agent (conversation branch) with optional system prompt."""
    project_dir = memory.get_project_dir(cwd)
    memory.create_agent(project_dir, agent_name, system_prompt or None)
    system_prompt_info = (
        f" with system prompt: {system_prompt}" if system_prompt else ""
    )
    return f"Agent '{agent_name}' created successfully{system_prompt_info}!"


def list_agents(cwd: Path | None = None) -> str:
    """List all agents (conversation branches) with their statistics."""
    project_dir = memory.get_project_dir(cwd)
    branches = memory.list_branches(project_dir)

    if not branches:
        return "No agents found. Create an agent with create_agent()."

    result = "**Agent List**\n\n"
    for branch_info in branches:
        name = branch_info["name"]
        try:
            stats = memory.agent_stats(project_dir, name)
            result += f"**{stats['name']}**\n"
            created_at = stats.get("created_at")
            if created_at:
                result += f"- Created: {created_at[:10]}\n"
            result += f"- Messages: {stats['message_count']}\n"
            result += f"- Tokens: {stats['total_tokens']:,}\n"
            result += f"- Cost: ${stats['total_cost']:.4f}\n"
            if stats.get("system_prompt"):
                result += f"- System Prompt: {stats['system_prompt']}\n"
            result += "\n"
        except ValueError:
            # Branch exists but might be empty
            result += f"**{name}**\n"
            result += f"- Messages: {branch_info['message_count']}\n\n"

    return result


def agent_stats(agent_name: str, cwd: Path | None = None) -> str:
    """Get detailed statistics for a specific agent."""
    project_dir = memory.get_project_dir(cwd)
    stats = memory.agent_stats(project_dir, agent_name)

    result = f"**Agent Statistics: {agent_name}**\n\n"
    result += "**Overview:**\n"
    created_at = stats.get("created_at")
    if created_at:
        result += f"- Created: {created_at}\n"
    result += f"- Total Messages: {stats['message_count']}\n"
    result += f"- Total Tokens: {stats['total_tokens']:,}\n"
    result += f"- Total Cost: ${stats['total_cost']:.4f}\n"

    if stats.get("system_prompt"):
        result += f"- System Prompt: {stats['system_prompt']}\n"

    # Get recent messages from content
    content = memory.read_branch(project_dir, agent_name)
    events = memory.parse_messages(content)

    # Find last clear boundary
    last_clear_idx = -1
    for i, event in enumerate(events):
        if event.get("type") == "clear":
            last_clear_idx = i

    # Get message events after last clear
    messages = [
        e
        for i, e in enumerate(events)
        if i > last_clear_idx and e.get("type") == "message"
    ]

    if messages:
        result += "\n**Recent Messages:**\n"
        recent_messages = messages[-5:]
        for i, msg in enumerate(recent_messages, 1):
            role = msg.get("role", "unknown")
            content_text = msg.get("content", "")

            # Truncate long messages
            if len(content_text) > 100:
                content_text = content_text[:97] + "..."

            result += f"{i}. **{role.title()}:** {content_text}\n"

    return result


def clear_agent(agent_name: str, cwd: Path | None = None) -> str:
    """Clear an agent's conversation history by appending a clear event."""
    project_dir = memory.get_project_dir(cwd)
    memory.clear_agent(project_dir, agent_name)
    return f"Agent '{agent_name}' history cleared successfully!"


def delete_agent(agent_name: str, cwd: Path | None = None) -> str:
    """Delete an agent (conversation branch) permanently.

    Note: In Git-backed storage, this removes the branch reference.
    The commits remain in the repository until garbage collected.
    """
    project_dir = memory.get_project_dir(cwd)

    if not memory.branch_exists(project_dir, agent_name):
        raise ValueError(f"Agent '{agent_name}' not found")

    # Delete the branch reference
    result = memory._git_unchecked(project_dir, "branch", "-D", agent_name)
    if result.returncode != 0:
        raise ValueError(f"Failed to delete agent '{agent_name}': {result.stderr}")

    return f"Agent '{agent_name}' deleted permanently!"


def get_response(agent_name: str, index: int = -1, cwd: Path | None = None) -> dict:
    """Get a full message from an agent's conversation history by index.

    Returns the full message dict including content and usage metadata.
    """
    project_dir = memory.get_project_dir(cwd)
    return memory.get_response(project_dir, agent_name, index)
