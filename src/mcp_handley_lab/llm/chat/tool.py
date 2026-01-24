"""Unified Chat Tool for AI interactions via MCP.

Provides a single entry point for multiple LLM providers (Gemini, OpenAI, Claude,
Mistral, Grok, Groq) with model-based provider inference and Git-backed memory.
"""

from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.llm import memory
from mcp_handley_lab.llm.registry import (
    get_adapter,
    resolve_model,
    validate_options,
)
from mcp_handley_lab.llm.shared import process_llm_request
from mcp_handley_lab.shared.models import LLMResult

mcp = FastMCP("Chat Tool")


@mcp.tool(
    description="Send a message to an LLM. Provider is auto-detected from model name. "
    "Supports Gemini, OpenAI, Claude, Mistral, Grok, and Groq. "
    "Use provider names for latest defaults (e.g., model='gemini'). "
    "Returns: {content, usage: {input_tokens, output_tokens, cost, model_used}, agent_name}."
)
def chat(
    prompt: str | None = Field(
        default=None,
        description="The message to send to the LLM.",
    ),
    prompt_file: str | None = Field(
        default=None,
        description="Path to a file containing the prompt. Cannot be used with 'prompt'.",
    ),
    prompt_vars: dict[str, str] = Field(
        default_factory=dict,
        description="Variables for template substitution using ${var} syntax.",
    ),
    output_file: str = Field(
        default="",
        description="File path to save the response. Empty string means no file output. "
        "Responses are always stored in memory (~/.mcp-handley-lab/) and can be "
        "retrieved via get_response().",
    ),
    branch: str = Field(
        default="session",
        description="Conversation branch name. 'session' uses a shared auto-generated ID "
        "(WARNING: collides across sub-agents). Use unique names for isolated conversations, "
        "'false' to disable memory.",
    ),
    model: str = Field(
        default="gemini",
        description="Model or provider name. Provider is inferred automatically. "
        "Use provider names (gemini, openai, claude) for latest defaults, "
        "or specific model IDs. Run list_models() to see available options.",
    ),
    temperature: float = Field(
        default=1.0,
        description="Controls randomness (0.0-2.0). Higher is more creative.",
    ),
    files: list[str] = Field(
        default_factory=list,
        description="Files to include as context. Accepts: local paths, URLs, "
        "or data URIs (data:mime/type;base64,...). Text files read as content, "
        "images/PDFs sent to multimodal models.",
    ),
    system_prompt: str | None = Field(
        default=None,
        description="System instructions for the conversation.",
    ),
    system_prompt_file: str | None = Field(
        default=None,
        description="Path to a file containing system instructions.",
    ),
    system_prompt_vars: dict[str, str] = Field(
        default_factory=dict,
        description="Variables for system prompt template substitution.",
    ),
    options: dict[str, Any] = Field(
        default_factory=dict,
        description="Provider-specific options. Use mcp-models list_models() to discover. "
        "Examples: grounding (Gemini), reasoning_effort (OpenAI), enable_thinking (Claude).",
    ),
    from_ref: str | None = Field(
        default=None,
        description="Fork from this ref when creating a new conversation branch. "
        "Use commit_sha from a previous response to fork from that point.",
    ),
) -> LLMResult:
    """Send a message to an LLM with automatic provider detection."""
    provider, canonical_model, model_config = resolve_model(model)
    validate_options(provider, model, model_config, options)

    # Route agent models (e.g., deep research) to their specialized adapter
    adapter_type = "deep_research" if model_config.get("is_agent") else "generation"
    generation_func = get_adapter(provider, adapter_type)

    return process_llm_request(
        prompt=prompt,
        prompt_file=prompt_file,
        prompt_vars=prompt_vars,
        output_file=output_file,
        branch=branch,
        model=canonical_model,
        provider=provider,
        generation_func=generation_func,
        mcp_instance=mcp,
        from_ref=from_ref,
        temperature=temperature,
        files=files,
        system_prompt=system_prompt,
        system_prompt_file=system_prompt_file,
        system_prompt_vars=system_prompt_vars,
        options=options,
    )


@mcp.tool(
    description="Analyze images with vision-capable LLMs. Provider auto-detected from model. "
    "Supports Gemini, OpenAI, Claude, Mistral, and Grok vision models. "
    "Returns: {content, usage: {input_tokens, output_tokens, cost, model_used}, agent_name}."
)
def analyze_image(
    prompt: str = Field(
        ...,
        description="Question about the image(s).",
    ),
    images: list[str] = Field(
        ...,
        description="Images to analyze. Accepts: local paths, URLs, "
        "data URIs (data:image/png;base64,...), or raw base64 strings.",
    ),
    output_file: str = Field(
        default="",
        description="File path to save analysis. Empty string means no file output. "
        "Responses stored in memory and retrievable via get_response().",
    ),
    model: str = Field(
        default="gemini",
        description="Vision model or provider name. Provider is inferred automatically. "
        "Use provider names for latest defaults, or specific model IDs.",
    ),
    focus: str = Field(
        default="general",
        description="Analysis focus (e.g., 'ocr', 'objects', 'general').",
    ),
    branch: str = Field(
        default="session",
        description="Conversation branch name. 'session' uses a shared auto-generated ID "
        "(WARNING: collides across sub-agents). Use unique names for isolated conversations.",
    ),
    system_prompt: str | None = Field(
        default=None,
        description="System instructions for the analysis.",
    ),
    options: dict[str, Any] = Field(
        default_factory=dict,
        description="Provider-specific options.",
    ),
) -> LLMResult:
    """Analyze images with vision-capable LLMs."""
    provider, canonical_model, model_config = resolve_model(model)
    validate_options(provider, model, model_config, options)

    analysis_func = get_adapter(provider, "image_analysis")

    return process_llm_request(
        prompt=prompt,
        output_file=output_file,
        branch=branch,
        model=canonical_model,
        provider=provider,
        generation_func=analysis_func,
        mcp_instance=mcp,
        images=images,
        focus=focus,
        system_prompt=system_prompt,
        options=options,
    )


@mcp.tool(
    description="Retrieve a past assistant response from a conversation branch. "
    "Only returns assistant messages (LLM responses), not user messages. "
    "Returns: {content, usage: {input_tokens, output_tokens, cost, model_used}, agent_name, ...metadata}."
)
def get_response(
    branch: str = Field(
        ...,
        description="The conversation branch to retrieve the response from. "
        "Use 'session' with provider param to get current session's responses.",
    ),
    index: int = Field(
        default=-1,
        description="Response index among assistant messages only. "
        "Use -1 for last response, -2 for second-to-last, 0 for first, etc.",
    ),
    provider: str = Field(
        default="",
        description="Provider name (gemini, openai, etc.) to resolve 'session' branch. "
        "Required when branch is 'session'.",
    ),
    output_file: str = Field(
        default="",
        description="File path to save response content. Empty string means no file output.",
    ),
) -> dict[str, Any]:
    """Retrieve an assistant response from a conversation branch.

    Returns the full message dict including content and usage metadata.
    Raises ValueError if branch not found, IndexError if no assistant responses or out of range.
    """
    from pathlib import Path

    from mcp_handley_lab.llm.common import get_session_id

    actual_branch = branch
    if branch == "session":
        if not provider:
            raise ValueError("provider parameter is required when branch is 'session'")
        actual_branch = get_session_id(mcp, provider)

    project_dir = memory.get_project_dir()
    response = memory.get_response(project_dir, actual_branch, index)

    if output_file:
        Path(output_file).expanduser().write_text(response["content"])

    return response


@mcp.tool(
    description="Review conversation histories in JSON format. "
    "Shows user questions in full and abbreviated assistant responses. "
    "Returns: {project, branches: [{name, stats, messages: [{role, content, timestamp, response_index?}]}]}."
)
def review_conversations(
    branch: str = Field(
        default="",
        description="Filter to specific branch. Empty for all branches.",
    ),
    output_file: str = Field(
        default="",
        description="File path to save review. Empty string means no file output.",
    ),
    max_response_chars: int = Field(
        default=200,
        description="Max characters per assistant response (truncated with ...).",
    ),
) -> dict[str, Any]:
    """Review all conversation branches with truncated assistant responses.

    Returns JSON with branch stats, system prompts, and messages.
    Assistant responses show response_index for use with get_response().
    """
    import json
    from pathlib import Path

    project_dir = memory.get_project_dir()
    branches = memory.list_branches(project_dir)

    # Filter to specific branch if requested
    if branch:
        branches = [b for b in branches if b["name"] == branch]
        if not branches:
            raise ValueError(f"Branch '{branch}' not found")

    branch_summaries = []
    for branch_info in branches:
        name = branch_info["name"]
        stats = memory.agent_stats(project_dir, name)

        # Get messages with truncation
        content = memory.read_branch(project_dir, name)
        events = memory.parse_messages(content)

        # Find last clear boundary
        last_clear_idx = -1
        for i, event in enumerate(events):
            if event.get("type") == "clear":
                last_clear_idx = i

        # Build messages list with response indices
        messages = []
        assistant_idx = 0
        for i, event in enumerate(events):
            if i <= last_clear_idx:
                continue
            if event.get("type") != "message":
                continue

            role = event.get("role")
            msg_content = event.get("content", "")
            timestamp = event.get("timestamp", "")

            msg = {
                "role": role,
                "content": msg_content,
                "timestamp": timestamp,
            }

            if role == "assistant":
                # Add response_index for retrieval
                msg["response_index"] = -(stats["message_count"] // 2 - assistant_idx)
                assistant_idx += 1

                # Truncate long assistant responses
                if len(msg_content) > max_response_chars:
                    msg["content"] = msg_content[:max_response_chars] + "..."
                    msg["truncated"] = True
                    msg["full_length"] = len(msg_content)

            messages.append(msg)

        branch_summary = {
            "name": name,
            "stats": {
                "messages": stats["message_count"],
                "tokens": stats["total_tokens"],
                "cost": stats["total_cost"],
            },
            "messages": messages,
        }

        if stats.get("system_prompt"):
            branch_summary["system_prompt"] = stats["system_prompt"]

        branch_summaries.append(branch_summary)

    result = {
        "project": str(project_dir),
        "agents": branch_summaries,  # "agents" key for backward compatibility
    }

    if output_file:
        Path(output_file).expanduser().write_text(json.dumps(result, indent=2))

    return result


@mcp.tool(
    description="Manage conversation branches. Actions: "
    "'list' (all branches), 'log' (history with hashes), 'show' (content at ref), "
    "'edit' (start editing session with worktree), 'done' (end editing session)."
)
def conversation(
    action: str = Field(
        ...,
        description="Action to perform: 'list', 'log', 'show', 'edit', 'done'.",
    ),
    branch: str = Field(
        default="",
        description="Target branch for log/show actions.",
    ),
    ref: str = Field(
        default="",
        description="Specific commit ref for show action. If provided, takes precedence over branch.",
    ),
    limit: int = Field(
        default=20,
        description="For log action: maximum number of entries to return.",
    ),
    force: bool = Field(
        default=False,
        description="For done action: force removal even if lock not held by this process.",
    ),
) -> dict[str, Any]:
    """Git interface for conversation management.

    Actions:
    - list: List all conversation branches with stats (message count, latest timestamp)
    - log: Get commit history for a branch with hashes for progressive disclosure
    - show: Get full JSONL content at branch tip or specific ref
    - edit: Start editing session (creates lock + worktree for Git operations)
    - done: End editing session (removes worktree + lock)
    """
    project_dir = memory.get_project_dir()

    if action == "list":
        branches = memory.list_branches(project_dir)
        return {"branches": branches}

    elif action == "log":
        if not branch:
            raise ValueError("branch parameter is required for 'log' action")
        log_entries = memory.get_log(project_dir, branch, limit)
        return {"branch": branch, "entries": log_entries}

    elif action == "show":
        if not ref and not branch:
            raise ValueError(
                "Either 'ref' or 'branch' must be specified for 'show' action"
            )

        if ref:
            content, resolved_sha = memory.read_ref(project_dir, ref)
            return {"content": content, "ref": resolved_sha}
        else:
            # Check branch exists before reading
            sha = memory.get_branch_sha(project_dir, branch)
            if sha is None:
                raise ValueError(f"Branch '{branch}' not found")
            content = memory.read_branch(project_dir, branch)
            return {"content": content, "ref": sha, "branch": branch}

    elif action == "edit":
        # start_edit returns {"path": str} dict directly
        return memory.start_edit(project_dir)

    elif action == "done":
        memory.end_edit(project_dir, force=force)
        return {"status": "success", "message": "Edit session ended"}

    else:
        raise ValueError(
            f"Unknown action: {action}. Valid actions: list, log, show, edit, done"
        )
