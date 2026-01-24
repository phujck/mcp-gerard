"""Standalone LLM functions for direct use (no MCP required).

Reuses the exact same validation and adapter-calling pattern as the MCP chat tool
to ensure consistent behavior across all providers.
"""

import json
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from mcp_handley_lab.llm import memory
from mcp_handley_lab.llm.memory import normalize_branch_input
from mcp_handley_lab.llm.registry import get_adapter, resolve_model, validate_options
from mcp_handley_lab.shared.models import LLMResult, UsageStats


@dataclass
class QueryResult:
    """Result from query() with text, usage info, and raw response."""

    text: str
    input_tokens: int
    output_tokens: int
    model_used: str
    raw: dict[str, Any]


def conversation(branch: str) -> tuple[list[dict[str, str]], str | None]:
    """Get conversation messages and system prompt.

    Args:
        branch: Branch name to read

    Returns:
        Tuple of (messages_list, system_prompt). Messages are dicts with
        'role' and 'content' keys. System prompt may be None.

    Example:
        >>> from mcp_handley_lab import llm
        >>> messages, sys = llm.conversation("main")
        >>> for msg in messages[-3:]:
        ...     print(f"{msg['role']}: {msg['content'][:50]}")
    """
    project_dir = memory.get_project_dir()
    return memory.get_llm_context(project_dir, branch)


def chat(
    prompt: str,
    branch: str = "session",
    model: str = "gemini",
    system_prompt: str = "",
    temperature: float = 1.0,
    options: dict[str, Any] | None = None,
) -> LLMResult:
    """Chat with conversation memory. Works in REPL, scripts, anywhere.

    Like query() but persists conversation to a branch.

    Args:
        prompt: The message to send
        branch: Conversation branch name (default "session")
        model: Model name or provider (gemini, openai, claude, etc.)
        system_prompt: System instructions (only used for new branches)
        temperature: Creativity (0.0-2.0)
        options: Provider-specific options dict

    Returns:
        LLMResult with .content, .usage, .branch, .commit_sha

    Example:
        >>> from mcp_handley_lab import llm
        >>> result = llm.chat("What is 2+2?", branch="math")
        >>> print(result.content)
        >>> result = llm.chat("Double that", branch="math")  # continues conversation
    """
    project_dir = memory.get_project_dir()
    actual_branch = normalize_branch_input(branch)

    # Get existing history or start fresh
    if memory.branch_exists(project_dir, actual_branch):
        history, existing_sys = memory.get_llm_context(project_dir, actual_branch)
        sys_instruction = existing_sys or system_prompt or None
    else:
        history = []
        sys_instruction = system_prompt or None

    # Resolve model and call adapter
    provider, model_id, config = resolve_model(model)
    options = options or {}
    validate_options(provider, model, config, options)

    adapter_type = "deep_research" if config.get("is_agent") else "generation"
    adapter = get_adapter(provider, adapter_type)

    response = adapter(
        prompt=prompt,
        model=model_id,
        history=history,
        system_instruction=sys_instruction,
        temperature=temperature,
        options=options,
    )

    # Build JSONL content - append new turn to existing content
    now = datetime.now().isoformat()
    user_line = json.dumps(
        {"v": 1, "type": "message", "timestamp": now, "role": "user", "content": prompt}
    )
    asst_line = json.dumps(
        {
            "v": 1,
            "type": "message",
            "timestamp": now,
            "role": "assistant",
            "content": response["text"],
            "usage": {
                "provider": provider,
                "model": model_id,
                **{
                    k: response.get(k)
                    for k in ["input_tokens", "output_tokens", "cost"]
                    if response.get(k) is not None
                },
            },
        }
    )
    new_turn = user_line + "\n" + asst_line

    # Append to existing content
    existing = (
        memory.read_branch(project_dir, actual_branch)
        if memory.branch_exists(project_dir, actual_branch)
        else ""
    )
    content = existing + "\n" + new_turn if existing else new_turn

    # Write to branch
    write_result = memory.write_conversation(
        project_dir, actual_branch, content, f"chat: {prompt[:50]}"
    )

    return LLMResult(
        content=response["text"],
        usage=UsageStats(
            input_tokens=response.get("input_tokens", 0),
            output_tokens=response.get("output_tokens", 0),
            cost=response.get("cost", 0.0),
            model_used=model_id,
        ),
        branch=write_result["branch"],
        commit_sha=write_result.get("commit_sha"),
    )


def query(
    prompt: str,
    model: str = "gemini",
    system_prompt: str = "",
    temperature: float = 1.0,
    options: dict[str, Any] | None = None,
) -> QueryResult:
    """Query an LLM directly. Works in REPL, scripts, anywhere.

    Uses the same validation and adapter infrastructure as the MCP chat tool.

    Args:
        prompt: The question/task for the LLM
        model: Model name or provider (gemini, openai, claude, etc.)
        system_prompt: System instructions (empty string = none)
        temperature: Creativity (0.0-2.0)
        options: Provider-specific options dict (grounding, reasoning_effort, etc.)

    Returns:
        QueryResult with .text, .input_tokens, .output_tokens, .model_used, .raw

    Example:
        >>> from mcp_handley_lab import llm
        >>> result = llm.query("What is 2+2?")
        >>> print(result.text)
        'The answer is 4.'
        >>> llm.query("Search for X", model="gemini", options={"grounding": True})
    """
    provider, model_id, config = resolve_model(model)
    options = options or {}
    validate_options(provider, model, config, options)

    adapter_type = "deep_research" if config.get("is_agent") else "generation"
    adapter = get_adapter(provider, adapter_type)

    response = adapter(
        prompt=prompt,
        model=model_id,
        history=[],
        system_instruction=system_prompt or None,
        temperature=temperature,
        options=options,
    )

    return QueryResult(
        text=response["text"],
        input_tokens=response.get("input_tokens", 0),
        output_tokens=response.get("output_tokens", 0),
        model_used=model_id,
        raw=response,
    )
