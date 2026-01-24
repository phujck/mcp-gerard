"""Standalone LLM functions for direct use (no MCP required)."""

import json
from datetime import datetime
from typing import Any

from mcp_handley_lab.llm import memory
from mcp_handley_lab.llm.memory import normalize_branch_input
from mcp_handley_lab.llm.registry import get_adapter, resolve_model, validate_options
from mcp_handley_lab.shared.models import LLMResult, UsageStats


def conversation(branch: str) -> tuple[list[dict[str, str]], str | None]:
    """Get conversation messages and system prompt."""
    return memory.get_llm_context(memory.get_project_dir(), branch)


def chat(
    prompt: str,
    branch: str = "session",
    model: str = "gemini",
    system_prompt: str = "",
    temperature: float = 1.0,
    options: dict[str, Any] | None = None,
) -> LLMResult:
    """Chat with conversation memory."""
    project_dir = memory.get_project_dir()
    actual_branch = normalize_branch_input(branch)
    options = options or {}

    # Read existing content once
    branch_exists = memory.branch_exists(project_dir, actual_branch)
    existing_content = (
        memory.read_branch(project_dir, actual_branch) if branch_exists else ""
    )
    history, existing_sys = (
        memory.get_llm_context(project_dir, actual_branch)
        if branch_exists
        else ([], None)
    )
    sys_instruction = existing_sys or system_prompt or None

    # Call LLM
    provider, model_id, config = resolve_model(model)
    validate_options(provider, model, config, options)
    adapter = get_adapter(
        provider, "deep_research" if config.get("is_agent") else "generation"
    )

    response = adapter(
        prompt=prompt,
        model=model_id,
        history=history,
        system_instruction=sys_instruction,
        temperature=temperature,
        options=options,
    )

    # Build and write JSONL
    now = datetime.now().isoformat()
    lines = [
        existing_content,
        json.dumps(
            {
                "v": 1,
                "type": "message",
                "timestamp": now,
                "role": "user",
                "content": prompt,
            }
        ),
        json.dumps(
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
                        k: response[k]
                        for k in ["input_tokens", "output_tokens", "cost"]
                        if response.get(k)
                    },
                },
            }
        ),
    ]
    content = "\n".join(line for line in lines if line)

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
