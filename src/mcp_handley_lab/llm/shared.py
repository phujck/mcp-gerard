"""Shared utilities for LLM providers."""

import tempfile
import uuid
from collections.abc import Callable
from pathlib import Path

from mcp_handley_lab.common.pricing import calculate_cost
from mcp_handley_lab.llm import memory
from mcp_handley_lab.llm.common import (
    get_session_id,
    load_prompt_text,
)
from mcp_handley_lab.shared.models import (
    GroundingMetadata,
    ImageGenerationResult,
    LLMResult,
)


def should_use_memory(branch: str) -> bool:
    """Determines if memory should be used based on the branch parameter.

    Returns False when branch is empty or "false" (case-insensitive).
    Returns True for any other non-empty string.

    Note: This is a simple check. For full validation with whitespace handling,
    use memory.normalize_branch_input() which raises ValueError on invalid input.
    """
    return bool(branch) and branch.lower() != "false"


def normalize_branch(branch: str) -> str | None:
    """Normalize branch input, returning None if memory should be disabled.

    Args:
        branch: Branch name to normalize

    Returns:
        Normalized branch name, or None if memory should be disabled

    Raises:
        ValueError: If branch name is whitespace-only or invalid
    """
    return memory.normalize_branch_input(branch)


def _handle_memory_setup(
    branch: str,
    system_prompt: str | None,
    mcp_instance,
    provider: str,
    from_ref: str | None = None,
) -> tuple[bool, str, list, str | None, Path | None, str | None]:
    """Set up memory for the LLM request.

    Returns:
        (use_memory, actual_branch, history, system_instruction, project_dir, pending_system_prompt)

    The pending_system_prompt is returned when:
    - Branch is new and system_prompt was provided
    - This should be included in the first commit by _save_conversation_turn()
    """
    # Normalize branch - returns None if memory should be disabled
    normalized = normalize_branch(branch) if branch else None

    use_memory = normalized is not None
    actual_branch = branch
    history = []
    system_instruction = None
    project_dir = None
    pending_system_prompt = None

    if use_memory:
        actual_branch = normalized  # Use normalized branch name
        if actual_branch == "session":
            actual_branch = get_session_id(mcp_instance, provider)

        project_dir = memory.get_project_dir()

        # Check if editing is in progress
        lock_info = memory.is_locked(project_dir)
        if lock_info is not None:
            raise ValueError(
                f"Editing in progress (pid={lock_info.get('pid')}). "
                "Use conversation(action='done') to finish editing before sending messages."
            )

        # Handle from_ref for forking
        if from_ref and not memory.branch_exists(project_dir, actual_branch):
            memory.fork_branch(project_dir, actual_branch, from_ref)

        branch_exists = memory.branch_exists(project_dir, actual_branch)

        if not branch_exists:
            # New branch - don't create yet, let _save_conversation_turn() do it
            # Pass system_prompt to be included in first commit
            pending_system_prompt = system_prompt
            system_instruction = system_prompt
        else:
            # Existing branch - handle system prompt changes
            if system_prompt is not None:
                content = memory.read_branch(project_dir, actual_branch)
                events = memory.parse_messages(content)

                # Find current system prompt (after last clear)
                last_clear_idx = -1
                for i, event in enumerate(events):
                    if event.get("type") == "clear":
                        last_clear_idx = i

                current_system_prompt = None
                for i, event in enumerate(events):
                    if i > last_clear_idx and event.get("type") == "system_prompt":
                        current_system_prompt = event.get("content")

                if system_prompt != current_system_prompt:
                    content = memory.append_system_prompt(content, system_prompt)
                    memory.write_conversation(
                        project_dir, actual_branch, content, "Update system prompt"
                    )

            # Get conversation context
            history, system_instruction = memory.get_llm_context(
                project_dir, actual_branch
            )

    return (
        use_memory,
        actual_branch,
        history,
        system_instruction,
        project_dir,
        pending_system_prompt,
    )


def _extract_response_metadata(response_data: dict, model: str, provider: str) -> dict:
    """Extract metadata from provider response."""
    input_tokens = response_data["input_tokens"]
    output_tokens = response_data["output_tokens"]

    return {
        "response_text": response_data["text"],
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "cost": calculate_cost(model, input_tokens, output_tokens, provider),
        "finish_reason": response_data.get("finish_reason", ""),
        "avg_logprobs": response_data.get("avg_logprobs") or 0.0,
        "model_version": response_data.get("model_version", ""),
        "generation_time_ms": response_data.get("generation_time_ms", 0),
        "response_id": response_data.get("response_id", ""),
        "system_fingerprint": response_data.get("system_fingerprint", ""),
        "service_tier": response_data.get("service_tier", ""),
        "completion_tokens_details": response_data.get("completion_tokens_details", {}),
        "prompt_tokens_details": response_data.get("prompt_tokens_details", {}),
        "stop_sequence": response_data.get("stop_sequence", ""),
        "cache_creation_input_tokens": response_data.get(
            "cache_creation_input_tokens", 0
        ),
        "cache_read_input_tokens": response_data.get("cache_read_input_tokens", 0),
        "grounding_metadata_dict": response_data.get("grounding_metadata"),
        # Enhanced metadata fields (from GPT-5 review)
        "total_tokens": response_data.get("total_tokens", input_tokens + output_tokens),
        "reasoning_text": response_data.get("reasoning_text", ""),
        "created_at": response_data.get("created_at") or 0.0,
        "completed_at": response_data.get("completed_at") or 0.0,
        "timing": response_data.get("timing", {}),
        "token_modalities": response_data.get("token_modalities", {}),
        "cache_creation_details": response_data.get("cache_creation_details", {}),
        "groq_metadata": response_data.get("groq_metadata", {}),
        "citations": response_data.get("citations", []),
        "refusal": response_data.get("refusal") or "",
    }


def _enhance_prompt_for_images(
    prompt: str, user_prompt: str, kwargs: dict
) -> tuple[str, str]:
    """Enhance prompt for image analysis."""
    if "image_data" in kwargs or "images" in kwargs:
        focus = kwargs.get("focus", "general")
        if focus != "general":
            prompt = f"Focus on {focus} aspects. {prompt}"

        image_count = 0
        if kwargs.get("image_data"):
            image_count += 1
        if kwargs.get("images"):
            image_count += len(kwargs.get("images", []))
        if image_count > 0:
            user_prompt = f"{user_prompt} [Image analysis: {image_count} image(s)]"

    return prompt, user_prompt


def _save_conversation_turn(
    project_dir: Path,
    branch: str,
    user_prompt: str,
    response_text: str,
    provider: str,
    model: str,
    metadata: dict | None = None,
    pending_system_prompt: str | None = None,
) -> dict:
    """Save a conversation turn (user + assistant messages) to memory.

    For new branches, creates the branch with the first commit containing
    the optional system_prompt event followed by the conversation turn.

    Returns the write result including commit_sha and forking info.
    """
    # Build usage dict for storage
    usage = None
    if metadata:
        usage = {
            "provider": provider,
            "model": model,
            "input_tokens": metadata.get("input_tokens", 0),
            "output_tokens": metadata.get("output_tokens", 0),
            "cost": metadata.get("cost", 0.0),
        }
        # Include additional metadata fields
        for field in [
            "finish_reason",
            "avg_logprobs",
            "model_version",
            "generation_time_ms",
            "response_id",
            "system_fingerprint",
            "service_tier",
            "completion_tokens_details",
            "prompt_tokens_details",
            "stop_sequence",
            "cache_creation_input_tokens",
            "cache_read_input_tokens",
            "grounding_metadata",
        ]:
            if metadata.get(field):
                usage[field] = metadata[field]

    # Read current content (empty string for new branches)
    content = memory.read_branch(project_dir, branch)

    # For new branches with a system prompt, prepend the system_prompt event
    if not content and pending_system_prompt:
        content = memory.append_system_prompt("", pending_system_prompt)

    # Add user message
    content = memory.append_message(content, "user", user_prompt)

    # Add assistant message with usage
    content = memory.append_message(content, "assistant", response_text, usage=usage)

    # Write back with user message preview as commit message
    commit_message = user_prompt[:50] + "..." if len(user_prompt) > 50 else user_prompt
    # Replace newlines with spaces for cleaner git log
    commit_message = commit_message.replace("\n", " ").strip()
    return memory.write_conversation(project_dir, branch, content, commit_message)


def process_llm_request(
    prompt: str | None,
    output_file: str,
    branch: str,
    model: str,
    provider: str,
    generation_func: Callable,
    mcp_instance,
    from_ref: str | None = None,
    **kwargs,
) -> LLMResult:
    """Generic handler for LLM requests that abstracts common patterns.

    Args:
        prompt: The prompt text
        output_file: File path to save response
        branch: Conversation branch name (replaces agent_name)
        model: Model identifier
        provider: Provider name
        generation_func: Provider-specific generation function
        mcp_instance: MCP server instance
        from_ref: Optional ref to fork from when creating new branch
        **kwargs: Additional arguments for the generation function
    """
    # Extract prompt resolution parameters
    prompt_file = kwargs.pop("prompt_file", None)
    prompt_vars = kwargs.pop("prompt_vars", None)
    system_prompt = kwargs.pop("system_prompt", None)
    system_prompt_file = kwargs.pop("system_prompt_file", None)
    system_prompt_vars = kwargs.pop("system_prompt_vars", None)

    # Resolve final prompt and system prompt
    final_prompt = load_prompt_text(prompt, prompt_file, prompt_vars)
    final_system_prompt = None
    if system_prompt or system_prompt_file:
        final_system_prompt = load_prompt_text(
            system_prompt, system_prompt_file, system_prompt_vars
        )

    user_prompt = final_prompt

    # Set up memory and get conversation context
    (
        use_memory,
        actual_branch,
        history,
        system_instruction,
        project_dir,
        pending_system_prompt,
    ) = _handle_memory_setup(
        branch, final_system_prompt, mcp_instance, provider, from_ref
    )

    # Enhance prompt for image analysis
    final_prompt, user_prompt = _enhance_prompt_for_images(
        final_prompt, user_prompt, kwargs
    )

    # Call provider-specific generation function
    response_data = generation_func(
        prompt=final_prompt,
        model=model,
        history=history,
        system_instruction=system_instruction,
        **kwargs,
    )

    # Extract response metadata
    metadata = _extract_response_metadata(response_data, model, provider)

    # Handle memory with full response metadata
    commit_sha = None
    if use_memory and project_dir:
        write_result = _save_conversation_turn(
            project_dir,
            actual_branch,
            user_prompt,
            metadata["response_text"],
            provider=provider,
            model=model,
            metadata=metadata,
            pending_system_prompt=pending_system_prompt,
        )
        commit_sha = write_result.get("commit_sha")

    # Handle output - write to file if path provided
    if output_file:
        output_path = Path(output_file).expanduser()
        output_path.write_text(metadata["response_text"])

    from mcp_handley_lab.shared.models import UsageStats

    usage_stats = UsageStats(
        input_tokens=metadata["input_tokens"],
        output_tokens=metadata["output_tokens"],
        cost=metadata["cost"],
        model_used=model,
    )

    grounding_metadata = None
    if metadata["grounding_metadata_dict"]:
        grounding_metadata = GroundingMetadata(**metadata["grounding_metadata_dict"])

    return LLMResult(
        content=metadata["response_text"],
        usage=usage_stats,
        branch=actual_branch if use_memory else "",
        commit_sha=commit_sha,
        grounding_metadata=grounding_metadata,
        finish_reason=metadata["finish_reason"],
        avg_logprobs=metadata["avg_logprobs"],
        model_version=metadata["model_version"],
        generation_time_ms=metadata["generation_time_ms"],
        response_id=metadata["response_id"],
        system_fingerprint=metadata["system_fingerprint"],
        service_tier=metadata["service_tier"],
        completion_tokens_details=metadata["completion_tokens_details"],
        prompt_tokens_details=metadata["prompt_tokens_details"],
        stop_sequence=metadata["stop_sequence"],
        cache_creation_input_tokens=metadata["cache_creation_input_tokens"],
        cache_read_input_tokens=metadata["cache_read_input_tokens"],
        # Enhanced metadata fields
        total_tokens=metadata["total_tokens"],
        reasoning_text=metadata["reasoning_text"],
        created_at=metadata["created_at"],
        completed_at=metadata["completed_at"],
        timing=metadata["timing"],
        token_modalities=metadata["token_modalities"],
        cache_creation_details=metadata["cache_creation_details"],
        groq_metadata=metadata["groq_metadata"],
        citations=metadata["citations"],
        refusal=metadata["refusal"],
    )


def process_image_generation(
    prompt: str,
    branch: str,
    model: str,
    provider: str,
    generation_func: Callable,
    mcp_instance,
    **kwargs,
) -> ImageGenerationResult:
    """Generic handler for LLM image generation requests."""
    if not prompt.strip():
        raise ValueError("Prompt is required and cannot be empty")

    # Normalize branch - returns None if memory should be disabled
    normalized = normalize_branch(branch) if branch else None
    use_memory = normalized is not None
    actual_branch = branch

    # Call the provider-specific generation function to get the image
    response_data = generation_func(prompt=prompt, model=model, **kwargs)
    image_bytes = response_data["image_bytes"]
    input_tokens = response_data.get("input_tokens", 0)
    output_tokens = response_data.get("output_tokens", 1)

    file_id = str(uuid.uuid4())[:8]
    filename = f"{provider}_generated_{file_id}.png"
    filepath = Path(tempfile.gettempdir()) / filename
    filepath.write_bytes(image_bytes)

    cost = calculate_cost(
        model, input_tokens, output_tokens, provider, images_generated=1
    )

    # Only handle memory if enabled
    commit_sha = None
    if use_memory:
        actual_branch = normalized  # Use normalized branch name
        if actual_branch == "session":
            actual_branch = get_session_id(mcp_instance, provider)

        project_dir = memory.get_project_dir()

        # Check if editing is in progress
        lock_info = memory.is_locked(project_dir)
        if lock_info is not None:
            raise ValueError(
                f"Editing in progress (pid={lock_info.get('pid')}). "
                "Use conversation(action='done') to finish editing."
            )

        # Build metadata for image generation
        image_metadata = {
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "cost": cost,
        }

        write_result = _save_conversation_turn(
            project_dir,
            actual_branch,
            f"Generate image: {prompt}",
            f"Generated image saved to {filepath}",
            provider=provider,
            model=model,
            metadata=image_metadata,
        )
        commit_sha = write_result.get("commit_sha")

    file_size = len(image_bytes)

    from mcp_handley_lab.shared.models import UsageStats

    usage_stats = UsageStats(
        input_tokens=input_tokens,
        output_tokens=output_tokens,
        cost=cost,
        model_used=model,
    )

    return ImageGenerationResult(
        message="Image Generated Successfully",
        file_path=str(filepath),
        file_size_bytes=file_size,
        usage=usage_stats,
        branch=actual_branch if use_memory else "",
        commit_sha=commit_sha,
        # Metadata from provider response
        generation_timestamp=response_data.get("generation_timestamp", 0),
        enhanced_prompt=response_data.get("enhanced_prompt", ""),
        original_prompt=response_data.get("original_prompt", prompt),
        # Request parameters
        requested_size=response_data.get("requested_size", ""),
        requested_quality=response_data.get("requested_quality", ""),
        requested_format=response_data.get("requested_format", ""),
        aspect_ratio=response_data.get("aspect_ratio", ""),
        # Safety and filtering
        safety_attributes=response_data.get("safety_attributes", {}),
        content_filter_reason=response_data.get("content_filter_reason", ""),
        # Provider-specific metadata
        openai_metadata=response_data.get("openai_metadata", {}),
        gemini_metadata=response_data.get("gemini_metadata", {}),
        # Technical details
        mime_type=response_data.get("mime_type", ""),
        cloud_uri=response_data.get("cloud_uri", ""),
        original_url=response_data.get("original_url", ""),
    )
