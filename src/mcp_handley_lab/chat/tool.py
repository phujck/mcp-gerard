"""Unified Chat Tool for AI interactions via MCP.

Provides a single entry point for multiple LLM providers (Gemini, OpenAI, Claude,
Mistral, Grok, Groq) with model-based provider inference.
"""

from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.llm.registry import (
    get_adapter,
    list_all_models,
    resolve_model,
    validate_options,
)
from mcp_handley_lab.llm.shared import process_image_generation, process_llm_request
from mcp_handley_lab.shared.models import ImageGenerationResult, LLMResult

mcp = FastMCP("Chat Tool")


@mcp.tool(
    description="Send a message to an LLM. Provider is auto-detected from model name. "
    "Supports all major providers: Gemini (gemini-*), OpenAI (gpt-*, o3, o4-*), "
    "Claude (claude-*, sonnet, opus, haiku), Mistral (mistral-*, codestral-*, pixtral-*), "
    "Grok (grok-*), Groq (llama-*, mixtral-*). "
    "Use options dict for provider-specific features like grounding (Gemini), "
    "reasoning_effort (OpenAI), or enable_thinking (Claude)."
)
def ask(
    prompt: str = Field(
        default=None,
        description="The message to send to the LLM.",
    ),
    prompt_file: str = Field(
        default=None,
        description="Path to a file containing the prompt. Cannot be used with 'prompt'.",
    ),
    prompt_vars: dict[str, str] = Field(
        default_factory=dict,
        description="Variables for template substitution using ${var} syntax.",
    ),
    output_file: str = Field(
        ...,
        description="File path to save the response. Use '-' for stdout only.",
    ),
    agent_name: str = Field(
        default="session",
        description="Conversation thread name. Use 'session' for temporary, custom name for persistent, 'false' to disable.",
    ),
    model: str = Field(
        default="gemini-2.5-flash",
        description="Model to use. Provider is inferred from model name.",
    ),
    temperature: float = Field(
        default=1.0,
        description="Controls randomness (0.0-2.0). Higher is more creative.",
    ),
    files: list[str] = Field(
        default_factory=list,
        description="File paths to include as context.",
    ),
    system_prompt: str = Field(
        default=None,
        description="System instructions for the conversation.",
    ),
    system_prompt_file: str = Field(
        default=None,
        description="Path to a file containing system instructions.",
    ),
    system_prompt_vars: dict[str, str] = Field(
        default_factory=dict,
        description="Variables for system prompt template substitution.",
    ),
    options: dict[str, Any] = Field(
        default_factory=dict,
        description="Provider-specific options. Use capabilities(model) to discover. "
        "Examples: grounding (Gemini), reasoning_effort (OpenAI), enable_thinking (Claude).",
    ),
) -> LLMResult:
    """Send a message to an LLM with automatic provider detection."""
    provider, canonical_model, model_config = resolve_model(model)
    validate_options(provider, model, model_config, options)

    generation_func = get_adapter(provider, "generation")

    return process_llm_request(
        prompt=prompt,
        prompt_file=prompt_file,
        prompt_vars=prompt_vars,
        output_file=output_file,
        agent_name=agent_name,
        model=canonical_model,
        provider=provider,
        generation_func=generation_func,
        mcp_instance=mcp,
        temperature=temperature,
        files=files,
        system_prompt=system_prompt,
        system_prompt_file=system_prompt_file,
        system_prompt_vars=system_prompt_vars,
        options=options,
    )


@mcp.tool(
    description="Analyze images with vision-capable LLMs. Provider auto-detected from model. "
    "Supports: Gemini (gemini-*), OpenAI (gpt-4o), Claude (claude-*), "
    "Mistral (pixtral-*), Grok (grok-*-vision)."
)
def analyze_image(
    prompt: str = Field(
        ...,
        description="Question about the image(s).",
    ),
    files: list[str] = Field(
        ...,
        description="Image file paths or base64 strings to analyze.",
    ),
    output_file: str = Field(
        ...,
        description="File path to save analysis. Use '-' for stdout only.",
    ),
    model: str = Field(
        default="gemini-2.5-flash",
        description="Vision model to use. Provider is inferred.",
    ),
    focus: str = Field(
        default="general",
        description="Analysis focus (e.g., 'ocr', 'objects', 'general').",
    ),
    agent_name: str = Field(
        default="session",
        description="Conversation thread name.",
    ),
    system_prompt: str = Field(
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
        agent_name=agent_name,
        model=canonical_model,
        provider=provider,
        generation_func=analysis_func,
        mcp_instance=mcp,
        images=files,
        focus=focus,
        system_prompt=system_prompt,
        options=options,
    )


@mcp.tool(
    description="Generate images from text prompts. "
    "Supports: Gemini (imagen-*), OpenAI (dall-e-*), Grok (grok-*-image)."
)
def generate_image(
    prompt: str = Field(
        ...,
        description="Detailed description of the image to generate.",
    ),
    model: str = Field(
        default="imagen-3.0-generate-002",
        description="Image generation model. Provider is inferred.",
    ),
    agent_name: str = Field(
        default="session",
        description="Conversation thread for prompt history.",
    ),
    size: str = Field(
        default="1024x1024",
        description="Image dimensions (OpenAI DALL-E only).",
    ),
    quality: str = Field(
        default="standard",
        description="Image quality: 'standard' or 'hd' (OpenAI DALL-E 3 only).",
    ),
    aspect_ratio: str = Field(
        default="1:1",
        description="Aspect ratio (Gemini Imagen only).",
    ),
) -> ImageGenerationResult:
    """Generate images from text descriptions."""
    provider, canonical_model, _ = resolve_model(model)

    generation_func = get_adapter(provider, "image_generation")

    return process_image_generation(
        prompt=prompt,
        agent_name=agent_name,
        model=canonical_model,
        provider=provider,
        generation_func=generation_func,
        mcp_instance=mcp,
        size=size,
        quality=quality,
        aspect_ratio=aspect_ratio,
    )


@mcp.tool(
    description="Fill-in-the-middle code completion using Mistral Codestral. "
    "Provide code before and after cursor to get intelligent completions."
)
def fill_in_middle(
    prefix: str = Field(
        ...,
        description="Code before the cursor position.",
    ),
    output_file: str = Field(
        ...,
        description="File path to save the completion result.",
    ),
    suffix: str = Field(
        default="",
        description="Code after the cursor position.",
    ),
    model: str = Field(
        default="codestral-latest",
        description="Model: 'codestral-latest' or 'devstral-small-latest'.",
    ),
    max_tokens: int = Field(
        default=256,
        description="Maximum tokens to generate.",
    ),
    temperature: float = Field(
        default=0.0,
        description="Temperature for sampling. Use 0 for deterministic.",
    ),
    stop: list[str] = Field(
        default_factory=list,
        description="Stop sequences to end generation.",
    ),
) -> dict[str, Any]:
    """Fill-in-the-middle code completion."""
    provider, canonical_model, _ = resolve_model(model)

    if provider != "mistral":
        raise ValueError(
            f"fill_in_middle only supports Mistral models. Got: {model} ({provider})"
        )

    from mcp_handley_lab.llm.mistral import adapter

    result = adapter.fill_in_middle_adapter(
        prefix=prefix,
        suffix=suffix,
        model=canonical_model,
        max_tokens=max_tokens,
        temperature=temperature,
        stop=stop if stop else None,
    )

    # Write to output file
    Path(output_file).write_text(result["full_code"])

    return result


@mcp.tool(
    description="Analyze text for harmful content using Mistral's moderation model. "
    "Detects violence, hate speech, sexual content, and self-harm."
)
def moderate_content(
    text: str = Field(
        ...,
        description="Text to analyze for harmful content.",
    ),
) -> dict[str, Any]:
    """Moderate text content for safety."""
    from mcp_handley_lab.llm.mistral import adapter

    return adapter.moderation_adapter(text)


@mcp.tool(
    description="List all available models from all providers with full details including "
    "capabilities, supported options, and constraints."
)
def list_models() -> dict[str, list[dict[str, Any]]]:
    """List all available models grouped by provider with capabilities."""
    return list_all_models()
