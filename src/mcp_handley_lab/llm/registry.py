"""Model registry for unified LLM provider routing.

Provides model resolution (model name → provider) and option validation
for the unified mcp-chat tool.
"""

from pathlib import Path
from typing import Any

import yaml

# All supported providers
PROVIDERS = ["gemini", "openai", "claude", "mistral", "grok", "groq"]

# Prefix fallback rules (longest match wins)
# Order matters: longer prefixes should match before shorter ones
MODEL_PREFIXES = [
    # Gemini
    ("gemini-embedding", "gemini"),
    ("gemini-", "gemini"),
    ("imagen-", "gemini"),
    ("veo-", "gemini"),
    # OpenAI
    ("gpt-image", "openai"),
    ("gpt-", "openai"),
    ("dall-e", "openai"),
    ("text-embedding", "openai"),
    ("o3", "openai"),
    ("o4", "openai"),
    # Claude
    ("claude-", "claude"),
    # Mistral
    ("codestral", "mistral"),
    ("pixtral", "mistral"),
    ("ministral", "mistral"),
    ("mistral-", "mistral"),
    ("mistral-embed", "mistral"),
    # Grok
    ("grok-", "grok"),
    # Groq (often hosts llama/mixtral models)
    ("llama-", "groq"),
    ("mixtral-", "groq"),
]

# Claude model aliases
CLAUDE_ALIASES = {
    "sonnet": "claude-sonnet-4-5-20250929",
    "opus": "claude-opus-4-1-20250805",
    "haiku": "claude-haiku-4-5-20251001",
}

# Provider-specific options and their valid values
PROVIDER_OPTIONS = {
    "gemini": {
        "grounding": {"type": "bool", "description": "Enable Google Search grounding"},
        "thinking_level": {
            "type": "str",
            "values": ["low", "high"],
            "description": "Thinking effort level for Gemini 3 models",
        },
        "thinking_budget": {
            "type": "int",
            "description": "Token budget for thinking (128-32768, or -1 for dynamic)",
        },
        "include_thoughts": {
            "type": "bool",
            "description": "Include model's thinking in output",
        },
    },
    "openai": {
        "reasoning_effort": {
            "type": "str",
            "values": ["none", "low", "medium", "high", "xhigh"],
            "description": "Reasoning effort for GPT-5.x and o-series models",
        },
        "reasoning_summary": {
            "type": "str",
            "values": ["auto", "concise", "detailed"],
            "description": "Reasoning summary format",
        },
        "verbosity": {
            "type": "str",
            "values": ["low", "medium", "high"],
            "description": "Output verbosity (GPT-5.1+ only)",
        },
    },
    "claude": {
        "enable_thinking": {
            "type": "bool",
            "description": "Enable extended thinking mode",
        },
        "thinking_budget": {
            "type": "int",
            "description": "Maximum tokens for thinking (min 1024)",
        },
    },
    "mistral": {
        "include_thinking": {
            "type": "bool",
            "description": "Include reasoning model thinking in output",
        },
    },
    "grok": {},
    "groq": {},
}


def _load_models_yaml(provider: str) -> dict[str, Any]:
    """Load models.yaml for a provider."""
    yaml_path = Path(__file__).parent / provider / "models.yaml"
    if not yaml_path.exists():
        return {}
    with open(yaml_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)
    return config.get("models", {})


def build_model_registry() -> dict[str, tuple[str, dict[str, Any]]]:
    """Build unified model registry from all providers' models.yaml files.

    Returns:
        Dict mapping model_id → (provider, model_config)
    """
    registry: dict[str, tuple[str, dict[str, Any]]] = {}

    for provider in PROVIDERS:
        models = _load_models_yaml(provider)
        for model_id, config in models.items():
            registry[model_id] = (provider, config)

    # Add Claude aliases
    for alias, full_id in CLAUDE_ALIASES.items():
        if full_id in registry:
            registry[alias] = registry[full_id]

    return registry


# Build registry at module load time
MODEL_REGISTRY = build_model_registry()


def get_default_model(provider: str) -> str:
    """Get the default model for a provider."""
    defaults = {
        "gemini": "gemini-2.5-flash",
        "openai": "gpt-5.2",
        "claude": "claude-sonnet-4-5-20250929",
        "mistral": "mistral-large-latest",
        "grok": "grok-4-fast-reasoning",
        "groq": "llama-3.3-70b-versatile",
    }
    return defaults.get(provider, "")


def resolve_model(model: str) -> tuple[str, str, dict[str, Any]]:
    """Resolve a model name to its provider and configuration.

    Args:
        model: Model name (e.g., "gemini-2.5-flash", "sonnet", "gpt-5.2")

    Returns:
        Tuple of (provider, canonical_model_id, model_config)

    Raises:
        ValueError: If model cannot be resolved to any provider
    """
    # 1. Exact match in registry (includes aliases)
    if model in MODEL_REGISTRY:
        provider, config = MODEL_REGISTRY[model]
        # Resolve alias to canonical ID if needed
        canonical_id = CLAUDE_ALIASES.get(model, model)
        return provider, canonical_id, config

    # 2. Prefix fallback (sorted by length, longest first)
    for prefix, provider in sorted(MODEL_PREFIXES, key=lambda x: -len(x[0])):
        if model.startswith(prefix):
            # Unknown model variant - return with empty config
            return provider, model, {}

    # 3. Error with helpful message
    available_providers = ", ".join(PROVIDERS)
    raise ValueError(
        f"Unknown model: '{model}'. Cannot infer provider.\n"
        f"Use list_models() to see available models, or specify a model from: {available_providers}"
    )


def get_supported_options(provider: str, model_config: dict[str, Any]) -> set[str]:
    """Get the set of supported options for a provider/model combination.

    Args:
        provider: Provider name
        model_config: Model configuration from YAML

    Returns:
        Set of supported option names
    """
    # Start with provider-level options
    supported = set(PROVIDER_OPTIONS.get(provider, {}).keys())

    # Model-specific capability flags can restrict options
    # e.g., only models with supports_grounding=True can use grounding
    if provider == "gemini":
        if not model_config.get("supports_grounding", False):
            supported.discard("grounding")
        if not model_config.get("supports_thinking_level", False):
            supported.discard("thinking_level")
            supported.discard("thinking_budget")
            supported.discard("include_thoughts")

    if provider == "openai":
        if not model_config.get("supports_reasoning", False):
            supported.discard("reasoning_effort")
            supported.discard("reasoning_summary")
        if not model_config.get("supports_verbosity", False):
            supported.discard("verbosity")

    if provider == "claude" and not model_config.get(
        "supports_extended_thinking", True
    ):
        supported.discard("enable_thinking")
        supported.discard("thinking_budget")

    if provider == "mistral" and not model_config.get("supports_reasoning", False):
        supported.discard("include_thinking")

    return supported


def validate_options(
    provider: str, model: str, model_config: dict[str, Any], options: dict[str, Any]
) -> None:
    """Validate that options are supported by the provider/model.

    Implements strict validation: raises error if user sets unsupported option.

    Args:
        provider: Provider name
        model: Model ID
        model_config: Model configuration from YAML
        options: User-provided options dict

    Raises:
        ValueError: If an unsupported option is explicitly set
    """
    supported = get_supported_options(provider, model_config)

    for key, value in options.items():
        # Skip if value is None, False, or empty string (not explicitly set)
        if value is None or value is False or value == "":
            continue

        # Skip if value equals the default "none" for reasoning_effort
        if key == "reasoning_effort" and value == "none":
            continue

        if key not in supported:
            supported_list = sorted(supported) if supported else ["none"]
            raise ValueError(
                f"'{key}' is not supported by {provider} models.\n"
                f"Inferred provider: {provider} (from model '{model}')\n"
                f"Supported options for this model: {', '.join(supported_list)}\n"
                f"Use capabilities('{model}') for full details."
            )


def get_model_capabilities(model: str) -> dict[str, Any]:
    """Get capabilities and supported options for a model.

    This is the data source for the capabilities() tool.

    Args:
        model: Model name

    Returns:
        Dict with provider, model info, and supported options
    """
    provider, canonical_id, config = resolve_model(model)
    supported = get_supported_options(provider, config)

    # Build option details
    option_details = {}
    for opt_name in supported:
        opt_info = PROVIDER_OPTIONS.get(provider, {}).get(opt_name, {})
        option_details[opt_name] = {
            "type": opt_info.get("type", "unknown"),
            "description": opt_info.get("description", ""),
        }
        if "values" in opt_info:
            option_details[opt_name]["values"] = opt_info["values"]

    return {
        "model": canonical_id,
        "provider": provider,
        "description": config.get("description", ""),
        "context_window": config.get("context_window", ""),
        "capabilities": {
            "vision": config.get("supports_vision", False),
            "grounding": config.get("supports_grounding", False),
            "reasoning": config.get("supports_reasoning", False),
            "extended_thinking": config.get("supports_extended_thinking", False),
            "image_generation": config.get("pricing_type") == "per_image",
        },
        "supported_options": option_details,
        "constraints": _get_model_constraints(provider, config),
    }


def _get_model_constraints(provider: str, config: dict[str, Any]) -> list[str]:
    """Get usage constraints for a model."""
    constraints = []

    if provider == "openai" and not config.get("supports_temperature", True):
        constraints.append("temperature only supported when reasoning_effort='none'")

    if provider == "claude" and config.get("supports_extended_thinking", False):
        constraints.append("temperature not allowed when enable_thinking=True")

    return constraints


def list_all_models() -> dict[str, list[dict[str, Any]]]:
    """List all available models grouped by provider with full details.

    Returns:
        Dict mapping provider → list of model info dicts with capabilities
    """
    result: dict[str, list[dict[str, Any]]] = {p: [] for p in PROVIDERS}

    for model_id, (provider, config) in MODEL_REGISTRY.items():
        # Skip aliases (they duplicate the canonical entry)
        if model_id in CLAUDE_ALIASES:
            continue

        # Get supported options for this model
        supported = get_supported_options(provider, config)
        option_details = {}
        for opt_name in supported:
            opt_info = PROVIDER_OPTIONS.get(provider, {}).get(opt_name, {})
            option_details[opt_name] = {
                "type": opt_info.get("type", "unknown"),
                "description": opt_info.get("description", ""),
            }
            if "values" in opt_info:
                option_details[opt_name]["values"] = opt_info["values"]

        result[provider].append(
            {
                "id": model_id,
                "description": config.get("description", ""),
                "context_window": config.get("context_window", ""),
                "tags": config.get("tags", []),
                "capabilities": {
                    "vision": config.get("supports_vision", False),
                    "grounding": config.get("supports_grounding", False),
                    "reasoning": config.get("supports_reasoning", False),
                    "extended_thinking": config.get(
                        "supports_extended_thinking", False
                    ),
                    "image_generation": config.get("pricing_type") == "per_image",
                },
                "supported_options": option_details,
                "constraints": _get_model_constraints(provider, config),
            }
        )

    return result


def get_adapter(provider: str, adapter_type: str):
    """Dynamically import and return the appropriate adapter function.

    This is the central routing function for all provider adapters.
    Supported adapter types depend on the provider's capabilities.

    Args:
        provider: Provider name (gemini, openai, claude, mistral, grok, groq)
        adapter_type: Type of adapter (generation, image_analysis, image_generation,
                      fill_in_middle, moderation)

    Returns:
        The adapter function for the specified provider and type

    Raises:
        ValueError: If provider is unknown or doesn't support the adapter type
    """
    if provider == "gemini":
        from mcp_handley_lab.llm.gemini import adapter

        adapters = {
            "generation": adapter.generation_adapter,
            "image_analysis": adapter.image_analysis_adapter,
            "image_generation": adapter.image_generation_adapter,
        }
    elif provider == "openai":
        from mcp_handley_lab.llm.openai import adapter

        adapters = {
            "generation": adapter.generation_adapter,
            "image_analysis": adapter.image_analysis_adapter,
            "image_generation": adapter.image_generation_adapter,
        }
    elif provider == "claude":
        from mcp_handley_lab.llm.claude import adapter

        adapters = {
            "generation": adapter.generation_adapter,
            "image_analysis": adapter.image_analysis_adapter,
        }
    elif provider == "mistral":
        from mcp_handley_lab.llm.mistral import adapter

        adapters = {
            "generation": adapter.generation_adapter,
            "image_analysis": adapter.image_analysis_adapter,
            "fill_in_middle": adapter.fill_in_middle_adapter,
            "moderation": adapter.moderation_adapter,
        }
    elif provider == "grok":
        from mcp_handley_lab.llm.grok import adapter

        adapters = {
            "generation": adapter.generation_adapter,
            "image_analysis": adapter.image_analysis_adapter,
            "image_generation": adapter.image_generation_adapter,
        }
    elif provider == "groq":
        from mcp_handley_lab.llm.groq import adapter

        adapters = {"generation": adapter.generation_adapter}
    else:
        raise ValueError(f"Unknown provider: {provider}")

    if adapter_type not in adapters:
        raise ValueError(f"Provider '{provider}' does not support '{adapter_type}'")

    return adapters[adapter_type]
