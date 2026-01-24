"""Model Registry Tool for discovering available LLM models.

Provides model discovery across all providers (Gemini, OpenAI, Claude,
Mistral, Grok, Groq) with capabilities, pricing, and supported options.
"""

from typing import Any

from mcp.server.fastmcp import FastMCP

from mcp_handley_lab.llm.registry import list_all_models

mcp = FastMCP("Models Tool")


@mcp.tool(
    description="List all available models from all providers with full details including "
    "capabilities, supported options, pricing, and constraints. Use this to discover "
    "which models to use with mcp-chat, mcp-image, mcp-embeddings, etc."
)
def list_models() -> dict[str, list[dict[str, Any]]]:
    """List all available models grouped by provider with capabilities."""
    return list_all_models()
