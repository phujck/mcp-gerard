"""Shared MCP instance for unified email tool with dynamic description injection."""

import asyncio
import logging
from collections.abc import Callable
from contextlib import asynccontextmanager
from typing import Any, TypedDict

from mcp.server.fastmcp import FastMCP

logger = logging.getLogger(__name__)


class ToolConfig(TypedDict):
    fn: Callable[..., Any]
    description: str


# Tool configs for lifespan-based description injection.
# Populated by provider modules (notmuch, mutt, offlineimap) before mcp.run().
_TOOL_CONFIGS: dict[str, ToolConfig] = {}


def _fetch_email_context() -> tuple[list[str], list[str]]:
    """Fetch tags and folders for injection into tool descriptions."""
    from mcp_handley_lab.email.notmuch.tool import _list_folders, _list_tags

    return _list_tags(), _list_folders()


@asynccontextmanager
async def _lifespan(app: FastMCP):
    """Inject tags/folders into tool descriptions at startup."""
    try:
        tags, folders = await asyncio.wait_for(
            asyncio.to_thread(_fetch_email_context),
            timeout=5.0,
        )
    except Exception:
        logger.warning("Failed to fetch email context for injection", exc_info=True)
        yield
        return

    # Format for injection (cap tags at 50 to avoid huge descriptions)
    tags_text = "\n".join(f"- {t}" for t in sorted(tags)[:50])
    if len(tags) > 50:
        tags_text += f"\n... and {len(tags) - 50} more tags"
    folders_text = "\n".join(f"- {f}" for f in sorted(folders))

    # Tool-specific injection content
    # read: tags only (for query building)
    # update: tags + folders (for move/archive operations)
    injection_map = {
        "read": f"\n\nAvailable tags:\n{tags_text}",
        "update": f"\n\nAvailable tags:\n{tags_text}\n\nAvailable folders:\n{folders_text}",
    }

    # Re-register tools with injected descriptions
    for tool_name, injection_text in injection_map.items():
        if tool_name not in _TOOL_CONFIGS:
            continue
        config = _TOOL_CONFIGS[tool_name]
        try:
            app.remove_tool(tool_name)
            app.add_tool(
                config["fn"],
                name=tool_name,
                description=config["description"] + injection_text,
            )
        except Exception:
            logger.warning(f"Failed to inject into {tool_name}", exc_info=True)

    yield


# Single, shared MCP instance for the entire email tool.
# All provider modules will import and use this instance.
mcp = FastMCP("Email", lifespan=_lifespan)
