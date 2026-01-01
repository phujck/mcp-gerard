"""Unified email client MCP tool integrating all email providers."""

# Import the shared mcp instance (re-exported for tests)
from mcp_handley_lab.email.common import mcp  # noqa: F401

# Import tool modules to register their @mcp.tool decorators
from mcp_handley_lab.email.msmtp import tool as _msmtp  # noqa: F401
from mcp_handley_lab.email.mutt import tool as _mutt  # noqa: F401
from mcp_handley_lab.email.notmuch import tool as _notmuch  # noqa: F401
from mcp_handley_lab.email.offlineimap import tool as _offlineimap  # noqa: F401
