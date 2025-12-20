"""Unified Chat Tool - mcp-chat entry point.

This module provides a unified interface to multiple LLM providers through
model-based provider inference. Instead of using separate tools per provider,
users specify the model and the provider is automatically determined.
"""

from mcp_handley_lab.chat.tool import mcp

__all__ = ["mcp"]
