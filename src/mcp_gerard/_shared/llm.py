"""Thin Anthropic API wrapper for use in simulation projects.

Copy this file into your project directory and import as:
    from llm import ask, ask_stream
"""

from __future__ import annotations

import os
from typing import Iterator

import anthropic

_client: anthropic.Anthropic | None = None


def _get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY environment variable not set.")
        _client = anthropic.Anthropic(api_key=api_key)
    return _client


def ask(
    prompt: str,
    model: str = "claude-sonnet-4-6",
    system: str = "",
    max_tokens: int = 8096,
) -> str:
    """Send a single prompt and return the response text."""
    client = _get_client()
    kwargs: dict = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}],
    }
    if system:
        kwargs["system"] = system
    response = client.messages.create(**kwargs)
    return response.content[0].text


def ask_stream(
    prompt: str,
    model: str = "claude-sonnet-4-6",
    system: str = "",
    max_tokens: int = 8096,
) -> Iterator[str]:
    """Stream a single prompt, yielding text chunks as they arrive."""
    client = _get_client()
    kwargs: dict = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}],
    }
    if system:
        kwargs["system"] = system
    with client.messages.stream(**kwargs) as stream:
        for text in stream.text_stream:
            yield text
