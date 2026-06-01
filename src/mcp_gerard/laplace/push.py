"""Explicit canon push - the one git.exe action, deliberately off the hot path.

Versioning the canon (commit / log / revert) is pure-Python Dulwich (see gitio),
so the constant hot path can never spawn git.exe or deadlock. Publishing to
GitHub is different: it needs the user's existing git credential setup, which
Dulwich does not reuse. So pushing shells ``git push`` - but only as an explicit,
occasional action a human or agent triggers, never in the read/assess/dream
loop. It still captures via temp files (no pipe for a git-fsmonitor--daemon to
inherit), so even this one subprocess is deadlock-proof, and the tool watchdog
bounds it regardless.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from typing import Any


def push_canon(root, remote: str = "origin", timeout: int = 120) -> dict[str, Any]:
    """Push the canon repo at ``root`` to ``remote`` using the system git creds."""
    root = Path(root)
    if not (root / ".git").exists():
        return {"ok": False, "error": f"no git repo at {root}"}
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as out, tempfile.TemporaryFile(
        mode="w+", encoding="utf-8", errors="replace"
    ) as err:
        try:
            proc = subprocess.run(
                ["git", "-c", "core.fsmonitor=false", "-C", str(root), "push", remote],
                stdout=out,
                stderr=err,
                timeout=timeout,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            return {"ok": False, "error": str(exc), "remote": remote, "root": str(root)}
        out.seek(0)
        err.seek(0)
        return {
            "ok": proc.returncode == 0,
            "returncode": proc.returncode,
            "stdout": out.read(),
            "stderr": err.read(),
            "remote": remote,
            "root": str(root),
        }
