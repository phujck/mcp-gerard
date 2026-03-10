"""Vault tool: capture ideas, search notes, maintain project dashboard."""

from __future__ import annotations

import os
import subprocess
from datetime import datetime
from pathlib import Path

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("vault")

VAULT_PATH = Path(os.environ.get("VAULT_PATH", "~/Projects/vault")).expanduser()
PROJECTS_ROOT = Path("~/Projects").expanduser()


def _git_commit(path: Path, message: str) -> None:
    """Stage all changes and commit in the given git repo."""
    try:
        subprocess.run(
            ["git", "add", "-A"],
            cwd=str(path),
            capture_output=True,
            timeout=30,
        )
        subprocess.run(
            ["git", "commit", "-m", message],
            cwd=str(path),
            capture_output=True,
            timeout=30,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass


def _ensure_vault() -> None:
    """Create vault directory structure if it doesn't exist."""
    for subdir in ("ideas", "notes", "references"):
        (VAULT_PATH / subdir).mkdir(parents=True, exist_ok=True)
    if not (VAULT_PATH / "_index.md").exists():
        (VAULT_PATH / "_index.md").write_text("# Project Dashboard\n\n", encoding="utf-8")


@mcp.tool()
def vault_capture(text: str, tags: list[str] = []) -> str:  # noqa: B006
    """Capture a quick idea to the vault.

    Appends a timestamped entry to ideas/YYYY-MM.md with optional #tags.
    Returns confirmation with the file path.
    """
    _ensure_vault()
    now = datetime.now()
    ideas_file = VAULT_PATH / "ideas" / f"{now.strftime('%Y-%m')}.md"
    tag_str = "  " + " ".join(f"#{t}" for t in tags) if tags else ""
    entry = f"\n## {now.strftime('%Y-%m-%d %H:%M')}{tag_str}\n\n{text}\n"

    with ideas_file.open("a", encoding="utf-8") as f:
        f.write(entry)

    _git_commit(VAULT_PATH, f"capture: {text[:50]!r}")
    return f"Captured to {ideas_file}"


@mcp.tool()
def vault_search(query: str) -> list[dict]:
    """Full-text search across all .md files in the vault.

    Returns a list of matches with file, line number, and surrounding context.
    """
    _ensure_vault()
    matches = []
    query_lower = query.lower()

    for md_file in sorted(VAULT_PATH.rglob("*.md")):
        try:
            lines = md_file.read_text(encoding="utf-8", errors="ignore").splitlines()
        except OSError:
            continue
        for i, line in enumerate(lines):
            if query_lower in line.lower():
                context_start = max(0, i - 1)
                context_end = min(len(lines), i + 2)
                matches.append(
                    {
                        "file": str(md_file.relative_to(VAULT_PATH)),
                        "line": i + 1,
                        "match": line.strip(),
                        "context": "\n".join(lines[context_start:context_end]),
                    }
                )
    return matches


@mcp.tool()
def vault_new_note(title: str, category: str = "notes") -> str:
    """Create a new note in the vault.

    Creates a .md file in the appropriate subdirectory with title/date template.
    category must be one of: notes, ideas, references.
    Returns the file path.
    """
    _ensure_vault()
    valid_categories = {"notes", "ideas", "references"}
    if category not in valid_categories:
        category = "notes"

    now = datetime.now()
    slug = title.lower().replace(" ", "-").replace("/", "-")
    slug = "".join(c for c in slug if c.isalnum() or c == "-")
    filename = f"{now.strftime('%Y-%m-%d')}-{slug}.md"
    note_path = VAULT_PATH / category / filename

    content = (
        f"# {title}\n\n"
        f"*Created: {now.strftime('%Y-%m-%d')}*\n\n"
        "## Summary\n\n\n"
        "## Notes\n\n\n"
        "## References\n\n"
    )
    note_path.write_text(content, encoding="utf-8")
    _git_commit(VAULT_PATH, f"note: {title}")
    return str(note_path)


@mcp.tool()
def vault_update_dashboard() -> str:
    """Scan ~/Projects/ for all git repos and update the vault dashboard.

    Rewrites vault/_index.md with a structured table of all projects.
    Returns the updated dashboard as a string.
    """
    _ensure_vault()

    rows: list[dict] = []
    if PROJECTS_ROOT.exists():
        for entry in sorted(PROJECTS_ROOT.iterdir()):
            if not entry.is_dir():
                continue
            category = entry.name
            # Two-level: ~/Projects/<category>/<project>
            for project_dir in sorted(entry.iterdir()):
                if not project_dir.is_dir():
                    continue
                git_dir = project_dir / ".git"
                if not git_dir.exists():
                    continue
                # Get last commit info
                try:
                    result = subprocess.run(
                        ["git", "log", "-1", "--format=%h|%s|%cr"],
                        cwd=str(project_dir),
                        capture_output=True,
                        text=True,
                        timeout=10,
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        parts = result.stdout.strip().split("|", 2)
                        last_hash = parts[0] if len(parts) > 0 else ""
                        last_msg = parts[1] if len(parts) > 1 else ""
                        last_when = parts[2] if len(parts) > 2 else ""
                    else:
                        last_hash, last_msg, last_when = "", "no commits", ""
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    last_hash, last_msg, last_when = "", "error", ""

                rows.append(
                    {
                        "category": category,
                        "name": project_dir.name,
                        "last_commit": last_msg[:60],
                        "when": last_when,
                        "path": str(project_dir),
                    }
                )

    now = datetime.now()
    lines = [
        "# Project Dashboard",
        "",
        f"*Last updated: {now.strftime('%Y-%m-%d %H:%M')}*",
        "",
        "| Category | Project | Last Commit | When |",
        "|----------|---------|-------------|------|",
    ]
    for row in rows:
        lines.append(
            f"| {row['category']} | {row['name']} | {row['last_commit']} | {row['when']} |"
        )

    dashboard = "\n".join(lines) + "\n"
    index_path = VAULT_PATH / "_index.md"
    index_path.write_text(dashboard, encoding="utf-8")
    _git_commit(VAULT_PATH, "dashboard: auto-update")
    return dashboard


@mcp.tool()
def vault_dashboard() -> str:
    """Get the current project dashboard (regenerates it first).

    Equivalent to vault_update_dashboard().
    """
    return vault_update_dashboard()


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
