"""Overleaf tool: sync Overleaf LaTeX projects with GitHub."""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("overleaf")

CONFIG_DIR = Path("~/.mcp_gerard").expanduser()
PROJECTS_FILE = CONFIG_DIR / "overleaf_projects.json"
OVERLEAF_TOKEN = os.environ.get("OVERLEAF_TOKEN", "")


def _load_projects() -> list[dict]:
    if not PROJECTS_FILE.exists():
        return []
    try:
        return json.loads(PROJECTS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []


def _save_projects(projects: list[dict]) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    PROJECTS_FILE.write_text(json.dumps(projects, indent=2), encoding="utf-8")


def _get_project(name: str) -> dict | None:
    for p in _load_projects():
        if p.get("name") == name:
            return p
    return None


def _authed_url(url: str) -> str:
    """Inject OVERLEAF_TOKEN into an overleaf git URL."""
    if OVERLEAF_TOKEN and "git.overleaf.com" in url:
        return url.replace("https://", f"https://git:{OVERLEAF_TOKEN}@")
    return url


def _run_git(args: list[str], cwd: str, timeout: int = 60) -> tuple[int, str]:
    result = subprocess.run(
        ["git"] + args,
        cwd=cwd,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    output = (result.stdout + result.stderr).strip()
    return result.returncode, output


@mcp.tool()
def overleaf_list_projects() -> list[dict]:
    """List all known Overleaf projects and their sync status.

    Reads from ~/.mcp_gerard/overleaf_projects.json.
    Returns a list of project dicts with name, overleaf_url, local_path, github_repo.
    """
    projects = _load_projects()
    result = []
    for p in projects:
        local_path = Path(p.get("local_path", "")).expanduser()
        info = dict(p)
        info["local_exists"] = local_path.exists()
        # Get last commit if local repo exists
        if local_path.exists():
            rc, out = _run_git(
                ["log", "-1", "--format=%h %s (%cr)", "--"], str(local_path)
            )
            info["last_commit"] = out if rc == 0 else "unknown"
        result.append(info)
    return result


@mcp.tool()
def overleaf_sync(project_name: str, direction: str = "both") -> str:
    """Sync an Overleaf project with local and GitHub.

    direction: 'pull' (overleaf→local), 'push' (local→overleaf), 'both' (pull then push).
    Also pushes to GitHub origin after sync.
    Returns a summary of what was synced.
    """
    if direction not in ("pull", "push", "both"):
        raise ValueError(f"Invalid direction '{direction}'. Must be pull, push, or both.")

    project = _get_project(project_name)
    if not project:
        return f"Error: project '{project_name}' not found. Use overleaf_list_projects() to see known projects."

    local_path = str(Path(project["local_path"]).expanduser())
    overleaf_url = _authed_url(project["overleaf_url"])
    summary: list[str] = []

    # Ensure overleaf remote exists
    rc, remotes = _run_git(["remote"], local_path)
    if "overleaf" not in remotes.split():
        _run_git(["remote", "add", "overleaf", overleaf_url], local_path)
        summary.append("Added overleaf remote.")
    else:
        # Update remote URL in case token changed
        _run_git(["remote", "set-url", "overleaf", overleaf_url], local_path)

    if direction in ("pull", "both"):
        rc, out = _run_git(["pull", "overleaf", "master", "--no-rebase"], local_path)
        summary.append(f"Pull from Overleaf: {out[:200]}")

    if direction in ("push", "both"):
        rc, out = _run_git(["push", "overleaf", "HEAD:master"], local_path)
        summary.append(f"Push to Overleaf: {out[:200]}")

    # Always push to GitHub origin
    rc, out = _run_git(["push", "origin"], local_path)
    summary.append(f"Push to GitHub: {out[:200]}")

    return "\n".join(summary)


@mcp.tool()
def overleaf_add_project(
    name: str, overleaf_url: str, local_path: str, github_repo: str
) -> str:
    """Register a new Overleaf project and clone it locally.

    Adds entry to ~/.mcp_gerard/overleaf_projects.json.
    Clones the Overleaf repo if local_path doesn't exist.
    Sets up 'origin' remote pointing to github_repo.
    Returns a confirmation message.
    """
    projects = _load_projects()
    if any(p["name"] == name for p in projects):
        return f"Error: project '{name}' already exists."

    local = Path(local_path).expanduser()
    authed_url = _authed_url(overleaf_url)

    if not local.exists():
        local.parent.mkdir(parents=True, exist_ok=True)
        result = subprocess.run(
            ["git", "clone", authed_url, str(local)],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            return f"Error cloning Overleaf repo: {result.stderr}"

    # Set up origin remote
    rc, remotes = _run_git(["remote"], str(local))
    if "origin" not in remotes.split():
        _run_git(["remote", "add", "origin", github_repo], str(local))
    else:
        _run_git(["remote", "set-url", "origin", github_repo], str(local))

    entry = {
        "name": name,
        "overleaf_url": overleaf_url,
        "local_path": local_path,
        "github_repo": github_repo,
    }
    projects.append(entry)
    _save_projects(projects)

    return f"Added project '{name}'. Local path: {local}. GitHub: {github_repo}."


@mcp.tool()
def overleaf_compile(project_name: str) -> str:
    """Compile an Overleaf project with latexmk.

    Runs 'latexmk -pdf main.tex' in the project directory.
    Returns the path to the compiled PDF or an error log excerpt.
    """
    project = _get_project(project_name)
    if not project:
        return f"Error: project '{project_name}' not found."

    local_path = Path(project["local_path"]).expanduser()
    if not local_path.exists():
        return f"Error: local path does not exist: {local_path}"

    try:
        result = subprocess.run(
            ["latexmk", "-pdf", "main.tex"],
            cwd=str(local_path),
            capture_output=True,
            text=True,
            timeout=300,
        )
    except FileNotFoundError:
        return "Error: latexmk not found. Install a LaTeX distribution (e.g. MiKTeX or TeX Live)."
    except subprocess.TimeoutExpired:
        return "Error: compilation timed out after 5 minutes."

    pdf_path = local_path / "main.pdf"
    if result.returncode == 0 and pdf_path.exists():
        return f"Compiled successfully: {pdf_path}"

    # Return last 50 lines of log on error
    log_lines = (result.stdout + result.stderr).splitlines()
    excerpt = "\n".join(log_lines[-50:])
    return f"Compilation failed. Log excerpt:\n{excerpt}"


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
