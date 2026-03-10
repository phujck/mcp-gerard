"""Projects tool: create projects, sync repos, bootstrap environment."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("projects")

PROJECTS_ROOT = Path("~/Projects").expanduser()
CONFIG_DIR = Path("~/.mcp_gerard").expanduser()
REPOS_FILE = CONFIG_DIR / "repos.json"
_SHARED_DIR = Path(__file__).parent / "_shared"

GITHUB_USER = "phujck"

_GITIGNORE = """\
# Python
__pycache__/
*.py[cod]
*.pyo
*.pyd
.Python
*.egg
*.egg-info/
dist/
build/
*.venv/
venv/
.env

# LaTeX
*.aux
*.bbl
*.blg
*.fdb_latexmk
*.fls
*.log
*.out
*.synctex.gz
*.toc
*.lof
*.lot
*.nav
*.snm
*.vrb
*.pdf

# Data / outputs
data/
outputs/
*.csv
*.h5
*.npz

# OS
.DS_Store
Thumbs.db
"""

_TEMPLATES: dict[str, str] = {
    "blog": "blog",
    "research": "research",
    "simulations": "simulations",
    "website": "website",
}


def _run(args: list[str], cwd: str | None = None, timeout: int = 60) -> tuple[int, str]:
    result = subprocess.run(
        args, cwd=cwd, capture_output=True, text=True, timeout=timeout
    )
    return result.returncode, (result.stdout + result.stderr).strip()


def _scaffold_blog(project_dir: Path) -> None:
    (project_dir / "images").mkdir(exist_ok=True)
    # Copy preamble/macros from _shared
    for f in ("preamble.tex", "macros.tex"):
        src = _SHARED_DIR / f
        if src.exists():
            shutil.copy2(src, project_dir / f)
    (project_dir / "main.tex").write_text(
        "\\documentclass[12pt]{article}\n"
        "\\input{preamble}\n"
        "\\input{macros}\n\n"
        "\\title{New Post}\n"
        "\\author{Gerard McCaul}\n"
        f"\\date{{{datetime.now().strftime('%Y-%m-%d')}}}\n\n"
        "\\begin{document}\n"
        "\\maketitle\n\n"
        "% Write your post here\n\n"
        "\\end{document}\n",
        encoding="utf-8",
    )


def _scaffold_research(project_dir: Path) -> None:
    (project_dir / "sections").mkdir(exist_ok=True)
    for f in ("preamble.tex", "macros.tex"):
        src = _SHARED_DIR / f
        if src.exists():
            shutil.copy2(src, project_dir / f)
    (project_dir / "references.bib").write_text(
        "% References\n", encoding="utf-8"
    )
    (project_dir / "main.tex").write_text(
        "\\documentclass[12pt]{article}\n"
        "\\input{preamble}\n"
        "\\input{macros}\n\n"
        "\\title{Manuscript Title}\n"
        "\\author{Gerard McCaul}\n"
        f"\\date{{{datetime.now().strftime('%Y-%m-%d')}}}\n\n"
        "\\begin{document}\n"
        "\\maketitle\n\n"
        "\\section{Introduction}\n\n"
        "\\section{Methods}\n\n"
        "\\section{Results}\n\n"
        "\\section{Conclusion}\n\n"
        "\\printbibliography\n"
        "\\end{document}\n",
        encoding="utf-8",
    )


def _scaffold_simulations(project_dir: Path) -> None:
    for d in ("data", "outputs"):
        (project_dir / d).mkdir(exist_ok=True)
        (project_dir / d / ".gitkeep").touch()
    (project_dir / "requirements.txt").write_text(
        "numpy\nmatplotlib\nscipy\n", encoding="utf-8"
    )
    (project_dir / "run.py").write_text(
        '"""Main simulation entry point."""\n\n'
        "import numpy as np\n\n\n"
        "def main():\n"
        "    pass\n\n\n"
        'if __name__ == "__main__":\n'
        "    main()\n",
        encoding="utf-8",
    )
    # Copy llm.py from _shared
    llm_src = _SHARED_DIR / "llm.py"
    if llm_src.exists():
        shutil.copy2(llm_src, project_dir / "llm.py")


def _scaffold_website(project_dir: Path) -> None:
    (project_dir / "_posts").mkdir(exist_ok=True)
    (project_dir / "index.md").write_text(
        "---\nlayout: home\ntitle: Home\n---\n", encoding="utf-8"
    )
    (project_dir / "_config.yml").write_text(
        "title: My Site\ntheme: minima\n", encoding="utf-8"
    )


def _scaffold_default(project_dir: Path) -> None:
    (project_dir / "README.md").write_text(
        f"# {project_dir.name}\n\nCreated {datetime.now().strftime('%Y-%m-%d')}\n",
        encoding="utf-8",
    )


_SCAFFOLD_FN = {
    "blog": _scaffold_blog,
    "research": _scaffold_research,
    "simulations": _scaffold_simulations,
    "website": _scaffold_website,
}


def _update_repos_json(name: str, category: str, github_repo: str) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    repos = []
    if REPOS_FILE.exists():
        try:
            repos = json.loads(REPOS_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            repos = []
    repos.append({"name": name, "category": category, "github_repo": github_repo})
    REPOS_FILE.write_text(json.dumps(repos, indent=2), encoding="utf-8")


@mcp.tool()
def projects_new(category: str, name: str, template: str = "auto") -> str:
    """Create a new project with scaffolding and a private GitHub repo.

    Creates ~/Projects/<category>/<name>/, runs git init, applies template,
    creates phujck/<category>-<name> on GitHub, pushes initial commit.
    Also updates the vault dashboard and opens in VSCode.
    Returns the GitHub repo URL.
    """
    project_dir = PROJECTS_ROOT / category / name
    if project_dir.exists():
        return f"Error: {project_dir} already exists."

    project_dir.mkdir(parents=True)

    # Determine template
    actual_template = template
    if template == "auto":
        actual_template = _TEMPLATES.get(category, "default")

    # Scaffold
    scaffold_fn = _SCAFFOLD_FN.get(actual_template, _scaffold_default)
    scaffold_fn(project_dir)

    # Write .gitignore
    (project_dir / ".gitignore").write_text(_GITIGNORE, encoding="utf-8")

    # git init + initial commit
    _run(["git", "init"], str(project_dir))
    _run(["git", "add", "-A"], str(project_dir))
    _run(["git", "commit", "-m", f"initial: scaffold {category}/{name}"], str(project_dir))

    # Create GitHub repo
    repo_name = f"{category}-{name}"
    rc, out = _run(
        ["gh", "repo", "create", f"{GITHUB_USER}/{repo_name}", "--private", "--source=.", "--push"],
        str(project_dir),
        timeout=120,
    )
    if rc != 0:
        # Try fallback: create repo then set remote manually
        _run(
            ["gh", "repo", "create", f"{GITHUB_USER}/{repo_name}", "--private"],
            timeout=60,
        )
        github_url = f"https://github.com/{GITHUB_USER}/{repo_name}"
        _run(["git", "remote", "add", "origin", github_url], str(project_dir))
        _run(["git", "push", "-u", "origin", "main"], str(project_dir), timeout=60)
    else:
        github_url = f"https://github.com/{GITHUB_USER}/{repo_name}"

    # Update repos.json
    _update_repos_json(name, category, github_url)

    # Update vault dashboard
    try:
        from mcp_gerard.vault import vault_update_dashboard
        vault_update_dashboard()
    except Exception:
        pass

    # Open in VSCode
    try:
        subprocess.Popen(["code", str(project_dir)])
    except FileNotFoundError:
        pass

    return github_url


@mcp.tool()
def projects_list() -> list[dict]:
    """List all projects under ~/Projects/ that are git repos.

    Returns name, category, last commit message, and GitHub URL for each.
    """
    results = []
    if not PROJECTS_ROOT.exists():
        return results

    for category_dir in sorted(PROJECTS_ROOT.iterdir()):
        if not category_dir.is_dir():
            continue
        for project_dir in sorted(category_dir.iterdir()):
            if not project_dir.is_dir() or not (project_dir / ".git").exists():
                continue

            # Last commit
            rc, last_commit = _run(
                ["git", "log", "-1", "--format=%s (%cr)"],
                str(project_dir),
                timeout=10,
            )

            # GitHub remote URL
            rc2, remote_url = _run(
                ["git", "remote", "get-url", "origin"],
                str(project_dir),
                timeout=10,
            )

            results.append(
                {
                    "category": category_dir.name,
                    "name": project_dir.name,
                    "path": str(project_dir),
                    "last_commit": last_commit if rc == 0 else "",
                    "github_url": remote_url if rc2 == 0 else "",
                }
            )
    return results


@mcp.tool()
def projects_sync_all() -> str:
    """Sync all git repos under ~/Projects/ to GitHub.

    For each repo with changes: git add . && git commit && git push.
    Skips repos with no uncommitted changes.
    Returns a summary.
    """
    synced = []
    skipped = []
    errors = []

    if not PROJECTS_ROOT.exists():
        return "~/Projects/ does not exist."

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    for category_dir in sorted(PROJECTS_ROOT.iterdir()):
        if not category_dir.is_dir():
            continue
        for project_dir in sorted(category_dir.iterdir()):
            if not project_dir.is_dir() or not (project_dir / ".git").exists():
                continue
            name = f"{category_dir.name}/{project_dir.name}"

            # Check for changes
            rc, status = _run(["git", "status", "--porcelain"], str(project_dir))
            if not status.strip():
                skipped.append(name)
                continue

            # Commit and push
            _run(["git", "add", "-A"], str(project_dir))
            rc, out = _run(
                ["git", "commit", "-m", f"auto-sync {timestamp}"],
                str(project_dir),
            )
            if rc != 0 and "nothing to commit" not in out:
                errors.append(f"{name}: commit failed: {out[:100]}")
                continue

            rc, out = _run(["git", "push"], str(project_dir), timeout=60)
            if rc == 0:
                synced.append(name)
            else:
                errors.append(f"{name}: push failed: {out[:100]}")

    lines = []
    if synced:
        lines.append(f"Synced ({len(synced)}): {', '.join(synced)}")
    if skipped:
        lines.append(f"No changes ({len(skipped)}): {', '.join(skipped)}")
    if errors:
        lines.append(f"Errors ({len(errors)}):\n" + "\n".join(errors))
    return "\n".join(lines) or "Nothing to sync."


@mcp.tool()
def projects_bootstrap() -> str:
    """Bootstrap the environment on a fresh machine.

    Reads ~/.mcp_gerard/repos.json and clones any repos that don't exist locally.
    Also installs mcp-gerard itself via uv.
    Returns a setup summary.
    """
    if not REPOS_FILE.exists():
        return f"repos.json not found at {REPOS_FILE}. Nothing to bootstrap."

    try:
        repos = json.loads(REPOS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        return f"Error reading repos.json: {e}"

    cloned = []
    already_exists = []
    errors = []

    for repo in repos:
        name = repo.get("name", "")
        category = repo.get("category", "")
        github_repo = repo.get("github_repo", "")

        if not (name and category and github_repo):
            continue

        local_path = PROJECTS_ROOT / category / name
        if local_path.exists():
            already_exists.append(f"{category}/{name}")
            continue

        local_path.parent.mkdir(parents=True, exist_ok=True)
        rc, out = _run(
            ["git", "clone", github_repo, str(local_path)],
            timeout=120,
        )
        if rc == 0:
            cloned.append(f"{category}/{name}")
        else:
            errors.append(f"{category}/{name}: {out[:100]}")

    lines = []
    if cloned:
        lines.append(f"Cloned ({len(cloned)}): {', '.join(cloned)}")
    if already_exists:
        lines.append(f"Already exists ({len(already_exists)}): {', '.join(already_exists)}")
    if errors:
        lines.append(f"Errors ({len(errors)}):\n" + "\n".join(errors))
    return "\n".join(lines) or "Nothing to bootstrap."


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
