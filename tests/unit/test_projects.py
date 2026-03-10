"""Unit tests for the projects tool."""

import json
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock, call


@pytest.mark.unit
class TestGitignoreContent:
    def test_gitignore_contains_latex_patterns(self, tmp_path):
        from mcp_gerard.projects import _GITIGNORE
        for pattern in ("*.aux", "*.log", "*.pdf", "*.synctex.gz"):
            assert pattern in _GITIGNORE

    def test_gitignore_contains_python_patterns(self, tmp_path):
        from mcp_gerard.projects import _GITIGNORE
        for pattern in ("__pycache__/", "*.py[cod]", ".env"):
            assert pattern in _GITIGNORE


@pytest.mark.unit
class TestTemplateSelection:
    def test_simulations_template_auto(self, tmp_path):
        from mcp_gerard import projects
        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")), \
             patch("subprocess.Popen"), \
             patch("mcp_gerard.projects._update_repos_json"):
            # Patch vault_update_dashboard import
            with patch.dict("sys.modules", {"mcp_gerard.vault": MagicMock()}):
                projects.projects_new("simulations", "my-sim")

        project_dir = tmp_path / "simulations" / "my-sim"
        assert project_dir.exists()
        assert (project_dir / "run.py").exists()
        assert (project_dir / "data").is_dir()
        assert (project_dir / "outputs").is_dir()
        assert (project_dir / "requirements.txt").exists()

    def test_research_template_auto(self, tmp_path):
        from mcp_gerard import projects
        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")), \
             patch("subprocess.Popen"), \
             patch("mcp_gerard.projects._update_repos_json"):
            with patch.dict("sys.modules", {"mcp_gerard.vault": MagicMock()}):
                projects.projects_new("research", "my-paper")

        project_dir = tmp_path / "research" / "my-paper"
        assert (project_dir / "main.tex").exists()
        assert (project_dir / "sections").is_dir()
        assert (project_dir / "references.bib").exists()

    def test_blog_template_auto(self, tmp_path):
        from mcp_gerard import projects
        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")), \
             patch("subprocess.Popen"), \
             patch("mcp_gerard.projects._update_repos_json"):
            with patch.dict("sys.modules", {"mcp_gerard.vault": MagicMock()}):
                projects.projects_new("blog", "my-post")

        project_dir = tmp_path / "blog" / "my-post"
        assert (project_dir / "main.tex").exists()
        assert (project_dir / "images").is_dir()

    def test_default_template_creates_readme(self, tmp_path):
        from mcp_gerard import projects
        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")), \
             patch("subprocess.Popen"), \
             patch("mcp_gerard.projects._update_repos_json"):
            with patch.dict("sys.modules", {"mcp_gerard.vault": MagicMock()}):
                projects.projects_new("misc", "random-thing")

        project_dir = tmp_path / "misc" / "random-thing"
        assert (project_dir / "README.md").exists()


@pytest.mark.unit
class TestPathConstruction:
    def test_creates_correct_path(self, tmp_path):
        from mcp_gerard import projects
        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")), \
             patch("subprocess.Popen"), \
             patch("mcp_gerard.projects._update_repos_json"):
            with patch.dict("sys.modules", {"mcp_gerard.vault": MagicMock()}):
                projects.projects_new("simulations", "test-exp")

        expected = tmp_path / "simulations" / "test-exp"
        assert expected.is_dir()

    def test_duplicate_raises_error(self, tmp_path):
        from mcp_gerard import projects
        (tmp_path / "simulations" / "existing").mkdir(parents=True)
        with patch.object(projects, "PROJECTS_ROOT", tmp_path):
            result = projects.projects_new("simulations", "existing")
        assert "Error" in result
        assert "already exists" in result


@pytest.mark.unit
class TestProjectsList:
    def test_returns_git_repos_only(self, tmp_path):
        from mcp_gerard import projects
        # Create one git repo and one plain directory
        repo = tmp_path / "research" / "paper"
        repo.mkdir(parents=True)
        (repo / ".git").mkdir()

        non_repo = tmp_path / "research" / "notes"
        non_repo.mkdir(parents=True)

        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")):
            result = projects.projects_list()

        names = [r["name"] for r in result]
        assert "paper" in names
        assert "notes" not in names

    def test_includes_category_info(self, tmp_path):
        from mcp_gerard import projects
        repo = tmp_path / "blog" / "my-post"
        repo.mkdir(parents=True)
        (repo / ".git").mkdir()

        with patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "")):
            result = projects.projects_list()

        assert result[0]["category"] == "blog"
        assert result[0]["name"] == "my-post"


@pytest.mark.unit
class TestBootstrap:
    def test_missing_repos_json(self, tmp_path):
        from mcp_gerard import projects
        with patch.object(projects, "REPOS_FILE", tmp_path / "missing.json"):
            result = projects.projects_bootstrap()
        assert "not found" in result

    def test_clones_missing_repos(self, tmp_path):
        from mcp_gerard import projects
        repos_file = tmp_path / "repos.json"
        repos_file.write_text(
            json.dumps([{"name": "my-sim", "category": "simulations", "github_repo": "https://github.com/phujck/simulations-my-sim"}]),
            encoding="utf-8",
        )
        with patch.object(projects, "REPOS_FILE", repos_file), \
             patch.object(projects, "PROJECTS_ROOT", tmp_path), \
             patch("mcp_gerard.projects._run", return_value=(0, "Cloning...")) as mock_run:
            result = projects.projects_bootstrap()

        assert "simulations/my-sim" in result
        assert "Cloned" in result
