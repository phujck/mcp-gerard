"""Unit tests for the vault tool."""

import pytest
from pathlib import Path
from datetime import datetime
from unittest.mock import patch, MagicMock


@pytest.mark.unit
class TestVaultCapture:
    def test_creates_ideas_file(self, tmp_path):
        from mcp_gerard import vault
        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch("mcp_gerard.vault._git_commit"):
            result = vault.vault_capture("My great idea")

        ideas_dir = tmp_path / "ideas"
        assert ideas_dir.is_dir()
        # Find the monthly file
        monthly_files = list(ideas_dir.glob("*.md"))
        assert len(monthly_files) == 1
        content = monthly_files[0].read_text(encoding="utf-8")
        assert "My great idea" in content
        assert "Captured to" in result

    def test_appends_tags(self, tmp_path):
        from mcp_gerard import vault
        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch("mcp_gerard.vault._git_commit"):
            vault.vault_capture("Tag test", tags=["physics", "todo"])

        monthly_files = list((tmp_path / "ideas").glob("*.md"))
        content = monthly_files[0].read_text(encoding="utf-8")
        assert "#physics" in content
        assert "#todo" in content

    def test_appends_to_existing_file(self, tmp_path):
        from mcp_gerard import vault
        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch("mcp_gerard.vault._git_commit"):
            vault.vault_capture("First idea")
            vault.vault_capture("Second idea")

        monthly_files = list((tmp_path / "ideas").glob("*.md"))
        content = monthly_files[0].read_text(encoding="utf-8")
        assert "First idea" in content
        assert "Second idea" in content


@pytest.mark.unit
class TestVaultSearch:
    def test_finds_match_in_notes(self, tmp_path):
        from mcp_gerard import vault
        notes_dir = tmp_path / "notes"
        notes_dir.mkdir(parents=True)
        (notes_dir / "test.md").write_text("This is about thermodynamics.", encoding="utf-8")
        # Create other required dirs
        (tmp_path / "ideas").mkdir()
        (tmp_path / "references").mkdir()
        (tmp_path / "_index.md").write_text("# Dashboard", encoding="utf-8")

        with patch.object(vault, "VAULT_PATH", tmp_path):
            results = vault.vault_search("thermodynamics")

        assert len(results) == 1
        assert "thermodynamics" in results[0]["match"]
        assert results[0]["line"] == 1

    def test_returns_empty_for_no_match(self, tmp_path):
        from mcp_gerard import vault
        notes_dir = tmp_path / "notes"
        notes_dir.mkdir(parents=True)
        (notes_dir / "test.md").write_text("Nothing relevant here.", encoding="utf-8")
        for d in ("ideas", "references"):
            (tmp_path / d).mkdir()

        with patch.object(vault, "VAULT_PATH", tmp_path):
            results = vault.vault_search("quantum")

        assert results == []

    def test_case_insensitive_search(self, tmp_path):
        from mcp_gerard import vault
        notes_dir = tmp_path / "notes"
        notes_dir.mkdir(parents=True)
        (notes_dir / "test.md").write_text("ENTROPY matters.", encoding="utf-8")
        for d in ("ideas", "references"):
            (tmp_path / d).mkdir()

        with patch.object(vault, "VAULT_PATH", tmp_path):
            results = vault.vault_search("entropy")

        assert len(results) == 1


@pytest.mark.unit
class TestVaultNewNote:
    def test_creates_note_in_notes_dir(self, tmp_path):
        from mcp_gerard import vault
        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch("mcp_gerard.vault._git_commit"):
            result = vault.vault_new_note("My Note")

        path = Path(result)
        assert path.exists()
        assert path.parent == tmp_path / "notes"
        content = path.read_text(encoding="utf-8")
        assert "# My Note" in content

    def test_invalid_category_defaults_to_notes(self, tmp_path):
        from mcp_gerard import vault
        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch("mcp_gerard.vault._git_commit"):
            result = vault.vault_new_note("Test", category="invalid")

        path = Path(result)
        assert "notes" in str(path)

    def test_references_category(self, tmp_path):
        from mcp_gerard import vault
        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch("mcp_gerard.vault._git_commit"):
            result = vault.vault_new_note("Paper Summary", category="references")

        assert "references" in result


@pytest.mark.unit
class TestVaultDashboard:
    def test_creates_index_file(self, tmp_path):
        from mcp_gerard import vault
        # Empty projects root
        projects_root = tmp_path / "Projects"
        projects_root.mkdir()

        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch.object(vault, "PROJECTS_ROOT", projects_root), \
             patch("mcp_gerard.vault._git_commit"):
            result = vault.vault_update_dashboard()

        index = tmp_path / "_index.md"
        assert index.exists()
        assert "# Project Dashboard" in result

    def test_table_format_with_repos(self, tmp_path):
        from mcp_gerard import vault
        # Create a fake git repo
        projects_root = tmp_path / "Projects"
        repo_dir = projects_root / "research" / "my-paper"
        repo_dir.mkdir(parents=True)
        (repo_dir / ".git").mkdir()

        with patch.object(vault, "VAULT_PATH", tmp_path), \
             patch.object(vault, "PROJECTS_ROOT", projects_root), \
             patch("subprocess.run") as mock_run, \
             patch("mcp_gerard.vault._git_commit"):
            mock_run.return_value = MagicMock(
                returncode=0,
                stdout="abc1234|initial commit|2 days ago\n"
            )
            result = vault.vault_update_dashboard()

        assert "| research |" in result
        assert "| my-paper |" in result
        assert "|----------|" in result  # table separator
