"""Unit tests for the overleaf tool."""

import json
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock


@pytest.mark.unit
class TestLoadSaveProjects:
    def test_load_empty_when_file_missing(self, tmp_path):
        from mcp_gerard import overleaf
        with patch.object(overleaf, "PROJECTS_FILE", tmp_path / "missing.json"):
            result = overleaf._load_projects()
        assert result == []

    def test_load_returns_projects(self, tmp_path):
        from mcp_gerard import overleaf
        data = [{"name": "thesis", "overleaf_url": "https://git.overleaf.com/abc"}]
        projects_file = tmp_path / "projects.json"
        projects_file.write_text(json.dumps(data), encoding="utf-8")
        with patch.object(overleaf, "PROJECTS_FILE", projects_file):
            result = overleaf._load_projects()
        assert result == data

    def test_save_creates_file(self, tmp_path):
        from mcp_gerard import overleaf
        projects_file = tmp_path / "projects.json"
        data = [{"name": "test"}]
        with patch.object(overleaf, "CONFIG_DIR", tmp_path), \
             patch.object(overleaf, "PROJECTS_FILE", projects_file):
            overleaf._save_projects(data)
        assert json.loads(projects_file.read_text()) == data


@pytest.mark.unit
class TestListProjects:
    def test_empty_when_no_projects(self, tmp_path):
        from mcp_gerard import overleaf
        with patch.object(overleaf, "PROJECTS_FILE", tmp_path / "missing.json"):
            result = overleaf.overleaf_list_projects()
        assert result == []

    def test_includes_local_exists_flag(self, tmp_path):
        from mcp_gerard import overleaf
        projects_file = tmp_path / "projects.json"
        data = [{"name": "thesis", "local_path": str(tmp_path / "nonexistent"), "overleaf_url": ""}]
        projects_file.write_text(json.dumps(data), encoding="utf-8")

        with patch.object(overleaf, "PROJECTS_FILE", projects_file):
            result = overleaf.overleaf_list_projects()

        assert result[0]["local_exists"] is False


@pytest.mark.unit
class TestSyncDirectionValidation:
    def test_invalid_direction_raises(self):
        from mcp_gerard import overleaf
        with pytest.raises(ValueError, match="Invalid direction"):
            overleaf.overleaf_sync("myproject", direction="sideways")

    def test_valid_directions_accepted(self):
        """Valid directions should not raise ValueError (may fail for other reasons in unit context)."""
        from mcp_gerard import overleaf
        with patch.object(overleaf, "PROJECTS_FILE", Path("/nonexistent/projects.json")):
            # Will fail because project not found, but not ValueError for direction
            result = overleaf.overleaf_sync("nonexistent", direction="both")
            assert "not found" in result


@pytest.mark.unit
class TestAddProject:
    def test_add_project_writes_json(self, tmp_path):
        from mcp_gerard import overleaf
        projects_file = tmp_path / "projects.json"
        projects_file.write_text("[]", encoding="utf-8")

        local = tmp_path / "local_repo"
        local.mkdir()

        with patch.object(overleaf, "CONFIG_DIR", tmp_path), \
             patch.object(overleaf, "PROJECTS_FILE", projects_file), \
             patch("subprocess.run") as mock_run, \
             patch.object(overleaf, "_run_git", return_value=(0, "origin")):
            mock_run.return_value = MagicMock(returncode=0)
            result = overleaf.overleaf_add_project(
                name="thesis",
                overleaf_url="https://git.overleaf.com/abc123",
                local_path=str(local),
                github_repo="https://github.com/phujck/thesis",
            )

        projects = json.loads(projects_file.read_text())
        assert len(projects) == 1
        assert projects[0]["name"] == "thesis"
        assert "Added project" in result

    def test_duplicate_name_rejected(self, tmp_path):
        from mcp_gerard import overleaf
        existing = [{"name": "thesis"}]
        projects_file = tmp_path / "projects.json"
        projects_file.write_text(json.dumps(existing), encoding="utf-8")

        with patch.object(overleaf, "PROJECTS_FILE", projects_file):
            result = overleaf.overleaf_add_project(
                "thesis", "url", "path", "repo"
            )
        assert "Error" in result
        assert "already exists" in result


@pytest.mark.unit
class TestAuthedUrl:
    def test_injects_token(self):
        from mcp_gerard import overleaf
        with patch.object(overleaf, "OVERLEAF_TOKEN", "mytoken"):
            result = overleaf._authed_url("https://git.overleaf.com/abc123")
        assert "git:mytoken@" in result

    def test_no_injection_without_token(self):
        from mcp_gerard import overleaf
        with patch.object(overleaf, "OVERLEAF_TOKEN", ""):
            result = overleaf._authed_url("https://git.overleaf.com/abc123")
        assert "git:" not in result
