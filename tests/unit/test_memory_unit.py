"""Unit tests for Git-backed memory module."""

import json
from pathlib import Path

import pytest

from mcp_handley_lab.llm import memory


class TestPathEncoding:
    """Test path encoding functions."""

    def test_encode_simple_path(self):
        """Test encoding a simple Unix path."""
        result = memory.encode_project_path(Path("/home/will/project"))
        assert result == "-home-will-project"

    def test_encode_handles_resolve(self, tmp_path):
        """Test that paths are resolved before encoding."""
        sub = tmp_path / "foo"
        sub.mkdir()

        relative = tmp_path / "foo" / ".." / "foo"
        result = memory.encode_project_path(relative)

        expected = memory.encode_project_path(sub)
        assert result == expected


class TestBranchValidation:
    """Test branch name validation."""

    def test_valid_names(self, tmp_path, monkeypatch):
        """Test valid branch names don't raise."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        for name in ["session", "my_agent", "agent-1", "test.name", "Agent123"]:
            memory.validate_branch_name(name)  # Should not raise

    def test_empty_name_raises(self, tmp_path, monkeypatch):
        """Test empty name raises ValueError."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        with pytest.raises(ValueError, match="cannot be empty"):
            memory.validate_branch_name("")

    def test_invalid_chars_raises(self, tmp_path, monkeypatch):
        """Test names with invalid characters raise ValueError."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        # Note: git check-ref-format rejects these
        for name in ["../evil", "name with space", "name@{evil}"]:
            with pytest.raises(ValueError, match="Invalid branch name"):
                memory.validate_branch_name(name)


class TestBranchInputNormalization:
    """Test branch input normalization and special values."""

    def test_empty_string_returns_none(self):
        """Test empty string returns None (disable memory)."""
        assert memory.normalize_branch_input("") is None

    def test_false_returns_none(self):
        """Test 'false' returns None (disable memory)."""
        assert memory.normalize_branch_input("false") is None
        assert memory.normalize_branch_input("FALSE") is None
        assert memory.normalize_branch_input("False") is None

    def test_whitespace_only_raises(self):
        """Test whitespace-only string raises ValueError."""
        with pytest.raises(ValueError, match="cannot be whitespace-only"):
            memory.normalize_branch_input("   ")

    def test_normal_values_returned_stripped(self):
        """Test normal values are returned with whitespace stripped."""
        assert memory.normalize_branch_input("session") == "session"
        assert memory.normalize_branch_input("  session  ") == "session"
        assert memory.normalize_branch_input("my-branch") == "my-branch"


class TestGitInfrastructure:
    """Test core Git infrastructure."""

    def test_get_project_dir_creates_repo(self, tmp_path, monkeypatch):
        """Test get_project_dir creates Git repo if not exists."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        assert project_dir.exists()
        assert (project_dir / ".git").exists()

    def test_get_project_dir_idempotent(self, tmp_path, monkeypatch):
        """Test get_project_dir is idempotent."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        dir1 = memory.get_project_dir()
        dir2 = memory.get_project_dir()

        assert dir1 == dir2

    def test_read_branch_nonexistent_returns_empty(self, tmp_path, monkeypatch):
        """Test reading non-existent branch returns empty string."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()
        content = memory.read_branch(project_dir, "nonexistent")

        assert content == ""

    def test_write_and_read_branch(self, tmp_path, monkeypatch):
        """Test writing and reading from a branch."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()
        test_content = (
            '{"v": 1, "type": "message", "role": "user", "content": "hello"}\n'
        )

        result = memory.write_conversation(
            project_dir, "test-branch", test_content, "Test commit"
        )

        assert result["branch"] == "test-branch"
        assert "commit_sha" in result

        content = memory.read_branch(project_dir, "test-branch")
        assert content == test_content


class TestJSONLHelpers:
    """Test JSONL parsing and formatting helpers."""

    def test_parse_messages_empty(self):
        """Test parsing empty content."""
        events = memory.parse_messages("")
        assert events == []

    def test_parse_messages_valid(self):
        """Test parsing valid JSONL."""
        content = (
            '{"v": 1, "type": "message", "role": "user", "content": "hello"}\n'
            '{"v": 1, "type": "message", "role": "assistant", "content": "hi"}\n'
        )
        events = memory.parse_messages(content)

        assert len(events) == 2
        assert events[0]["role"] == "user"
        assert events[1]["role"] == "assistant"

    def test_format_messages(self):
        """Test formatting events to JSONL."""
        events = [
            {"v": 1, "type": "message", "role": "user", "content": "hello"},
            {"v": 1, "type": "message", "role": "assistant", "content": "hi"},
        ]
        content = memory.format_messages(events)

        # Should be two JSON lines
        lines = content.strip().split("\n")
        assert len(lines) == 2

        # Verify they're valid JSON
        for line in lines:
            json.loads(line)

    def test_append_message(self):
        """Test appending a message to content."""
        content = ""
        content = memory.append_message(content, "user", "hello")

        events = memory.parse_messages(content)
        assert len(events) == 1
        assert events[0]["role"] == "user"
        assert events[0]["content"] == "hello"
        assert "timestamp" in events[0]

    def test_append_message_with_usage(self):
        """Test appending a message with usage metadata."""
        content = ""
        usage = {"input_tokens": 10, "output_tokens": 20, "cost": 0.01}
        content = memory.append_message(content, "assistant", "response", usage=usage)

        events = memory.parse_messages(content)
        assert events[0]["usage"] == usage

    def test_validate_jsonl_valid(self):
        """Test validation passes for valid JSONL."""
        content = '{"v": 1, "type": "message", "role": "user", "content": "hello"}\n'
        memory.validate_jsonl(content)  # Should not raise

    def test_validate_jsonl_invalid_json(self):
        """Test validation fails for invalid JSON."""
        content = "not valid json\n"
        with pytest.raises(ValueError, match="Invalid JSON"):
            memory.validate_jsonl(content)

    def test_validate_jsonl_missing_version(self):
        """Test validation fails for missing version field."""
        content = '{"type": "message"}\n'
        with pytest.raises(ValueError, match="Missing 'v' field"):
            memory.validate_jsonl(content)

    def test_validate_jsonl_invalid_type(self):
        """Test validation fails for invalid event type."""
        content = '{"v": 1, "type": "invalid"}\n'
        with pytest.raises(ValueError, match="Invalid type"):
            memory.validate_jsonl(content)


class TestSystemPromptAndClear:
    """Test system prompt and clear event handling."""

    def test_append_system_prompt(self):
        """Test appending a system prompt event."""
        content = ""
        content = memory.append_system_prompt(content, "You are helpful.")

        events = memory.parse_messages(content)
        assert len(events) == 1
        assert events[0]["type"] == "system_prompt"
        assert events[0]["content"] == "You are helpful."

    def test_append_clear(self):
        """Test appending a clear event."""
        content = '{"v": 1, "type": "message", "role": "user", "content": "old"}\n'
        content = memory.append_clear(content)

        events = memory.parse_messages(content)
        assert len(events) == 2
        assert events[1]["type"] == "clear"

    def test_get_llm_context_respects_clear(self, tmp_path, monkeypatch):
        """Test that get_llm_context only returns messages after last clear."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Build content with clear in the middle
        content = ""
        content = memory.append_message(content, "user", "old message")
        content = memory.append_message(content, "assistant", "old response")
        content = memory.append_clear(content)
        content = memory.append_message(content, "user", "new message")
        content = memory.append_message(content, "assistant", "new response")

        memory.write_conversation(project_dir, "test-branch", content, "Test")

        history, system_prompt = memory.get_llm_context(project_dir, "test-branch")

        # Should only have messages after clear
        assert len(history) == 2
        assert history[0]["content"] == "new message"
        assert history[1]["content"] == "new response"

    def test_get_llm_context_returns_system_prompt(self, tmp_path, monkeypatch):
        """Test that get_llm_context returns system prompt after last clear."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        content = ""
        content = memory.append_system_prompt(content, "Old prompt")
        content = memory.append_clear(content)
        content = memory.append_system_prompt(content, "New prompt")
        content = memory.append_message(content, "user", "hello")

        memory.write_conversation(project_dir, "test-branch", content, "Test")

        history, system_prompt = memory.get_llm_context(project_dir, "test-branch")

        assert system_prompt == "New prompt"


class TestBranchOperations:
    """Test branch-level operations."""

    def test_branch_exists(self, tmp_path, monkeypatch):
        """Test checking if branch exists."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        assert not memory.branch_exists(project_dir, "test-branch")

        content = memory.append_message("", "user", "hello")
        memory.write_conversation(project_dir, "test-branch", content, "Initial")

        assert memory.branch_exists(project_dir, "test-branch")

    def test_list_branches(self, tmp_path, monkeypatch):
        """Test listing all branches."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Initially empty
        branches = memory.list_branches(project_dir)
        assert len(branches) == 0

        # Create branches via write_conversation
        content1 = memory.append_message("", "user", "hello")
        memory.write_conversation(project_dir, "branch1", content1, "Initial")
        content2 = memory.append_message("", "user", "world")
        memory.write_conversation(project_dir, "branch2", content2, "Initial")

        branches = memory.list_branches(project_dir)
        names = [b["name"] for b in branches]

        assert len(names) == 2
        assert "branch1" in names
        assert "branch2" in names

    def test_get_branch_sha(self, tmp_path, monkeypatch):
        """Test getting branch SHA."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()
        content = memory.append_message("", "user", "hello")
        memory.write_conversation(project_dir, "test-branch", content, "Initial")

        sha = memory.get_branch_sha(project_dir, "test-branch")

        assert sha is not None
        assert len(sha) == 40  # Full SHA

    def test_fork_branch(self, tmp_path, monkeypatch):
        """Test forking a branch from a ref."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Create original branch with content
        content = memory.append_message("", "user", "hello")
        result = memory.write_conversation(project_dir, "original", content, "Initial")
        original_sha = result["commit_sha"]

        # Fork from that commit
        memory.fork_branch(project_dir, "forked", original_sha)

        # Both branches should have same content
        original_content = memory.read_branch(project_dir, "original")
        forked_content = memory.read_branch(project_dir, "forked")

        assert original_content == forked_content


class TestForkOnConflict:
    """Test fork-on-conflict concurrency handling."""

    def test_write_returns_commit_sha(self, tmp_path, monkeypatch):
        """Test write_conversation returns commit SHA."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()
        content = memory.append_message("", "user", "hello")

        result = memory.write_conversation(project_dir, "test", content, "Test")

        assert "commit_sha" in result
        assert len(result["commit_sha"]) == 40

    def test_concurrent_write_forks(self, tmp_path, monkeypatch):
        """Test that concurrent writes fork instead of losing data."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Create initial branch
        content1 = memory.append_message("", "user", "message 1")
        result1 = memory.write_conversation(project_dir, "main", content1, "First")
        original_sha = result1["commit_sha"]

        # Simulate concurrent write by creating commit with original SHA as parent
        # then manually advancing branch before fast-forward
        content2 = memory.append_message(content1, "assistant", "response 1")
        commit2 = memory.create_commit(project_dir, content2, original_sha, "Second")

        # Manually advance the branch to simulate another writer
        content3 = memory.append_message(content1, "assistant", "different response")
        memory.write_conversation(project_dir, "main", content3, "Concurrent")

        # Now try fast-forward which should fail
        success = memory.try_fast_forward(project_dir, "main", original_sha, commit2)

        assert not success  # Fast-forward should fail

        # Create fork branch
        fork_name = f"main-fork-{commit2}"
        memory.create_branch(project_dir, fork_name, commit2)

        # Verify fork has our content
        fork_content = memory.read_branch(project_dir, fork_name)
        assert "response 1" in fork_content


class TestGetResponse:
    """Test get_response functionality."""

    def test_get_response(self, tmp_path, monkeypatch):
        """Test getting response by index."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        content = ""
        content = memory.append_message(content, "user", "Q1")
        content = memory.append_message(content, "assistant", "A1")
        content = memory.append_message(content, "user", "Q2")
        content = memory.append_message(content, "assistant", "A2")
        memory.write_conversation(project_dir, "test-branch", content, "Add messages")

        # Last response (default)
        response = memory.get_response(project_dir, "test-branch")
        assert response["content"] == "A2"

        # First response
        response = memory.get_response(project_dir, "test-branch", 0)
        assert response["content"] == "A1"

        # Second-to-last
        response = memory.get_response(project_dir, "test-branch", -2)
        assert response["content"] == "A1"

    def test_get_response_empty_raises(self, tmp_path, monkeypatch):
        """Test get_response raises when no assistant messages."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Create branch with only user messages
        content = memory.append_message("", "user", "hello")
        memory.write_conversation(project_dir, "test-branch", content, "Initial")

        with pytest.raises(IndexError, match="no assistant responses"):
            memory.get_response(project_dir, "test-branch")


class TestWorktreeEditing:
    """Test worktree-based editing."""

    def test_is_locked_when_not_locked(self, tmp_path, monkeypatch):
        """Test is_locked returns None when not locked."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()
        lock_info = memory.is_locked(project_dir)

        assert lock_info is None

    def test_start_and_end_edit(self, tmp_path, monkeypatch):
        """Test starting and ending an edit session."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Start edit
        result = memory.start_edit(project_dir)
        worktree_path = result["path"]

        assert Path(worktree_path).exists()
        assert memory.is_locked(project_dir) is not None

        # End edit
        memory.end_edit(project_dir)

        assert not Path(worktree_path).exists()
        assert memory.is_locked(project_dir) is None

    def test_start_edit_when_locked_raises(self, tmp_path, monkeypatch):
        """Test start_edit raises when already locked."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        memory.start_edit(project_dir)

        with pytest.raises(ValueError, match="Editing already in progress"):
            memory.start_edit(project_dir)

        # Cleanup
        memory.end_edit(project_dir, force=True)

    def test_force_end_edit(self, tmp_path, monkeypatch):
        """Test force ending edit session."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        result = memory.start_edit(project_dir)
        worktree_path = result["path"]

        # Force end (simulating recovery from crash)
        memory.end_edit(project_dir, force=True)

        assert not Path(worktree_path).exists()
        assert memory.is_locked(project_dir) is None


class TestGetLog:
    """Test commit log retrieval."""

    def test_get_log(self, tmp_path, monkeypatch):
        """Test getting commit log for a branch."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Create branch with multiple commits
        content = memory.append_message("", "user", "first")
        memory.write_conversation(project_dir, "test", content, "First commit")

        content = memory.append_message(content, "assistant", "response")
        memory.write_conversation(project_dir, "test", content, "Second commit")

        log = memory.get_log(project_dir, "test", limit=10)

        assert len(log) == 2
        assert "sha" in log[0]
        assert "timestamp" in log[0]
        assert "message_preview" in log[0]

    def test_get_log_respects_limit(self, tmp_path, monkeypatch):
        """Test that get_log respects the limit parameter."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Create branch with multiple commits
        content = ""
        for i in range(5):
            content = memory.append_message(content, "user", f"message {i}")
            memory.write_conversation(project_dir, "test", content, f"Commit {i}")

        log = memory.get_log(project_dir, "test", limit=3)

        assert len(log) == 3


class TestReadRef:
    """Test reading content at specific refs."""

    def test_read_ref_by_sha(self, tmp_path, monkeypatch):
        """Test reading content at a specific SHA."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        # Create commits
        content1 = memory.append_message("", "user", "first")
        result1 = memory.write_conversation(project_dir, "test", content1, "First")
        sha1 = result1["commit_sha"]

        content2 = memory.append_message(content1, "assistant", "second")
        memory.write_conversation(project_dir, "test", content2, "Second")

        # Read at old SHA - returns (content, resolved_sha)
        old_content, resolved_sha = memory.read_ref(project_dir, sha1)
        assert "first" in old_content
        assert "second" not in old_content
        assert resolved_sha == sha1

    def test_read_ref_nonexistent_raises(self, tmp_path, monkeypatch):
        """Test reading non-existent ref raises ValueError."""
        monkeypatch.setenv("MCP_HANDLEY_LAB_MEMORY_DIR", str(tmp_path))
        monkeypatch.chdir(tmp_path)

        project_dir = memory.get_project_dir()

        with pytest.raises(ValueError, match="Ref not found"):
            memory.read_ref(project_dir, "0000000000000000000000000000000000000000")
