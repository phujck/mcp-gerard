"""Unit tests for the blog tool."""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock, mock_open
import re


@pytest.mark.unit
class TestSlugGeneration:
    def test_simple_title(self):
        from mcp_gerard.blog import _slug
        assert _slug("Hello World") == "hello-world"

    def test_special_characters(self):
        from mcp_gerard.blog import _slug
        assert _slug("My Post: A Story!") == "my-post-a-story"

    def test_multiple_spaces(self):
        from mcp_gerard.blog import _slug
        assert _slug("A   B   C") == "a-b-c"

    def test_leading_trailing_hyphens_stripped(self):
        from mcp_gerard.blog import _slug
        result = _slug("  Test  ")
        assert not result.startswith("-")
        assert not result.endswith("-")

    def test_numbers_preserved(self):
        from mcp_gerard.blog import _slug
        assert _slug("Chapter 2") == "chapter-2"


@pytest.mark.unit
class TestExtractTitle:
    def test_extracts_title(self, tmp_path):
        from mcp_gerard.blog import _extract_title_from_tex
        tex = tmp_path / "main.tex"
        tex.write_text(r"\title{My Great Post}", encoding="utf-8")
        assert _extract_title_from_tex(tex) == "My Great Post"

    def test_falls_back_to_dirname(self, tmp_path):
        from mcp_gerard.blog import _extract_title_from_tex
        tex = tmp_path / "main.tex"
        tex.write_text("no title here", encoding="utf-8")
        assert _extract_title_from_tex(tex) == tmp_path.name

    def test_missing_file_returns_dirname(self, tmp_path):
        from mcp_gerard.blog import _extract_title_from_tex
        tex = tmp_path / "missing.tex"
        assert _extract_title_from_tex(tex) == tmp_path.name


@pytest.mark.unit
class TestBlogNewDraft:
    def test_creates_directory_and_tex(self, tmp_path):
        from mcp_gerard import blog
        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            result = blog.blog_new_draft("Test Post")

        assert result.endswith("main.tex")
        main_tex = Path(result)
        assert main_tex.exists()
        content = main_tex.read_text(encoding="utf-8")
        assert "Test Post" in content
        assert r"\input{preamble}" in content
        assert r"\input{macros}" in content

    def test_creates_images_directory(self, tmp_path):
        from mcp_gerard import blog
        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            blog.blog_new_draft("My Post")

        images_dir = tmp_path / "my-post" / "images"
        assert images_dir.is_dir()

    def test_slug_used_as_directory_name(self, tmp_path):
        from mcp_gerard import blog
        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            blog.blog_new_draft("Hello World!")

        assert (tmp_path / "hello-world").is_dir()


@pytest.mark.unit
class TestBlogListDrafts:
    def test_empty_directory(self, tmp_path):
        from mcp_gerard import blog
        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            result = blog.blog_list_drafts()
        assert result == []

    def test_nonexistent_directory(self, tmp_path):
        from mcp_gerard import blog
        missing = tmp_path / "nonexistent"
        with patch.object(blog, "BLOG_DRAFTS_PATH", missing):
            result = blog.blog_list_drafts()
        assert result == []

    def test_lists_drafts(self, tmp_path):
        from mcp_gerard import blog
        # Create two draft directories
        for slug in ("post-one", "post-two"):
            d = tmp_path / slug
            d.mkdir()
            (d / "main.tex").write_text(
                rf"\title{{{slug.replace('-', ' ').title()}}}", encoding="utf-8"
            )

        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            result = blog.blog_list_drafts()

        assert len(result) == 2
        slugs = {r["slug"] for r in result}
        assert slugs == {"post-one", "post-two"}

    def test_compiled_flag_true_when_draft_md_exists(self, tmp_path):
        from mcp_gerard import blog
        d = tmp_path / "compiled-post"
        d.mkdir()
        (d / "main.tex").write_text(r"\title{Test}", encoding="utf-8")
        (d / "draft.md").write_text("# Test", encoding="utf-8")

        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            result = blog.blog_list_drafts()

        assert result[0]["compiled"] is True

    def test_compiled_flag_false_without_draft_md(self, tmp_path):
        from mcp_gerard import blog
        d = tmp_path / "uncompiled-post"
        d.mkdir()
        (d / "main.tex").write_text(r"\title{Test}", encoding="utf-8")

        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            result = blog.blog_list_drafts()

        assert result[0]["compiled"] is False


@pytest.mark.unit
class TestBlogOpenDraft:
    def test_missing_draft_returns_error(self, tmp_path):
        from mcp_gerard import blog
        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path):
            result = blog.blog_open_draft("nonexistent")
        assert "Error" in result

    def test_opens_vscode(self, tmp_path):
        from mcp_gerard import blog
        draft_dir = tmp_path / "my-draft"
        draft_dir.mkdir()

        with patch.object(blog, "BLOG_DRAFTS_PATH", tmp_path), \
             patch("subprocess.Popen") as mock_popen:
            result = blog.blog_open_draft("my-draft")

        mock_popen.assert_called_once()
        assert str(draft_dir) in result
