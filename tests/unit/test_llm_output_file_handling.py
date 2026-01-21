"""Unit tests for output_file handling across LLM tools.

Tests that all tools create parent directories automatically when output_file is specified.
"""

import json
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest


class TestOutputFileParentDirectoryCreation:
    """Test that all LLM tools create parent directories for output_file."""

    @pytest.mark.asyncio
    async def test_ocr_creates_parent_directories(self):
        """Test that OCR tool creates parent directories for output_file."""
        from mcp_handley_lab.llm.ocr.tool import mcp

        mock_result = {
            "pages": [{"markdown": "Test OCR content"}],
            "model": "mistral-ocr-latest",
            "usage_info": {"input_tokens": 10, "output_tokens": 5},
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            nested_output = Path(tmpdir) / "nested" / "deep" / "output.json"
            assert not nested_output.parent.exists()

            with patch(
                "mcp_handley_lab.llm.registry.get_adapter", return_value=lambda *args: mock_result
            ):
                _, response = await mcp.call_tool(
                    "process",
                    {
                        "document_path": "/fake/path.pdf",
                        "output_file": str(nested_output),
                        "include_images": False,
                    },
                )

            assert nested_output.exists(), "Output file should be created"
            assert nested_output.parent.exists(), "Parent directories should be created"
            assert response["output_file"] == str(nested_output)

            # Verify file contents
            saved_data = json.loads(nested_output.read_text())
            assert "pages" in saved_data
            assert saved_data["model"] == "mistral-ocr-latest"

    @pytest.mark.asyncio
    async def test_audio_creates_parent_directories(self):
        """Test that audio tool creates parent directories for output_file."""
        from mcp_handley_lab.llm.audio.tool import mcp

        mock_result = {
            "text": "Transcribed audio content",
            "model": "voxtral-latest",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            nested_output = Path(tmpdir) / "nested" / "deep" / "transcript.json"
            assert not nested_output.parent.exists()

            with patch(
                "mcp_handley_lab.llm.registry.get_adapter", return_value=lambda **kwargs: mock_result
            ):
                _, response = await mcp.call_tool(
                    "transcribe",
                    {
                        "audio_path": "/fake/audio.wav",
                        "output_file": str(nested_output),
                    },
                )

            assert nested_output.exists(), "Output file should be created"
            assert nested_output.parent.exists(), "Parent directories should be created"
            assert response["output_file"] == str(nested_output)

            # Verify file contents
            saved_data = json.loads(nested_output.read_text())
            assert saved_data["text"] == "Transcribed audio content"

    @pytest.mark.asyncio
    async def test_embeddings_get_embeddings_creates_parent_directories(self):
        """Test that get_embeddings creates parent directories for output_file."""
        from mcp_handley_lab.llm.embeddings.tool import mcp

        mock_embeddings = [[0.1, 0.2, 0.3] * 512]  # 1536 dimensions

        with tempfile.TemporaryDirectory() as tmpdir:
            nested_output = Path(tmpdir) / "nested" / "deep" / "embeddings.json"
            assert not nested_output.parent.exists()

            with patch(
                "mcp_handley_lab.llm.embeddings.tool._get_embeddings",
                return_value=mock_embeddings,
            ):
                _, result = await mcp.call_tool(
                    "get_embeddings",
                    {
                        "texts": ["Test text"],
                        "model": "text-embedding-3-small",
                        "output_file": str(nested_output),
                    },
                )

            assert nested_output.exists(), "Output file should be created"
            assert nested_output.parent.exists(), "Parent directories should be created"
            assert result["output_file"] == str(nested_output)

            # Verify file contents
            saved_data = json.loads(nested_output.read_text())
            assert "embeddings" in saved_data
            assert len(saved_data["embeddings"]) == 1

    @pytest.mark.asyncio
    async def test_embeddings_index_documents_creates_parent_directories(self):
        """Test that index_documents creates parent directories for output_index_path."""
        from mcp_handley_lab.llm.embeddings.tool import mcp

        mock_embeddings = [[0.1, 0.2, 0.3] * 512]  # 1536 dimensions

        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a test document
            doc_path = Path(tmpdir) / "test_doc.txt"
            doc_path.write_text("This is a test document.")

            nested_index = Path(tmpdir) / "nested" / "deep" / "index.json"
            assert not nested_index.parent.exists()

            with patch(
                "mcp_handley_lab.llm.embeddings.tool._get_embeddings",
                return_value=mock_embeddings,
            ):
                _, result = await mcp.call_tool(
                    "index_documents",
                    {
                        "document_paths": [str(doc_path)],
                        "output_index_path": str(nested_index),
                        "model": "text-embedding-3-small",
                    },
                )

            assert nested_index.exists(), "Index file should be created"
            assert nested_index.parent.exists(), "Parent directories should be created"
            assert result["index_path"] == str(nested_index)

            # Verify file contents
            saved_index = json.loads(nested_index.read_text())
            assert "documents" in saved_index
            assert len(saved_index["documents"]) == 1

    @pytest.mark.asyncio
    async def test_chat_creates_parent_directories(self):
        """Test that chat tool creates parent directories for output_file."""
        from mcp_handley_lab.llm.chat.tool import mcp

        mock_response = {
            "text": "Hello, world!",
            "input_tokens": 10,
            "output_tokens": 5,
            "finish_reason": "stop",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            nested_output = Path(tmpdir) / "nested" / "deep" / "response.txt"
            assert not nested_output.parent.exists()

            with patch(
                "mcp_handley_lab.llm.chat.tool.resolve_model",
                return_value=("openai", "gpt-4o-mini", {}),
            ), patch(
                "mcp_handley_lab.llm.chat.tool.validate_options"
            ), patch(
                "mcp_handley_lab.llm.chat.tool.get_adapter",
                return_value=lambda **kwargs: mock_response,
            ):
                _, result = await mcp.call_tool(
                    "ask",
                    {
                        "prompt": "Hello",
                        "output_file": str(nested_output),
                        "model": "gpt-4o-mini",
                        "agent_name": "",  # Disable memory
                    },
                )

            assert nested_output.exists(), "Output file should be created"
            assert nested_output.parent.exists(), "Parent directories should be created"

            # Verify file contents
            saved_content = nested_output.read_text()
            assert saved_content == "Hello, world!"


class TestOutputFileInResponse:
    """Test that tools include output_file in response when file is written."""

    @pytest.mark.asyncio
    async def test_ocr_includes_output_file_in_response(self):
        """Test OCR includes output_file in response."""
        from mcp_handley_lab.llm.ocr.tool import mcp

        mock_result = {
            "pages": [{"markdown": "Test"}],
            "model": "mistral-ocr-latest",
            "usage_info": {},
        }

        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            output_path = f.name

        try:
            with patch(
                "mcp_handley_lab.llm.registry.get_adapter",
                return_value=lambda *args: mock_result,
            ):
                _, response = await mcp.call_tool(
                    "process",
                    {
                        "document_path": "/fake/path.pdf",
                        "output_file": output_path,
                    },
                )

            assert "output_file" in response
            assert response["output_file"] == output_path
        finally:
            Path(output_path).unlink(missing_ok=True)

    @pytest.mark.asyncio
    async def test_audio_includes_output_file_in_response(self):
        """Test audio includes output_file in response."""
        from mcp_handley_lab.llm.audio.tool import mcp

        mock_result = {"text": "Transcribed", "model": "voxtral-latest"}

        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            output_path = f.name

        try:
            with patch(
                "mcp_handley_lab.llm.registry.get_adapter",
                return_value=lambda **kwargs: mock_result,
            ):
                _, response = await mcp.call_tool(
                    "transcribe",
                    {
                        "audio_path": "/fake/audio.wav",
                        "output_file": output_path,
                    },
                )

            assert "output_file" in response
            assert response["output_file"] == output_path
        finally:
            Path(output_path).unlink(missing_ok=True)

    @pytest.mark.asyncio
    async def test_embeddings_includes_output_file_in_response(self):
        """Test embeddings includes output_file in response."""
        from mcp_handley_lab.llm.embeddings.tool import mcp

        mock_embeddings = [[0.1, 0.2, 0.3]]

        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            output_path = f.name

        try:
            with patch(
                "mcp_handley_lab.llm.embeddings.tool._get_embeddings",
                return_value=mock_embeddings,
            ):
                _, result = await mcp.call_tool(
                    "get_embeddings",
                    {
                        "texts": ["Test"],
                        "model": "text-embedding-3-small",
                        "output_file": output_path,
                    },
                )

            assert "output_file" in result
            assert result["output_file"] == output_path
        finally:
            Path(output_path).unlink(missing_ok=True)
