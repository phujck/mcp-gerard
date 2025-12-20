"""OCR Tool for document text extraction via MCP.

Provides high-accuracy OCR using Mistral's OCR model.
Supports PDFs, images, PPTX, and DOCX files.
"""

import json
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

mcp = FastMCP("OCR Tool")


@mcp.tool(
    description="Extract text from documents using Mistral OCR. "
    "Supports PDFs, images (PNG, JPG), PPTX, and DOCX. "
    "Returns structured markdown with optional bounding box data."
)
def process(
    document_path: str = Field(
        ...,
        description="Path to document file or URL. Supports PDF, images, PPTX, DOCX.",
    ),
    output_file: str = Field(
        ...,
        description="File path to save OCR results as JSON.",
    ),
    include_images: bool = Field(
        default=True,
        description="Include base64-encoded images with bounding boxes.",
    ),
) -> dict[str, Any]:
    """Process document with Mistral OCR for text extraction."""
    from mcp_handley_lab.llm.registry import get_adapter

    adapter = get_adapter("mistral", "ocr")
    result = adapter(document_path, include_images)

    Path(output_file).write_text(json.dumps(result, indent=2))
    return result
