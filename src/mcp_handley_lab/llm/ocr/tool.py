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
    "Returns: {pages: [{markdown, images?}], model, usage_info, output_file?}. "
    "Creates parent directories automatically."
)
def process(
    document_path: str = Field(
        ...,
        description="Path to document file or URL. Supports PDF, images, PPTX, DOCX.",
    ),
    output_file: str = Field(
        default="",
        description="File path to save full OCR results as JSON. Empty means no file output.",
    ),
    include_images: bool = Field(
        default=False,
        description="Include base64-encoded images with bounding boxes in output.",
    ),
) -> dict[str, Any]:
    """Process document with Mistral OCR for text extraction."""
    from mcp_handley_lab.llm.registry import get_adapter

    adapter = get_adapter("mistral", "ocr")
    result = adapter(document_path, include_images)

    # Return full result directly (like audio tool does)
    # Copy to avoid mutating adapter's result
    response = dict(result)

    if output_file:
        output_path = Path(output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        response["output_file"] = output_file

    return response
