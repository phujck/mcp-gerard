"""Apple Pages document MCP tool."""

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.pages.applescript import PagesAppleScript
from mcp_handley_lab.pages.models import (
    ConversionResult,
    CreateDocumentResult,
    DocumentAnalysisResult,
    DocumentInfoResult,
    ExportResult,
    FileListResult,
    FindReplaceResult,
    InsertTextResult,
    SetTextResult,
    TextExtractionResult,
    WordCountResult,
)
from mcp_handley_lab.pages.parser import PagesParser
from mcp_handley_lab.shared.models import ServerInfo

mcp = FastMCP("Pages Tool")


# ============================================================================
# Read-only tools (IWA parser based)
# ============================================================================


@mcp.tool(
    description="List all files contained within a Pages document archive. "
    "Pages files are ZIP archives containing IWA protobuf data, images, and metadata."
)
def list_files(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
) -> FileListResult:
    """List files in a Pages archive."""
    parser = PagesParser(file_path)
    files = parser.list_files()

    return FileListResult(
        success=True,
        file_path=file_path,
        files=files,
        file_count=len(files),
        message=f"Found {len(files)} files in Pages archive",
    )


@mcp.tool(
    description="Extract all text content from a Pages document. "
    "Parses the IWA protobuf format to extract document text."
)
def extract_text(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
) -> TextExtractionResult:
    """Extract text from a Pages document."""
    parser = PagesParser(file_path)
    text = parser.extract_text()

    word_count = len(text.split())
    paragraph_count = text.count("\n\n") + 1 if text else 0

    return TextExtractionResult(
        success=True,
        file_path=file_path,
        text=text,
        word_count=word_count,
        paragraph_count=paragraph_count,
        message=f"Extracted {word_count} words in {paragraph_count} paragraphs",
    )


@mcp.tool(
    description="Analyze a Pages document comprehensively. "
    "Returns metadata, text content, equation count, and structure information."
)
def analyze_document(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
) -> DocumentAnalysisResult:
    """Analyze a Pages document."""
    parser = PagesParser(file_path)

    metadata = parser.get_metadata()
    text = parser.extract_text()
    word_count = len(text.split())
    paragraph_count = text.count("\n\n") + 1 if text else 0
    equation_count = parser.count_equations()
    has_preview = parser.has_preview()
    files = parser.list_files()

    return DocumentAnalysisResult(
        success=True,
        file_path=file_path,
        metadata=metadata,
        text=text,
        word_count=word_count,
        paragraph_count=paragraph_count,
        has_equations=equation_count > 0,
        equation_count=equation_count,
        has_preview=has_preview,
        file_count=len(files),
        message=f"Document: {word_count} words, {equation_count} equations, {len(files)} files",
    )


@mcp.tool(
    description="Convert a Pages document to plain text format using the parser. "
    "Extracts and saves the text content to a file."
)
def convert_to_text(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    output_path: str = Field(..., description="Path to save the text output."),
) -> ConversionResult:
    """Convert Pages to plain text."""
    parser = PagesParser(file_path)
    text = parser.extract_text()

    from pathlib import Path

    output = Path(output_path)
    output.write_text(text, encoding="utf-8")

    return ConversionResult(
        success=True,
        input_path=file_path,
        output_path=str(output),
        format="text",
        message=f"Converted to text: {len(text.split())} words",
    )


@mcp.tool(
    description="Extract the preview image from a Pages document. "
    "Saves the preview thumbnail as a JPEG file."
)
def extract_preview(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    output_path: str = Field(..., description="Path to save the preview image."),
) -> ConversionResult:
    """Extract preview image from Pages document."""
    parser = PagesParser(file_path)
    output = parser.extract_preview(output_path)

    return ConversionResult(
        success=True,
        input_path=file_path,
        output_path=output,
        format="jpeg",
        message=f"Extracted preview to {output}",
    )


# ============================================================================
# Editing tools (AppleScript based - requires Pages app)
# ============================================================================


@mcp.tool(
    description="Create a new Pages document. Requires the Pages app to be installed. "
    "Optionally set initial text content."
)
def create_document(
    file_path: str = Field(..., description="Path to save the new document."),
    template: str = Field(
        default="Blank",
        description="Pages template name (e.g., 'Blank', 'Essay').",
    ),
    initial_text: str = Field(
        default="",
        description="Optional initial text content for the document.",
    ),
) -> CreateDocumentResult:
    """Create a new Pages document."""
    result = PagesAppleScript.create_document(file_path, template, initial_text)
    return CreateDocumentResult(
        success=result["success"],
        file_path=result["file_path"],
        message=f"Created new Pages document at {result['file_path']}",
    )


@mcp.tool(
    description="Get the body text of a Pages document using AppleScript. "
    "Returns the full text content as seen in Pages."
)
def get_body_text(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
) -> TextExtractionResult:
    """Get body text via AppleScript."""
    result = PagesAppleScript.get_body_text(file_path)
    text = result["text"]
    word_count = len(text.split())
    paragraph_count = text.count("\n") + 1 if text else 0

    return TextExtractionResult(
        success=True,
        file_path=file_path,
        text=text,
        word_count=word_count,
        paragraph_count=paragraph_count,
        message=f"Retrieved {word_count} words via AppleScript",
    )


@mcp.tool(
    description="Replace the entire body text of a Pages document. "
    "WARNING: This replaces ALL text in the document."
)
def set_body_text(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    text: str = Field(..., description="New text content for the document."),
) -> SetTextResult:
    """Set the entire body text."""
    result = PagesAppleScript.set_body_text(file_path, text)
    return SetTextResult(
        success=result["success"],
        file_path=result["file_path"],
        message=f"Replaced document text with {len(text.split())} words",
    )


@mcp.tool(
    description="Find and replace text throughout a Pages document. "
    "Replaces all occurrences of the search text."
)
def find_replace(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    find_text: str = Field(..., description="Text to find."),
    replace_text: str = Field(..., description="Text to replace with."),
) -> FindReplaceResult:
    """Find and replace text in document."""
    result = PagesAppleScript.find_replace(file_path, find_text, replace_text)
    return FindReplaceResult(
        success=result["success"],
        file_path=result["file_path"],
        find_text=result["find_text"],
        replace_text=result["replace_text"],
        replacement_count=result["replacement_count"],
        message=f"Replaced {result['replacement_count']} occurrences of '{find_text}'",
    )


@mcp.tool(
    description="Append text to the end of a Pages document. "
    "Adds the text after all existing content."
)
def append_text(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    text: str = Field(..., description="Text to append."),
) -> SetTextResult:
    """Append text to document."""
    result = PagesAppleScript.append_text(file_path, text)
    return SetTextResult(
        success=result["success"],
        file_path=result["file_path"],
        message=f"Appended {len(text.split())} words to document",
    )


@mcp.tool(
    description="Insert text at a specific character position in a Pages document. "
    "Position is 1-indexed; use 0 to append at end."
)
def insert_text(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    text: str = Field(..., description="Text to insert."),
    position: int = Field(
        default=0,
        description="Character position (1-indexed, 0 = end of document).",
    ),
) -> InsertTextResult:
    """Insert text at position."""
    result = PagesAppleScript.insert_text(file_path, text, position)
    return InsertTextResult(
        success=result["success"],
        file_path=result["file_path"],
        position=result["position"],
        message=f"Inserted text at position {position}",
    )


@mcp.tool(
    description="Get word count and other statistics for a Pages document. "
    "Returns word count, character count, and paragraph count."
)
def get_word_count(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
) -> WordCountResult:
    """Get document statistics."""
    result = PagesAppleScript.get_word_count(file_path)
    return WordCountResult(
        success=result["success"],
        file_path=result["file_path"],
        word_count=result["word_count"],
        character_count=result["character_count"],
        paragraph_count=result["paragraph_count"],
        message=f"{result['word_count']} words, {result['character_count']} chars, {result['paragraph_count']} paragraphs",
    )


@mcp.tool(
    description="Export a Pages document to another format. "
    "Supported formats: PDF, Word (docx), text, EPUB."
)
def export_document(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
    output_path: str = Field(..., description="Path for the exported file."),
    format: str = Field(
        default="PDF",
        description="Export format: PDF, Word, text, or EPUB.",
    ),
) -> ExportResult:
    """Export document to another format."""
    result = PagesAppleScript.export_document(file_path, output_path, format)
    return ExportResult(
        success=result["success"],
        input_path=result["input_path"],
        output_path=result["output_path"],
        format=result["format"],
        message=f"Exported to {format}: {result['output_path']}",
    )


@mcp.tool(
    description="Get document information and properties. "
    "Returns name, modification status, and page count."
)
def get_document_info(
    file_path: str = Field(..., description="Path to the Pages (.pages) file."),
) -> DocumentInfoResult:
    """Get document info."""
    result = PagesAppleScript.get_document_info(file_path)
    return DocumentInfoResult(
        success=result["success"],
        file_path=result["file_path"],
        name=result["name"],
        modified=result["modified"],
        page_count=result["page_count"],
        message=f"{result['name']}: {result['page_count']} pages",
    )


@mcp.tool(
    description="Get server information including available functions and format support."
)
def server_info() -> ServerInfo:
    """Get Pages Tool server information."""
    dependencies = {
        "snappy": "python-snappy (for IWA decompression)",
        "applescript": "Pages app (for editing operations)",
        "supported_formats": ".pages (Pages 13.x and later)",
    }

    available_functions = [
        # Read-only (parser)
        "list_files",
        "extract_text",
        "analyze_document",
        "convert_to_text",
        "extract_preview",
        # Editing (AppleScript)
        "create_document",
        "get_body_text",
        "set_body_text",
        "find_replace",
        "append_text",
        "insert_text",
        "get_word_count",
        "export_document",
        "get_document_info",
        "server_info",
    ]

    return ServerInfo(
        name="Pages Tool",
        version="1.0.0",
        status="active",
        capabilities=available_functions,
        dependencies=dependencies,
    )
