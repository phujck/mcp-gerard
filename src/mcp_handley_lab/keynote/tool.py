"""Keynote presentation MCP tool."""

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.keynote.applescript import KeynoteAppleScript
from mcp_handley_lab.keynote.models import (
    AddSlideResult,
    CreatePresentationResult,
    DeleteSlideResult,
    DuplicateSlideResult,
    FileListResult,
    PackResult,
    PresentationInfo,
    ReplaceResult,
    SetSlideContentResult,
    SlideInfo,
    SlideLayoutsResult,
    TextExtractionResult,
    UnpackResult,
)
from mcp_handley_lab.keynote.parser import KeynoteParser
from mcp_handley_lab.shared.models import ServerInfo

mcp = FastMCP("Keynote Tool")


@mcp.tool(
    description="List all files contained within a Keynote presentation archive. "
    "Keynote files are ZIP archives containing YAML, images, and other resources."
)
def list_files(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
) -> FileListResult:
    """List files in a Keynote archive."""
    parser = KeynoteParser(file_path)
    files = parser.list_files()

    return FileListResult(
        success=True,
        file_path=file_path,
        files=files,
        file_count=len(files),
        message=f"Found {len(files)} files in Keynote archive",
    )


@mcp.tool(
    description="Get an overview of a Keynote presentation including slide count "
    "and basic information about each slide."
)
def get_presentation_info(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
) -> PresentationInfo:
    """Get presentation overview."""
    parser = KeynoteParser(file_path)
    slides_data = parser.get_all_slides_text()

    slides = [
        SlideInfo(
            number=s["number"],
            title=s["title"],
            text_content=s["text_content"],
            has_notes=s["has_notes"],
            notes=s["notes"],
        )
        for s in slides_data
    ]

    return PresentationInfo(
        success=True,
        file_path=file_path,
        slide_count=len(slides),
        slides=slides,
        message=f"Presentation has {len(slides)} slides",
    )


@mcp.tool(
    description="Extract all text content from a Keynote presentation, "
    "including slide titles, body text, and presenter notes."
)
def extract_text(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
) -> TextExtractionResult:
    """Extract all text from presentation."""
    parser = KeynoteParser(file_path)
    slides_data = parser.get_all_slides_text()

    slides = [
        SlideInfo(
            number=s["number"],
            title=s["title"],
            text_content=s["text_content"],
            has_notes=s["has_notes"],
            notes=s["notes"],
        )
        for s in slides_data
    ]

    # Compile full text
    all_text_parts = []
    for slide in slides:
        if slide.title:
            all_text_parts.append(slide.title)
        all_text_parts.extend(slide.text_content)
        if slide.notes:
            all_text_parts.append(slide.notes)

    full_text = "\n\n".join(all_text_parts)
    word_count = len(full_text.split())

    return TextExtractionResult(
        success=True,
        file_path=file_path,
        slides=slides,
        full_text=full_text,
        word_count=word_count,
        message=f"Extracted {word_count} words from {len(slides)} slides",
    )


@mcp.tool(
    description="Find and replace text throughout a Keynote presentation. "
    "Optionally save to a new file."
)
def find_replace(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
    find_text: str = Field(..., description="Text to find."),
    replace_text: str = Field(..., description="Text to replace with."),
    output_path: str = Field(
        default="",
        description="Output file path. If empty, modifies the original file.",
    ),
) -> ReplaceResult:
    """Find and replace text in presentation."""
    parser = KeynoteParser(file_path)
    output, count = parser.find_replace(
        find_text, replace_text, output_path if output_path else None
    )

    return ReplaceResult(
        success=True,
        file_path=file_path,
        output_path=output,
        find_text=find_text,
        replace_text=replace_text,
        replacement_count=count,
        message=f"Replaced {count} occurrences of '{find_text}' with '{replace_text}'",
    )


@mcp.tool(
    description="Unpack a Keynote file into a directory of editable YAML and media files. "
    "Useful for version control, bulk editing, or detailed inspection."
)
def unpack(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
    output_dir: str = Field(
        default="",
        description="Output directory. If empty, uses the filename without .key extension.",
    ),
) -> UnpackResult:
    """Unpack Keynote file to directory."""
    parser = KeynoteParser(file_path)
    output, file_count = parser.unpack(output_dir if output_dir else None)

    return UnpackResult(
        success=True,
        input_path=file_path,
        output_path=output,
        file_count=file_count,
        message=f"Unpacked {file_count} files to {output}",
    )


@mcp.tool(
    description="Pack a directory of unpacked Keynote files back into a .key file. "
    "Use after editing unpacked YAML files."
)
def pack(
    input_dir: str = Field(..., description="Path to the unpacked Keynote directory."),
    output_path: str = Field(
        default="",
        description="Output .key file path. If empty, uses directory name + .key.",
    ),
) -> PackResult:
    """Pack directory into Keynote file."""
    output = KeynoteParser.pack(input_dir, output_path if output_path else None)

    return PackResult(
        success=True,
        input_path=input_dir,
        output_path=output,
        message=f"Packed directory into {output}",
    )


# AppleScript-based tools for structural operations


@mcp.tool(
    description="Create a new Keynote presentation using AppleScript. "
    "Optionally set the title slide content."
)
def create_presentation(
    file_path: str = Field(..., description="Path to save the new presentation."),
    theme: str = Field(
        default="Basic White",
        description="Keynote theme name (e.g., 'Basic White', 'Gradient').",
    ),
    title: str = Field(default="", description="Optional title for the first slide."),
    subtitle: str = Field(
        default="", description="Optional subtitle/body for the first slide."
    ),
) -> CreatePresentationResult:
    """Create a new Keynote presentation."""
    result = KeynoteAppleScript.create_presentation(file_path, theme, title, subtitle)
    return CreatePresentationResult(
        success=result["success"],
        file_path=result["file_path"],
        slide_count=result["slide_count"],
        message=f"Created presentation with {result['slide_count']} slide(s)",
    )


@mcp.tool(
    description="Add a new slide to an existing Keynote presentation using AppleScript."
)
def add_slide(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
    title: str = Field(default="", description="Slide title."),
    body: str = Field(default="", description="Slide body text."),
    slide_layout: str = Field(
        default="Title & Bullets",
        description="Layout name (e.g., 'Title & Bullets', 'Title - Center', 'Blank').",
    ),
    position: int = Field(
        default=0,
        description="Position to insert (0 = end, 1 = first, etc.).",
    ),
) -> AddSlideResult:
    """Add a slide to a presentation."""
    result = KeynoteAppleScript.add_slide(file_path, title, body, slide_layout, position)
    return AddSlideResult(
        success=result["success"],
        file_path=result["file_path"],
        slide_number=result["slide_number"],
        total_slides=result["total_slides"],
        message=f"Added slide {result['slide_number']} of {result['total_slides']}",
    )


@mcp.tool(
    description="Set content of an existing slide in a Keynote presentation using AppleScript. "
    "Only provided fields are updated; others remain unchanged."
)
def set_slide_content(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
    slide_number: int = Field(..., description="Slide number (1-indexed)."),
    title: str = Field(default="", description="New title (empty to keep existing)."),
    body: str = Field(default="", description="New body text (empty to keep existing)."),
    notes: str = Field(
        default="", description="New presenter notes (empty to keep existing)."
    ),
) -> SetSlideContentResult:
    """Set slide content."""
    # Convert empty strings to None to preserve existing content
    result = KeynoteAppleScript.set_slide_content(
        file_path,
        slide_number,
        title if title else None,
        body if body else None,
        notes if notes else None,
    )
    return SetSlideContentResult(
        success=result["success"],
        file_path=result["file_path"],
        slide_number=result["slide_number"],
        message=f"Updated slide {slide_number}",
    )


@mcp.tool(
    description="Delete a slide from a Keynote presentation using AppleScript."
)
def delete_slide(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
    slide_number: int = Field(..., description="Slide number to delete (1-indexed)."),
) -> DeleteSlideResult:
    """Delete a slide from presentation."""
    result = KeynoteAppleScript.delete_slide(file_path, slide_number)
    return DeleteSlideResult(
        success=result["success"],
        file_path=result["file_path"],
        remaining_slides=result["remaining_slides"],
        message=f"Deleted slide {slide_number}, {result['remaining_slides']} slides remaining",
    )


@mcp.tool(
    description="Duplicate a slide in a Keynote presentation using AppleScript."
)
def duplicate_slide(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
    slide_number: int = Field(..., description="Slide number to duplicate (1-indexed)."),
) -> DuplicateSlideResult:
    """Duplicate a slide."""
    result = KeynoteAppleScript.duplicate_slide(file_path, slide_number)
    return DuplicateSlideResult(
        success=result["success"],
        file_path=result["file_path"],
        new_slide_number=result["new_slide_number"],
        total_slides=result["total_slides"],
        message=f"Duplicated slide {slide_number}, new slide is {result['new_slide_number']}",
    )


@mcp.tool(
    description="Get available slide layouts (master slides) in a Keynote presentation."
)
def get_slide_layouts(
    file_path: str = Field(..., description="Path to the Keynote (.key) file."),
) -> SlideLayoutsResult:
    """Get available slide layouts."""
    result = KeynoteAppleScript.get_slide_layouts(file_path)
    return SlideLayoutsResult(
        success=result["success"],
        file_path=result["file_path"],
        layouts=result["layouts"],
        message=f"Found {len(result['layouts'])} layouts: {', '.join(result['layouts'])}",
    )


@mcp.tool(
    description="Get server information including available functions and dependency status."
)
def server_info() -> ServerInfo:
    """Get Keynote Tool server information."""
    import subprocess

    # Check keynote-parser availability
    try:
        result = subprocess.run(
            ["keynote-parser", "--version"], capture_output=True, text=True, check=True
        )
        kp_version = result.stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError):
        kp_version = "not installed"

    dependencies = {
        "keynote-parser": kp_version,
        "supported_formats": ".key (Keynote 14.x)",
    }

    available_functions = [
        # Parser-based (keynote-parser)
        "list_files",
        "get_presentation_info",
        "extract_text",
        "find_replace",
        "unpack",
        "pack",
        # AppleScript-based (structural operations)
        "create_presentation",
        "add_slide",
        "set_slide_content",
        "delete_slide",
        "duplicate_slide",
        "get_slide_layouts",
        "server_info",
    ]

    return ServerInfo(
        name="Keynote Tool",
        version="1.0.0",
        status="active",
        capabilities=available_functions,
        dependencies=dependencies,
    )
