"""Pydantic models for Keynote tool responses."""

from pydantic import BaseModel


class SlideInfo(BaseModel):
    """Information about a single slide."""

    number: int
    title: str = ""
    text_content: list[str] = []
    has_notes: bool = False
    notes: str = ""


class PresentationInfo(BaseModel):
    """Overview information about a Keynote presentation."""

    success: bool
    file_path: str
    slide_count: int
    slides: list[SlideInfo]
    message: str


class FileListResult(BaseModel):
    """Result of listing files in a Keynote archive."""

    success: bool
    file_path: str
    files: list[str]
    file_count: int
    message: str


class TextExtractionResult(BaseModel):
    """Result of extracting text from a presentation."""

    success: bool
    file_path: str
    slides: list[SlideInfo]
    full_text: str
    word_count: int
    message: str


class ReplaceResult(BaseModel):
    """Result of a find/replace operation."""

    success: bool
    file_path: str
    output_path: str
    find_text: str
    replace_text: str
    replacement_count: int
    message: str


class UnpackResult(BaseModel):
    """Result of unpacking a Keynote file."""

    success: bool
    input_path: str
    output_path: str
    file_count: int
    message: str


class PackResult(BaseModel):
    """Result of packing a directory into a Keynote file."""

    success: bool
    input_path: str
    output_path: str
    message: str


# AppleScript operation results


class CreatePresentationResult(BaseModel):
    """Result of creating a new presentation via AppleScript."""

    success: bool
    file_path: str
    slide_count: int
    message: str


class AddSlideResult(BaseModel):
    """Result of adding a slide via AppleScript."""

    success: bool
    file_path: str
    slide_number: int
    total_slides: int
    message: str


class SetSlideContentResult(BaseModel):
    """Result of setting slide content via AppleScript."""

    success: bool
    file_path: str
    slide_number: int
    message: str


class DeleteSlideResult(BaseModel):
    """Result of deleting a slide via AppleScript."""

    success: bool
    file_path: str
    remaining_slides: int
    message: str


class DuplicateSlideResult(BaseModel):
    """Result of duplicating a slide via AppleScript."""

    success: bool
    file_path: str
    new_slide_number: int
    total_slides: int
    message: str


class SlideLayoutsResult(BaseModel):
    """Result of listing available slide layouts."""

    success: bool
    file_path: str
    layouts: list[str]
    message: str
