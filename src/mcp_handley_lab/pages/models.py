"""Pydantic models for Pages MCP tool."""

from pydantic import BaseModel


class FileListResult(BaseModel):
    """Result of listing files in a Pages archive."""

    success: bool
    file_path: str
    files: list[str]
    file_count: int
    message: str


class DocumentMetadata(BaseModel):
    """Metadata extracted from a Pages document."""

    document_uuid: str = ""
    file_format_version: str = ""
    is_multi_page: bool = False
    pages_version: str = ""


class TextExtractionResult(BaseModel):
    """Result of extracting text from a Pages document."""

    success: bool
    file_path: str
    text: str
    word_count: int
    paragraph_count: int
    message: str


class DocumentAnalysisResult(BaseModel):
    """Comprehensive analysis of a Pages document."""

    success: bool
    file_path: str
    metadata: DocumentMetadata
    text: str
    word_count: int
    paragraph_count: int
    has_equations: bool
    equation_count: int
    has_preview: bool
    file_count: int
    message: str


class ConversionResult(BaseModel):
    """Result of converting a Pages document."""

    success: bool
    input_path: str
    output_path: str
    format: str
    message: str


class CreateDocumentResult(BaseModel):
    """Result of creating a new Pages document."""

    success: bool
    file_path: str
    message: str


class SetTextResult(BaseModel):
    """Result of setting document text."""

    success: bool
    file_path: str
    message: str


class FindReplaceResult(BaseModel):
    """Result of find and replace operation."""

    success: bool
    file_path: str
    find_text: str
    replace_text: str
    replacement_count: int
    message: str


class InsertTextResult(BaseModel):
    """Result of inserting text."""

    success: bool
    file_path: str
    position: int
    message: str


class WordCountResult(BaseModel):
    """Word count and statistics."""

    success: bool
    file_path: str
    word_count: int
    character_count: int
    paragraph_count: int
    message: str


class ExportResult(BaseModel):
    """Result of exporting a document."""

    success: bool
    input_path: str
    output_path: str
    format: str
    message: str


class DocumentInfoResult(BaseModel):
    """Document information and properties."""

    success: bool
    file_path: str
    name: str
    modified: bool
    page_count: int
    message: str
