"""Parser for Apple Pages documents.

Pages files are ZIP archives containing IWA (iWork Archive) files.
IWA files use Apple's variant of Snappy compression with protobuf payloads.
"""

import contextlib
import plistlib
import re
import zipfile
from pathlib import Path

import snappy

from mcp_handley_lab.pages.models import DocumentMetadata


class PagesParser:
    """Parser for Apple Pages documents."""

    # Text patterns to skip (system/locale data embedded in protobuf)
    SKIP_PATTERNS = {
        # Calendar/date strings
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
        "Sun",
        "Mon",
        "Tue",
        "Wed",
        "Thu",
        "Fri",
        "Sat",
        # Locale/format strings
        "gregorian",
        "latn",
        "NaN",
        "1st quarter",
        "2nd quarter",
        "3rd quarter",
        "4th quarter",
        "Before Christ",
        "Anno Domini",
    }

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Pages file not found: {file_path}")
        if self.file_path.suffix.lower() != ".pages":
            raise ValueError(f"Not a Pages file: {file_path}")

    def list_files(self) -> list[str]:
        """List all files in the Pages archive."""
        with zipfile.ZipFile(self.file_path, "r") as zf:
            return sorted(zf.namelist())

    def get_metadata(self) -> DocumentMetadata:
        """Extract metadata from the Pages document."""
        metadata = DocumentMetadata()

        with zipfile.ZipFile(self.file_path, "r") as zf:
            # Read Properties.plist
            try:
                with zf.open("Metadata/Properties.plist") as f:
                    props = plistlib.load(f)
                    metadata.document_uuid = props.get("documentUUID", "")
                    metadata.file_format_version = props.get("fileFormatVersion", "")
                    metadata.is_multi_page = props.get("isMultiPage", False)
            except (KeyError, plistlib.InvalidFileException):
                pass

            # Read BuildVersionHistory.plist
            try:
                with zf.open("Metadata/BuildVersionHistory.plist") as f:
                    history = plistlib.load(f)
                    if history:
                        # Format: "Template: Blank (13.2)" or "M14.4-7043.0.93-4"
                        for item in history:
                            if item.startswith("M") or "(" in item:
                                metadata.pages_version = item
                                break
            except (KeyError, plistlib.InvalidFileException):
                pass

        return metadata

    def _decompress_iwa(self, data: bytes) -> bytes:
        """Decompress Apple IWA format (Snappy chunks).

        Apple's IWA format uses a custom Snappy framing:
        - Each chunk: 4-byte header (type in low nibble, length in upper 24 bits)
        - Type 0: Snappy-compressed chunk
        - Type 1: Uncompressed chunk
        """
        result = b""
        pos = 0

        while pos < len(data):
            if pos + 4 > len(data):
                break

            # Parse chunk header
            header = int.from_bytes(data[pos : pos + 4], "little")
            chunk_type = header & 0x0F
            chunk_len = header >> 8

            if chunk_len == 0:
                break

            chunk_data = data[pos + 4 : pos + 4 + chunk_len]

            if chunk_type == 0:  # Compressed
                with contextlib.suppress(Exception):
                    result += snappy.decompress(chunk_data)
            elif chunk_type == 1:  # Uncompressed
                result += chunk_data

            pos += 4 + chunk_len

        return result

    def _extract_text_from_protobuf(self, data: bytes) -> list[str]:
        """Extract text strings from protobuf data.

        Protobuf strings are stored as: wire_type=2 (length-delimited)
        We look for readable text runs that appear to be document content.
        """
        texts = []
        decoded = data.decode("utf-8", errors="ignore")

        # Find text runs - sequences of printable characters that look like content
        # Must start with uppercase letter and contain space (likely prose)
        runs = re.findall(
            r'[A-Z][A-Za-z0-9 .,;:!?\'"()\-\u2018\u2019\u201c\u201d]{14,}',
            decoded,
        )

        for run in runs:
            run = run.strip()
            # Skip system strings and non-content patterns
            if run in self.SKIP_PATTERNS:
                continue
            # Skip repeated character patterns (protobuf structure artifacts)
            if re.match(r"^([A-Za-z])\1{5,}", run):
                continue
            if re.match(r"^[iIoOjJlLQWTY]+\d*$", run):  # Binary artifacts
                continue
            if run.startswith("d MMM") or run.startswith("HH:mm"):  # Date formats
                continue
            if "#" in run and "," in run and "0" in run:  # Number formats
                continue
            # Strip leading single character if followed by uppercase (protobuf length byte)
            if len(run) > 2 and run[1].isupper() and run[0] != run[1]:
                run = run[1:]
            # Must contain at least one space (likely prose)
            if " " in run:
                texts.append(run)

        return texts

    def extract_text(self) -> str:
        """Extract all text content from the Pages document."""
        all_texts = []

        with zipfile.ZipFile(self.file_path, "r") as zf:
            # Process all IWA files in Index/
            for name in zf.namelist():
                if name.startswith("Index/") and name.endswith(".iwa"):
                    try:
                        with zf.open(name) as f:
                            raw_data = f.read()
                            decompressed = self._decompress_iwa(raw_data)
                            texts = self._extract_text_from_protobuf(decompressed)
                            all_texts.extend(texts)
                    except Exception:
                        continue

        # Remove duplicates while preserving order
        seen = set()
        unique_texts = []
        for text in all_texts:
            if text not in seen:
                seen.add(text)
                unique_texts.append(text)

        return "\n\n".join(unique_texts)

    def count_equations(self) -> int:
        """Count equation files in the document."""
        count = 0
        with zipfile.ZipFile(self.file_path, "r") as zf:
            for name in zf.namelist():
                if name.startswith("Data/equation-") and name.endswith(".pdf"):
                    count += 1
        return count

    def has_preview(self) -> bool:
        """Check if the document has a preview image."""
        with zipfile.ZipFile(self.file_path, "r") as zf:
            return "preview.jpg" in zf.namelist()

    def extract_preview(self, output_path: str) -> str:
        """Extract the preview image to a file."""
        with zipfile.ZipFile(self.file_path, "r") as zf:
            if "preview.jpg" not in zf.namelist():
                raise ValueError("No preview image found in document")
            with zf.open("preview.jpg") as src:
                output = Path(output_path)
                output.write_bytes(src.read())
        return str(output)
