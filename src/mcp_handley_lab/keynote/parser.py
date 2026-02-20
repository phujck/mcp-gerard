"""Keynote file parser using keynote-parser library."""

import re
import subprocess
from pathlib import Path

import yaml


class KeynoteParser:
    """Parser for Apple Keynote presentations."""

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Keynote file not found: {file_path}")
        if self.file_path.suffix.lower() != ".key":
            raise ValueError(f"Not a Keynote file: {file_path}")

    def list_files(self) -> list[str]:
        """List all files in the Keynote archive."""
        result = subprocess.run(
            ["keynote-parser", "ls", str(self.file_path)],
            capture_output=True,
            text=True,
            check=True,
        )
        return [
            line.strip() for line in result.stdout.strip().split("\n") if line.strip()
        ]

    def get_slide_files(self) -> list[str]:
        """Get list of slide files (not template slides).

        Matches patterns like:
        - Index/Slide.iwa
        - Index/Slide-1221831.iwa
        - Index/Slide-1222258-2.iwa
        """
        files = self.list_files()
        slide_pattern = re.compile(r"Index/Slide(?:-[\d-]+)?\.iwa$")
        return sorted([f for f in files if slide_pattern.match(f)])

    def cat_file(self, internal_path: str) -> str:
        """Read contents of a file within the Keynote archive."""
        result = subprocess.run(
            ["keynote-parser", "cat", str(self.file_path), internal_path],
            capture_output=True,
            text=True,
            check=True,
        )
        return result.stdout

    def extract_text_from_slide(self, slide_path: str) -> list[str]:
        """Extract all text content from a slide YAML."""
        yaml_content = self.cat_file(slide_path)
        data = yaml.safe_load(yaml_content)
        texts = []
        self._find_text_recursive(data, texts)
        # Filter out placeholder characters and empty strings
        return [t for t in texts if t and t.strip() and t != "\ufffc"]

    def _find_text_recursive(self, obj: object, texts: list[str]) -> None:
        """Recursively find all text arrays in the YAML structure."""
        if isinstance(obj, dict):
            if "text" in obj and isinstance(obj["text"], list):
                for item in obj["text"]:
                    if isinstance(item, str) and item.strip():
                        texts.append(item.strip())
            for value in obj.values():
                self._find_text_recursive(value, texts)
        elif isinstance(obj, list):
            for item in obj:
                self._find_text_recursive(item, texts)

    def extract_notes_from_slide(self, slide_path: str) -> str:
        """Extract presenter notes from a slide."""
        yaml_content = self.cat_file(slide_path)
        data = yaml.safe_load(yaml_content)
        notes = []
        self._find_notes_recursive(data, notes)
        return "\n".join(notes)

    def _find_notes_recursive(self, obj: object, notes: list[str]) -> None:
        """Find presenter notes in the YAML structure."""
        if isinstance(obj, dict):
            # Notes are typically in objects with _pbtype containing "Note"
            if obj.get("_pbtype", "").endswith("NoteArchive"):
                self._find_text_recursive(obj, notes)
            for value in obj.values():
                self._find_notes_recursive(value, notes)
        elif isinstance(obj, list):
            for item in obj:
                self._find_notes_recursive(item, notes)

    def get_all_slides_text(self) -> list[dict]:
        """Get text content from all slides."""
        slides = []
        slide_files = self.get_slide_files()

        for i, slide_file in enumerate(slide_files, 1):
            texts = self.extract_text_from_slide(slide_file)
            notes = self.extract_notes_from_slide(slide_file)

            # First text is typically the title
            title = texts[0] if texts else ""
            body_text = texts[1:] if len(texts) > 1 else []

            slides.append(
                {
                    "number": i,
                    "file": slide_file,
                    "title": title,
                    "text_content": body_text,
                    "notes": notes,
                    "has_notes": bool(notes.strip()),
                }
            )

        return slides

    def find_replace(
        self, find: str, replace: str, output_path: str | None = None
    ) -> tuple[str, int]:
        """Find and replace text in the presentation.

        Returns tuple of (output_path, replacement_count).
        """
        output = output_path or str(self.file_path)
        cmd = [
            "keynote-parser",
            "replace",
            str(self.file_path),
            "--find",
            find,
            "--replace",
            replace,
        ]
        if output_path:
            cmd.extend(["--output", output_path])

        result = subprocess.run(cmd, capture_output=True, text=True, check=True)

        # Parse output to get replacement count
        count = 0
        for line in result.stdout.split("\n"):
            if "Replaced" in line:
                # Extract count from "Replaced X with Y N times" or single replacement
                match = re.search(r"(\d+) times", line)
                if match:
                    count += int(match.group(1))
                elif "Replaced" in line:
                    count += 1

        return output, count

    def unpack(self, output_dir: str | None = None) -> tuple[str, int]:
        """Unpack Keynote file to a directory.

        Returns tuple of (output_directory, file_count).
        """
        output = output_dir or str(self.file_path).replace(".key", "")
        cmd = ["keynote-parser", "unpack", str(self.file_path), "--output", output]

        subprocess.run(cmd, capture_output=True, text=True, check=True)

        # Count files in output directory
        output_path = Path(output)
        file_count = sum(1 for _ in output_path.rglob("*") if _.is_file())

        return output, file_count

    @staticmethod
    def pack(input_dir: str, output_path: str | None = None) -> str:
        """Pack a directory into a Keynote file.

        Returns the output file path.
        """
        input_path = Path(input_dir)
        if not input_path.is_dir():
            raise ValueError(f"Input must be a directory: {input_dir}")

        output = output_path or f"{input_dir}.key"
        cmd = ["keynote-parser", "pack", input_dir, "--output", output]

        subprocess.run(cmd, capture_output=True, text=True, check=True)

        return output
