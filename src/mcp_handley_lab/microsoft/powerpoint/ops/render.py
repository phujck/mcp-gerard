"""Render PowerPoint presentations to PNG images or PDF using libreoffice."""

import subprocess
import tempfile
from pathlib import Path


def _convert_to_pdf(pptx_path: Path, output_dir: Path, profile_dir: Path) -> Path:
    """Convert a PowerPoint presentation to PDF using libreoffice."""
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--nologo",
            "--norestore",
            "--nolockcheck",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
        check=True,
    )
    return list(output_dir.glob("*.pdf"))[0]


def render_to_pdf(file_path: str) -> bytes:
    """Render a PowerPoint presentation to PDF."""
    pptx_path = Path(file_path).resolve()
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        profile_dir = tmp / "profile"
        profile_dir.mkdir()
        pdf_path = _convert_to_pdf(pptx_path, tmp, profile_dir)
        return pdf_path.read_bytes()


def render_to_images(
    file_path: str,
    slides: list[int],
    dpi: int = 150,
) -> list[tuple[int, bytes]]:
    """Render PowerPoint slides to PNG images."""
    if not slides:
        raise ValueError("slides is required")
    unique_slides = sorted(set(slides))
    if len(unique_slides) > 5:
        raise ValueError(f"max 5 slides allowed; requested {len(unique_slides)}")

    pptx_path = Path(file_path).resolve()
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        profile_dir = tmp / "profile"
        profile_dir.mkdir()

        pdf_path = _convert_to_pdf(pptx_path, tmp, profile_dir)

        result = []
        for slide_num in unique_slides:
            png_prefix = tmp / f"s{slide_num}"
            subprocess.run(
                [
                    "pdftoppm",
                    "-png",
                    "-r",
                    str(dpi),
                    "-f",
                    str(slide_num),
                    "-l",
                    str(slide_num),
                    str(pdf_path),
                    str(png_prefix),
                ],
                capture_output=True,
                text=True,
                timeout=60,
                check=True,
            )
            expected_png = tmp / f"s{slide_num}-{slide_num}.png"
            result.append((slide_num, expected_png.read_bytes()))

        return result
