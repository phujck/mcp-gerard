"""Blog tool: draft and prepare LaTeX blog posts for Substack."""

from __future__ import annotations

import os
import re
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("blog")

BLOG_DRAFTS_PATH = Path(
    os.environ.get("BLOG_DRAFTS_PATH", "~/Projects/blog")
).expanduser()

_SHARED_DIR = Path(__file__).parent / "_shared"


def _slug(title: str) -> str:
    """Convert a title to a URL-safe slug."""
    return re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-")


def _extract_title_from_tex(tex_path: Path) -> str:
    """Extract \\title{} from a .tex file, or return the filename stem."""
    try:
        text = tex_path.read_text(encoding="utf-8", errors="ignore")
        m = re.search(r"\\title\{([^}]+)\}", text)
        if m:
            return m.group(1)
    except OSError:
        pass
    return tex_path.parent.name


def _render_equation_pdflatex(equation: str, is_display: bool, output_path: Path) -> bool:
    """Render a LaTeX equation to PNG using pdflatex + pdftoppm. Returns True on success."""
    size = "12pt" if is_display else "10pt"
    display_env = f"\\[ {equation} \\]" if is_display else f"$\\displaystyle {equation}$"
    tex_source = (
        f"\\documentclass[{size}]{{standalone}}\n"
        "\\usepackage{amsmath,amssymb}\n"
        "\\begin{document}\n"
        f"{display_env}\n"
        "\\end{document}\n"
    )
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        tex_file = tmp / "eq.tex"
        tex_file.write_text(tex_source, encoding="utf-8")
        try:
            result = subprocess.run(
                ["pdflatex", "-interaction=nonstopmode", "eq.tex"],
                cwd=tmpdir,
                capture_output=True,
                timeout=30,
            )
            if result.returncode != 0:
                return False
            pdf_file = tmp / "eq.pdf"
            if not pdf_file.exists():
                return False
            # Convert PDF page 1 to PNG at 150 dpi
            subprocess.run(
                ["pdftoppm", "-r", "150", "-png", "-singlefile", "eq.pdf", "eq"],
                cwd=tmpdir,
                capture_output=True,
                timeout=30,
                check=True,
            )
            png_file = tmp / "eq.png"
            if not png_file.exists():
                # pdftoppm may add -1 suffix
                candidates = list(tmp.glob("eq*.png"))
                if candidates:
                    png_file = candidates[0]
                else:
                    return False
            import shutil
            shutil.copy2(png_file, output_path)
            return True
        except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError):
            return False


def _render_equation_matplotlib(equation: str, is_display: bool, output_path: Path) -> bool:
    """Render a LaTeX equation to PNG using matplotlib.mathtext as fallback."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        fig = plt.figure(figsize=(0.01, 0.01))
        fontsize = 14 if is_display else 11
        text = fig.text(0, 0, f"${equation}$", fontsize=fontsize)
        fig.savefig(
            str(output_path),
            dpi=150,
            bbox_inches="tight",
            pad_inches=0.05,
            transparent=True,
        )
        plt.close(fig)
        return True
    except Exception:
        return False


def _render_equation(equation: str, is_display: bool, output_path: Path) -> bool:
    """Render equation to PNG, trying pdflatex first then matplotlib."""
    return _render_equation_pdflatex(equation, is_display, output_path) or \
           _render_equation_matplotlib(equation, is_display, output_path)


@mcp.tool()
def blog_new_draft(title: str) -> str:
    """Create a new LaTeX blog draft directory with a main.tex template.

    Creates ~/Projects/blog/<slug>/main.tex with standard preamble/macros includes.
    Returns the path to main.tex.
    """
    slug = _slug(title)
    draft_dir = BLOG_DRAFTS_PATH / slug
    draft_dir.mkdir(parents=True, exist_ok=True)
    (draft_dir / "images").mkdir(exist_ok=True)

    date_str = datetime.now().strftime("%Y-%m-%d")
    tex_content = (
        f"\\documentclass[12pt]{{article}}\n"
        "\\input{preamble}\n"
        "\\input{macros}\n"
        "\n"
        f"\\title{{{title}}}\n"
        "\\author{Gerard McCaul}\n"
        f"\\date{{{date_str}}}\n"
        "\n"
        "\\begin{document}\n"
        "\\maketitle\n"
        "\n"
        "% Write your post here\n"
        "\n"
        "\\end{document}\n"
    )
    main_tex = draft_dir / "main.tex"
    main_tex.write_text(tex_content, encoding="utf-8")

    # Copy preamble and macros from _shared if they exist
    for shared_file in ("preamble.tex", "macros.tex"):
        shared_src = _SHARED_DIR / shared_file
        dest = draft_dir / shared_file
        if shared_src.exists() and not dest.exists():
            import shutil
            shutil.copy2(shared_src, dest)

    return str(main_tex)


@mcp.tool()
def blog_compile(draft_path: str) -> str:
    """Convert a LaTeX draft to Substack-ready Markdown.

    Renders all equations ($...$ and $$...$$) as PNG images.
    Saves images to <draft_dir>/images/eq-NNN.png.
    Outputs draft.md next to the .tex file.
    Returns a summary with the path to draft.md and equation count.
    """
    tex_path = Path(draft_path).expanduser().resolve()
    if not tex_path.exists():
        return f"Error: file not found: {tex_path}"

    draft_dir = tex_path.parent
    images_dir = draft_dir / "images"
    images_dir.mkdir(exist_ok=True)

    tex_source = tex_path.read_text(encoding="utf-8", errors="ignore")

    # Extract equations: display ($$...$$) and inline ($...$)
    # We process display first to avoid double-matching
    equations: list[tuple[str, bool, str]] = []  # (equation, is_display, original_text)

    # Collect all math spans with their positions
    math_spans: list[tuple[int, int, str, bool]] = []  # (start, end, eq, is_display)

    # Display math: $$...$$ or \[...\]
    for m in re.finditer(r'\$\$(.+?)\$\$', tex_source, re.DOTALL):
        math_spans.append((m.start(), m.end(), m.group(1).strip(), True))
    for m in re.finditer(r'\\\[(.+?)\\\]', tex_source, re.DOTALL):
        math_spans.append((m.start(), m.end(), m.group(1).strip(), True))

    # Inline math: $...$ (not $$)
    for m in re.finditer(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', tex_source, re.DOTALL):
        # Skip if this position overlaps with a display math span
        pos_start, pos_end = m.start(), m.end()
        overlaps = any(
            not (pos_end <= ds or pos_start >= de)
            for ds, de, _, _ in math_spans
        )
        if not overlaps:
            math_spans.append((pos_start, pos_end, m.group(1).strip(), False))

    # Sort by position
    math_spans.sort(key=lambda x: x[0])

    eq_count = 0
    rendered_count = 0

    # Build replacement map: position -> image markdown
    replacements: list[tuple[int, int, str]] = []
    seen_equations: dict[str, str] = {}  # eq_text -> image_path (dedup)

    for start, end, eq_text, is_display in math_spans:
        if eq_text in seen_equations:
            img_path = seen_equations[eq_text]
        else:
            eq_count += 1
            img_filename = f"eq-{eq_count:03d}.png"
            img_path = str(images_dir / img_filename)
            success = _render_equation(eq_text, is_display, Path(img_path))
            if success:
                rendered_count += 1
                seen_equations[eq_text] = img_path
            else:
                # Keep original if rendering failed
                replacements.append((start, end, tex_source[start:end]))
                continue

        rel_path = f"images/{Path(img_path).name}"
        if is_display:
            md_ref = f"\n\n![equation]({rel_path})\n\n"
        else:
            md_ref = f"![eq]({rel_path})"
        replacements.append((start, end, md_ref))

    # Apply replacements in reverse order
    result_source = tex_source
    for start, end, replacement in sorted(replacements, key=lambda x: x[0], reverse=True):
        result_source = result_source[:start] + replacement + result_source[end:]

    # Try to convert with pandoc, fallback to basic stripping
    draft_md_path = draft_dir / "draft.md"

    pandoc_available = False
    try:
        subprocess.run(["pandoc", "--version"], capture_output=True, timeout=10)
        pandoc_available = True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    if pandoc_available:
        # Write modified tex to temp file and run pandoc
        with tempfile.NamedTemporaryFile(suffix=".tex", mode="w", encoding="utf-8", delete=False) as tmp:
            tmp.write(result_source)
            tmp_path = tmp.name
        try:
            result = subprocess.run(
                ["pandoc", tmp_path, "-f", "latex", "-t", "markdown", "-o", str(draft_md_path)],
                capture_output=True,
                timeout=60,
            )
            if result.returncode != 0:
                pandoc_available = False
        finally:
            Path(tmp_path).unlink(missing_ok=True)

    if not pandoc_available:
        # Basic LaTeX stripping fallback
        md_content = result_source
        # Remove LaTeX preamble
        body_match = re.search(r"\\begin\{document\}(.+?)\\end\{document\}", md_content, re.DOTALL)
        if body_match:
            md_content = body_match.group(1).strip()
        # Strip common LaTeX commands
        md_content = re.sub(r"\\maketitle", "", md_content)
        md_content = re.sub(r"\\section\*?\{([^}]+)\}", r"## \1", md_content)
        md_content = re.sub(r"\\subsection\*?\{([^}]+)\}", r"### \1", md_content)
        md_content = re.sub(r"\\textbf\{([^}]+)\}", r"**\1**", md_content)
        md_content = re.sub(r"\\emph\{([^}]+)\}", r"*\1*", md_content)
        md_content = re.sub(r"\\[a-zA-Z]+\{([^}]*)\}", r"\1", md_content)
        md_content = re.sub(r"\\[a-zA-Z]+", "", md_content)
        draft_md_path.write_text(md_content.strip(), encoding="utf-8")

    return (
        f"Compiled: {draft_md_path}\n"
        f"Equations rendered: {rendered_count}/{eq_count}\n"
        f"Images saved to: {images_dir}"
    )


@mcp.tool()
def blog_list_drafts() -> list[dict]:
    """List all blog drafts in the blog drafts directory.

    Returns a list of dicts with title, slug, last_modified, and compiled status.
    """
    if not BLOG_DRAFTS_PATH.exists():
        return []

    drafts = []
    for entry in sorted(BLOG_DRAFTS_PATH.iterdir()):
        if not entry.is_dir():
            continue
        main_tex = entry / "main.tex"
        draft_md = entry / "draft.md"

        title = _extract_title_from_tex(main_tex) if main_tex.exists() else entry.name
        last_modified = (
            datetime.fromtimestamp(main_tex.stat().st_mtime).isoformat()
            if main_tex.exists()
            else None
        )
        drafts.append(
            {
                "slug": entry.name,
                "title": title,
                "last_modified": last_modified,
                "compiled": draft_md.exists(),
                "path": str(entry),
            }
        )
    return drafts


@mcp.tool()
def blog_open_draft(slug: str) -> str:
    """Open a blog draft folder in VSCode.

    Returns the path to the draft directory.
    """
    draft_dir = BLOG_DRAFTS_PATH / slug
    if not draft_dir.exists():
        return f"Error: draft not found: {draft_dir}"
    try:
        subprocess.Popen(["code", str(draft_dir)])
    except FileNotFoundError:
        return f"VSCode not found in PATH. Draft is at: {draft_dir}"
    return str(draft_dir)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
