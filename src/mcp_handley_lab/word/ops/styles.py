"""Style, font, and paragraph formatting operations.

Contains functions for:
- Run and paragraph formatting
- Style management (list, get, create, edit, delete)
- Tab stops
- Hyperlinks
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_COLOR_INDEX,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.shared import Inches, Pt, RGBColor
from docx.text.hyperlink import Hyperlink

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph

from mcp_handley_lab.word.models import (
    HyperlinkInfo,
    ParagraphFormatInfo,
    RunInfo,
    StyleFormatInfo,
    StyleInfo,
    TabStopInfo,
)
from mcp_handley_lab.word.ops.core import _iter_all_paragraphs

# =============================================================================
# Constants
# =============================================================================

_HIGHLIGHT_MAP = {
    "yellow": WD_COLOR_INDEX.YELLOW,
    "green": WD_COLOR_INDEX.BRIGHT_GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE,
    "pink": WD_COLOR_INDEX.PINK,
    "blue": WD_COLOR_INDEX.BLUE,
    "red": WD_COLOR_INDEX.RED,
    "dark_blue": WD_COLOR_INDEX.DARK_BLUE,
    "dark_red": WD_COLOR_INDEX.DARK_RED,
    "dark_yellow": WD_COLOR_INDEX.DARK_YELLOW,
    "gray": WD_COLOR_INDEX.GRAY_25,
    "dark_gray": WD_COLOR_INDEX.GRAY_50,
    "black": WD_COLOR_INDEX.BLACK,
    "white": WD_COLOR_INDEX.WHITE,
}
_HIGHLIGHT_REVERSE = {v: k for k, v in _HIGHLIGHT_MAP.items()}

_RUN_ATTRS = {
    "bold": "bold",
    "italic": "italic",
    "underline": "underline",
    "strike": "font.strike",
    "double_strike": "font.double_strike",
    "subscript": "font.subscript",
    "superscript": "font.superscript",
    "all_caps": "font.all_caps",
    "small_caps": "font.small_caps",
    "hidden": "font.hidden",
    "emboss": "font.emboss",
    "imprint": "font.imprint",
    "outline": "font.outline",
    "shadow": "font.shadow",
    "font_name": "font.name",
}

_PARA_INCH_ATTRS = {"left_indent", "right_indent", "first_line_indent"}
_PARA_PT_ATTRS = {"space_before", "space_after"}
_PARA_DIRECT_ATTRS = {"keep_with_next", "page_break_before"}
_RUN_FORMAT_KEYS = set(_RUN_ATTRS) | {"style", "font_size", "color", "highlight_color"}


# =============================================================================
# Run Building
# =============================================================================


def _build_run_info(
    run, index: int, is_hyperlink: bool = False, hyperlink_url: str | None = None
) -> RunInfo:
    """Build RunInfo from a Run object."""
    return RunInfo(
        index=index,
        text=run.text or "",
        bold=run.bold,
        italic=run.italic,
        underline=run.underline,
        font_name=run.font.name,
        font_size=run.font.size.pt if run.font.size else None,
        color=str(run.font.color.rgb)
        if run.font.color and run.font.color.rgb
        else None,
        highlight_color=_HIGHLIGHT_REVERSE.get(run.font.highlight_color),
        strike=run.font.strike,
        double_strike=run.font.double_strike,
        subscript=run.font.subscript,
        superscript=run.font.superscript,
        style=run.style.name if run.style else None,
        is_hyperlink=is_hyperlink,
        hyperlink_url=hyperlink_url,
        all_caps=run.font.all_caps,
        small_caps=run.font.small_caps,
        hidden=run.font.hidden,
        emboss=run.font.emboss,
        imprint=run.font.imprint,
        outline=run.font.outline,
        shadow=run.font.shadow,
    )


def build_runs(paragraph: Paragraph) -> list[RunInfo]:
    """Build list of RunInfo for all runs in a paragraph, including hyperlink runs."""
    result = []
    idx = 0
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            url = item.url
            for run in item.runs:
                result.append(
                    _build_run_info(run, idx, is_hyperlink=True, hyperlink_url=url)
                )
                idx += 1
        else:  # Run
            result.append(_build_run_info(item, idx))
            idx += 1
    return result


def build_hyperlinks(doc: Document) -> list[HyperlinkInfo]:
    """Build list of all hyperlinks in the document."""
    result = []
    idx = 0
    for para, _el in _iter_all_paragraphs(doc):
        for item in para.iter_inner_content():
            if isinstance(item, Hyperlink):
                result.append(
                    HyperlinkInfo(
                        index=idx,
                        text=item.text,
                        url=item.url,
                        address=item.address,
                        fragment=item.fragment,
                        is_external=bool(item._hyperlink.rId),
                    )
                )
                idx += 1
    return result


# =============================================================================
# Style Management
# =============================================================================


def build_styles(doc: Document) -> list[StyleInfo]:
    """Build list of all styles in the document."""
    style_type_map = {
        WD_STYLE_TYPE.PARAGRAPH: "paragraph",
        WD_STYLE_TYPE.CHARACTER: "character",
        WD_STYLE_TYPE.TABLE: "table",
        WD_STYLE_TYPE.LIST: "list",
    }
    result = []
    for style in doc.styles:
        base = getattr(style, "base_style", None)
        next_style = getattr(style, "next_paragraph_style", None)
        hidden = getattr(style, "hidden", False)
        quick_style = getattr(style, "quick_style", False)
        result.append(
            StyleInfo(
                name=style.name,
                style_id=style.style_id,
                type=style_type_map.get(style.type, "unknown"),
                builtin=style.builtin,
                base_style=base.name if base else None,
                next_style=next_style.name
                if next_style and next_style != style
                else None,
                hidden=hidden,
                quick_style=quick_style,
            )
        )
    return result


def get_style_format(doc: Document, style_name: str) -> StyleFormatInfo:
    """Get detailed formatting for a specific style."""
    alignment_to_api = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }

    style = doc.styles[style_name]
    font = style.font

    style_type_map = {
        WD_STYLE_TYPE.PARAGRAPH: "paragraph",
        WD_STYLE_TYPE.CHARACTER: "character",
        WD_STYLE_TYPE.TABLE: "table",
        WD_STYLE_TYPE.LIST: "list",
    }
    style_type = style_type_map.get(style.type, "unknown")

    pf = getattr(style, "paragraph_format", None)
    alignment = None
    left_indent = None
    space_before = None
    space_after = None
    line_spacing = None

    if pf:
        alignment = alignment_to_api.get(pf.alignment) if pf.alignment else None
        left_indent = pf.left_indent.inches if pf.left_indent else None
        space_before = pf.space_before.pt if pf.space_before else None
        space_after = pf.space_after.pt if pf.space_after else None
        line_spacing = (
            pf.line_spacing
            if isinstance(pf.line_spacing, float)
            else (pf.line_spacing.pt if pf.line_spacing else None)
        )

    return StyleFormatInfo(
        name=style.name,
        style_id=style.style_id,
        type=style_type,
        font_name=font.name,
        font_size=font.size.pt if font.size else None,
        bold=font.bold,
        italic=font.italic,
        color=str(font.color.rgb) if font.color and font.color.rgb else None,
        alignment=alignment,
        left_indent=left_indent,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
    )


def edit_style(doc: Document, style_name: str, fmt: dict) -> None:
    """Modify a style definition."""
    style = doc.styles[style_name]
    font = style.font

    if "font_name" in fmt:
        font.name = fmt["font_name"]
    if "font_size" in fmt:
        font.size = Pt(fmt["font_size"])
    if "bold" in fmt:
        font.bold = fmt["bold"]
    if "italic" in fmt:
        font.italic = fmt["italic"]
    if "color" in fmt:
        font.color.rgb = RGBColor.from_string(fmt["color"].lstrip("#"))

    pf = getattr(style, "paragraph_format", None)
    if pf:
        if "alignment" in fmt:
            pf.alignment = getattr(WD_ALIGN_PARAGRAPH, fmt["alignment"].upper())
        if "left_indent" in fmt:
            pf.left_indent = Inches(fmt["left_indent"])
        if "space_before" in fmt:
            pf.space_before = Pt(fmt["space_before"])
        if "space_after" in fmt:
            pf.space_after = Pt(fmt["space_after"])
        if "line_spacing" in fmt:
            val = fmt["line_spacing"]
            pf.line_spacing = val if val < 5 else Pt(val)


def create_style(
    doc: Document,
    name: str,
    style_type: str = "paragraph",
    base_style: str = "Normal",
    formatting: dict | None = None,
) -> str:
    """Create a new custom style. Returns the style ID."""
    type_map = {
        "paragraph": WD_STYLE_TYPE.PARAGRAPH,
        "character": WD_STYLE_TYPE.CHARACTER,
        "table": WD_STYLE_TYPE.TABLE,
    }

    wd_style_type = type_map.get(style_type.lower(), WD_STYLE_TYPE.PARAGRAPH)
    style = doc.styles.add_style(name, wd_style_type)

    if base_style and base_style in doc.styles:
        style.base_style = doc.styles[base_style]

    if formatting:
        edit_style(doc, name, formatting)

    return style.style_id


def delete_style(doc: Document, style_name: str) -> bool:
    """Delete a custom style. Returns True if deleted, False if builtin or not found."""
    try:
        style = doc.styles[style_name]
    except KeyError:
        return False

    if style.builtin:
        return False

    style._element.getparent().remove(style._element)
    return True


# =============================================================================
# Paragraph Formatting
# =============================================================================


def apply_paragraph_formatting(p: Paragraph, fmt: dict) -> None:
    """Apply direct formatting to paragraph (affects all runs)."""
    if "alignment" in fmt:
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, fmt["alignment"].upper())

    pf = p.paragraph_format
    for key, value in fmt.items():
        if key in _PARA_INCH_ATTRS:
            setattr(pf, key, Inches(value))
        elif key in _PARA_PT_ATTRS:
            setattr(pf, key, Pt(value))
        elif key == "line_spacing":
            pf.line_spacing = value if value < 5 else Pt(value)
        elif key in _PARA_DIRECT_ATTRS:
            setattr(pf, key, value)

    for run in p.runs:
        for key, value in fmt.items():
            if key in _RUN_FORMAT_KEYS:
                _set_run_attr(run, key, value)


def build_paragraph_format(paragraph: Paragraph) -> ParagraphFormatInfo:
    """Extract paragraph formatting properties."""
    alignment_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    pf = paragraph.paragraph_format
    alignment = alignment_map.get(paragraph.alignment)
    return ParagraphFormatInfo(
        alignment=alignment,
        left_indent=pf.left_indent.inches if pf.left_indent else None,
        right_indent=pf.right_indent.inches if pf.right_indent else None,
        first_line_indent=pf.first_line_indent.inches if pf.first_line_indent else None,
        space_before=pf.space_before.pt if pf.space_before else None,
        space_after=pf.space_after.pt if pf.space_after else None,
        line_spacing=(
            pf.line_spacing
            if isinstance(pf.line_spacing, float)
            else (pf.line_spacing.pt if pf.line_spacing else None)
        ),
        keep_with_next=pf.keep_with_next,
        page_break_before=pf.page_break_before,
        tab_stops=build_tab_stops(paragraph),
    )


# =============================================================================
# Tab Stops
# =============================================================================


def build_tab_stops(paragraph: Paragraph) -> list[TabStopInfo]:
    """Build list of tab stops for a paragraph."""
    tab_align_map = {
        WD_TAB_ALIGNMENT.LEFT: "left",
        WD_TAB_ALIGNMENT.CENTER: "center",
        WD_TAB_ALIGNMENT.RIGHT: "right",
        WD_TAB_ALIGNMENT.DECIMAL: "decimal",
        WD_TAB_ALIGNMENT.BAR: "bar",
    }
    tab_leader_map = {
        WD_TAB_LEADER.SPACES: "spaces",
        WD_TAB_LEADER.DOTS: "dots",
        WD_TAB_LEADER.HEAVY: "heavy",
        WD_TAB_LEADER.MIDDLE_DOT: "middle_dot",
        None: "spaces",
    }

    result = []
    for tab in paragraph.paragraph_format.tab_stops:
        alignment = tab_align_map.get(tab.alignment, "unknown")
        leader = tab_leader_map.get(tab.leader, "unknown")
        result.append(
            TabStopInfo(
                position_inches=round(tab.position.inches, 4),
                alignment=alignment,
                leader=leader,
            )
        )
    return result


def add_tab_stop(
    paragraph: Paragraph,
    position_inches: float,
    alignment: str = "left",
    leader: str = "spaces",
) -> None:
    """Add a tab stop to a paragraph."""
    align_map = {
        "left": WD_TAB_ALIGNMENT.LEFT,
        "center": WD_TAB_ALIGNMENT.CENTER,
        "right": WD_TAB_ALIGNMENT.RIGHT,
        "decimal": WD_TAB_ALIGNMENT.DECIMAL,
    }
    leader_map = {
        "spaces": WD_TAB_LEADER.SPACES,
        "dots": WD_TAB_LEADER.DOTS,
        "heavy": WD_TAB_LEADER.HEAVY,
        "middle_dot": WD_TAB_LEADER.MIDDLE_DOT,
    }
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(position_inches),
        align_map[alignment.lower()],
        leader_map[leader.lower()],
    )


# =============================================================================
# Run Formatting
# =============================================================================


def _set_run_attr(run, key: str, value) -> None:
    """Set a run attribute by key. Handles nested paths like 'font.bold'."""
    if key == "style":
        run.style = value
    elif key == "font_size":
        run.font.size = Pt(float(value))
    elif key == "color":
        run.font.color.rgb = RGBColor.from_string(value.lstrip("#"))
    elif key == "highlight_color":
        run.font.highlight_color = _HIGHLIGHT_MAP[value.lower()]
    elif key in _RUN_ATTRS:
        path = _RUN_ATTRS[key]
        obj, attr = (run, path) if "." not in path else (run.font, path.split(".")[1])
        setattr(obj, attr, value)


def _resolve_run_by_inner_index(paragraph: Paragraph, run_index: int):
    """Resolve run by iter_inner_content index (matching build_runs indexing)."""
    idx = 0
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            for run in item.runs:
                if idx == run_index:
                    return run
                idx += 1
        else:  # Run
            if idx == run_index:
                return item
            idx += 1
    raise IndexError(f"Run index {run_index} out of range (paragraph has {idx} runs)")


def edit_run_text(paragraph: Paragraph, run_index: int, text: str) -> None:
    """Edit run text. Uses iter_inner_content indexing (includes hyperlink runs)."""
    run = _resolve_run_by_inner_index(paragraph, run_index)
    run.text = text


def edit_run_formatting(paragraph: Paragraph, run_index: int, fmt: dict) -> None:
    """Apply formatting to a specific run. Uses iter_inner_content indexing."""
    run = _resolve_run_by_inner_index(paragraph, run_index)
    for key, value in fmt.items():
        _set_run_attr(run, key, value)
