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
from lxml import etree

from mcp_handley_lab.word.opc.constants import RT, qn

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


def build_hyperlinks(pkg) -> list[HyperlinkInfo]:
    """Build list of all hyperlinks in the document.

    Args:
        pkg: WordPackage or python-docx Document (duck-typed)
    """
    # Get document element and relationships (duck-typed)
    if hasattr(pkg, "document_xml"):
        # WordPackage
        doc_element = pkg.document_xml
        doc_rels = pkg.get_rels("/word/document.xml")
        return _build_hyperlinks_ooxml(doc_element, doc_rels)
    else:
        # python-docx Document - use legacy implementation
        return _build_hyperlinks_docx(pkg)


def _build_hyperlinks_ooxml(doc_element, doc_rels) -> list[HyperlinkInfo]:
    """Build hyperlinks using pure OOXML."""
    result = []

    # Find all w:hyperlink elements in body (including in tables)
    for idx, hyperlink in enumerate(doc_element.iter(qn("w:hyperlink"))):
        # Get rId for external links
        rId = hyperlink.get(qn("r:id"))
        # Get anchor for internal bookmarks
        anchor = hyperlink.get(qn("w:anchor"))

        # Extract text from all runs
        text_parts = []
        for t in hyperlink.iter(qn("w:t")):
            if t.text:
                text_parts.append(t.text)
        text = "".join(text_parts)

        # Resolve URL from relationships
        address = ""
        is_external = False
        if rId:
            rel = doc_rels.get(rId)
            if rel:
                # Only use target as URL if it's an external relationship
                if rel.is_external:
                    address = rel.target
                    is_external = True
                # Non-external hyperlinks (rare) - rel.target is a part path, not URL

        # Build URL (external address or internal #fragment)
        if address:
            url = f"{address}#{anchor}" if anchor else address
        elif anchor:
            url = f"#{anchor}"
        else:
            url = ""

        result.append(
            HyperlinkInfo(
                index=idx,
                text=text,
                url=url,
                address=address,
                fragment=anchor or "",
                is_external=is_external,
            )
        )

    return result


def _build_hyperlinks_docx(doc) -> list[HyperlinkInfo]:
    """Build hyperlinks using python-docx (legacy)."""
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph as DocxParagraph

    def iter_all_paragraphs_docx():
        """Iterate paragraphs using python-docx (for Hyperlink support)."""
        for child in doc.element.body.iterchildren():
            if isinstance(child, CT_P):
                yield DocxParagraph(child, doc)
            elif isinstance(child, CT_Tbl):
                tbl = Table(child, doc)
                for row in tbl.rows:
                    for cell in row.cells:
                        yield from cell.paragraphs

    result = []
    idx = 0
    for para in iter_all_paragraphs_docx():
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


def add_hyperlink(
    paragraph: Paragraph,
    text: str,
    address: str = "",
    fragment: str = "",
) -> None:
    """Add hyperlink to paragraph using OOXML.

    Args:
        paragraph: Target paragraph
        text: Visible link text
        address: URL or file path (external link)
        fragment: Bookmark name (internal link) or URL anchor

    Either address or fragment must be provided.
    For external links: address="https://example.com"
    For internal links: address="", fragment="BookmarkName"
    For external with anchor: address="https://example.com", fragment="section"
    """
    if not text:
        raise ValueError("Hyperlink text cannot be empty")
    if not address and not fragment:
        raise ValueError("Either address or fragment must be provided")

    # Normalize fragment (strip leading #)
    if fragment.startswith("#"):
        fragment = fragment[1:]

    # Validate: fragment must not be empty after normalization for internal links
    if not address and not fragment:
        raise ValueError("Fragment cannot be empty for internal links")

    # Create the w:hyperlink element
    hyperlink = etree.Element(qn("w:hyperlink"))

    if address:
        # External link - create relationship
        # Use _parent.part for stability across python-docx versions
        part = paragraph._parent.part
        full_url = f"{address}#{fragment}" if fragment else address
        r_id = part.relate_to(full_url, RT.HYPERLINK, is_external=True)
        # Note: qn("r:id") is correct - CT_Hyperlink uses OptionalAttribute("r:id", ...)
        hyperlink.set(qn("r:id"), r_id)
    else:
        # Internal link - use anchor attribute (validate bookmark name format)
        from mcp_handley_lab.word.ops.bookmarks import _validate_bookmark_name

        _validate_bookmark_name(fragment)
        hyperlink.set(qn("w:anchor"), fragment)

    # Set history attribute (standard Word behavior)
    hyperlink.set(qn("w:history"), "1")

    # Create run with text
    run = etree.Element(qn("w:r"))
    run_text = etree.Element(qn("w:t"))
    # Preserve whitespace if needed
    if text[:1].isspace() or text[-1:].isspace() or "  " in text:
        run_text.set(qn("xml:space"), "preserve")
    run_text.text = text
    run.append(run_text)
    hyperlink.append(run)

    # Append hyperlink to paragraph
    paragraph._p.append(hyperlink)


# =============================================================================
# Style Management
# =============================================================================


def _build_styles_ooxml(styles_xml) -> list[StyleInfo]:
    """Build styles list from styles.xml element (pure OOXML)."""
    # Map OOXML type values to API strings
    type_map = {
        "paragraph": "paragraph",
        "character": "character",
        "table": "table",
        "numbering": "list",  # OOXML uses "numbering" for list styles
    }

    # First pass: build styleId -> name mapping for resolving references
    id_to_name = {}
    for style in styles_xml.findall(qn("w:style")):
        style_id = style.get(qn("w:styleId"))
        name_el = style.find(qn("w:name"))
        name = name_el.get(qn("w:val")) if name_el is not None else style_id
        if style_id:
            id_to_name[style_id] = name

    # Second pass: build StyleInfo list
    result = []
    for style in styles_xml.findall(qn("w:style")):
        style_id = style.get(qn("w:styleId"))
        style_type = style.get(qn("w:type"), "paragraph")
        # w:customStyle="1" indicates a custom/user-defined style
        # Absence means built-in (NOT w:default which means "default for type")
        is_builtin = style.get(qn("w:customStyle")) != "1"

        # Get name
        name_el = style.find(qn("w:name"))
        name = name_el.get(qn("w:val")) if name_el is not None else style_id

        # Get base style (w:basedOn)
        based_on = style.find(qn("w:basedOn"))
        base_style_id = based_on.get(qn("w:val")) if based_on is not None else None
        base_style = id_to_name.get(base_style_id) if base_style_id else None

        # Get next style (w:next)
        next_el = style.find(qn("w:next"))
        next_style_id = next_el.get(qn("w:val")) if next_el is not None else None
        next_style = id_to_name.get(next_style_id) if next_style_id else None
        # Don't report next style if it's same as current
        if next_style == name:
            next_style = None

        # Check hidden (w:semiHidden or w:hidden)
        hidden = (
            style.find(qn("w:semiHidden")) is not None
            or style.find(qn("w:hidden")) is not None
        )

        # Check quick style (w:qFormat)
        quick_style = style.find(qn("w:qFormat")) is not None

        result.append(
            StyleInfo(
                name=name,
                style_id=style_id,
                type=type_map.get(style_type, "unknown"),
                builtin=is_builtin,
                base_style=base_style,
                next_style=next_style,
                hidden=hidden,
                quick_style=quick_style,
            )
        )
    return result


def build_styles(pkg) -> list[StyleInfo]:
    """Build list of all styles in the document.

    Args:
        pkg: WordPackage or python-docx Document (duck-typed)
    """
    # WordPackage path (pure OOXML)
    if hasattr(pkg, "document_xml"):
        if not pkg.has_part("/word/styles.xml"):
            return []
        return _build_styles_ooxml(pkg.styles_xml)

    # python-docx Document path (legacy)
    style_type_map = {
        WD_STYLE_TYPE.PARAGRAPH: "paragraph",
        WD_STYLE_TYPE.CHARACTER: "character",
        WD_STYLE_TYPE.TABLE: "table",
        WD_STYLE_TYPE.LIST: "list",
    }
    result = []
    for style in pkg.styles:
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


def _get_style_format_ooxml(styles_xml, style_name: str) -> StyleFormatInfo:
    """Get style formatting from styles.xml (pure OOXML)."""
    # OOXML alignment values
    align_map = {
        "left": "left",
        "center": "center",
        "right": "right",
        "both": "justify",
    }
    type_map = {
        "paragraph": "paragraph",
        "character": "character",
        "table": "table",
        "numbering": "list",
    }

    # Find style by name (w:name/@w:val)
    style_el = None
    for s in styles_xml.findall(qn("w:style")):
        name_el = s.find(qn("w:name"))
        if name_el is not None and name_el.get(qn("w:val")) == style_name:
            style_el = s
            break

    if style_el is None:
        raise ValueError(f"Style not found: {style_name}")

    style_id = style_el.get(qn("w:styleId"))
    style_type = type_map.get(style_el.get(qn("w:type"), "paragraph"), "unknown")
    name_el = style_el.find(qn("w:name"))
    name = name_el.get(qn("w:val")) if name_el is not None else style_id

    # Extract run properties (w:rPr)
    rPr = style_el.find(qn("w:rPr"))
    font_name = None
    font_size = None
    bold = None
    italic = None
    color = None

    if rPr is not None:
        # Font name (w:rFonts/@w:ascii)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            font_name = rFonts.get(qn("w:ascii"))

        # Font size (w:sz/@w:val in half-points)
        sz = rPr.find(qn("w:sz"))
        if sz is not None:
            val = sz.get(qn("w:val"))
            font_size = int(val) / 2 if val else None  # half-points to points

        # Bold (w:b presence or @w:val != "0")
        b = rPr.find(qn("w:b"))
        if b is not None:
            val = b.get(qn("w:val"))
            bold = val != "0" if val else True

        # Italic (w:i)
        i = rPr.find(qn("w:i"))
        if i is not None:
            val = i.get(qn("w:val"))
            italic = val != "0" if val else True

        # Color (w:color/@w:val)
        color_el = rPr.find(qn("w:color"))
        if color_el is not None:
            val = color_el.get(qn("w:val"))
            if val and val != "auto":
                color = val.upper()

    # Extract paragraph properties (w:pPr)
    pPr = style_el.find(qn("w:pPr"))
    alignment = None
    left_indent = None
    space_before = None
    space_after = None
    line_spacing = None

    if pPr is not None:
        # Alignment (w:jc/@w:val)
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            alignment = align_map.get(jc.get(qn("w:val")))

        # Indentation (w:ind/@w:left in twips)
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            left_val = ind.get(qn("w:left"))
            if left_val:
                left_indent = int(left_val) / 1440  # twips to inches

        # Spacing (w:spacing/@w:before, @w:after in twips)
        # w:line interpretation depends on w:lineRule:
        # - "auto" (default): 240ths of a line (240 = single, 360 = 1.5)
        # - "exact" or "atLeast": twips (20 per point)
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            before = spacing.get(qn("w:before"))
            after = spacing.get(qn("w:after"))
            line = spacing.get(qn("w:line"))
            line_rule = spacing.get(qn("w:lineRule"))
            if before:
                space_before = int(before) / 20  # twips to points
            if after:
                space_after = int(after) / 20  # twips to points
            if line:
                if line_rule in ("exact", "atLeast"):
                    line_spacing = int(line) / 20  # twips to points
                else:
                    line_spacing = int(line) / 240  # 240ths of a line

    return StyleFormatInfo(
        name=name,
        style_id=style_id,
        type=style_type,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        color=color,
        alignment=alignment,
        left_indent=left_indent,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
    )


def get_style_format(pkg, style_name: str) -> StyleFormatInfo:
    """Get detailed formatting for a specific style.

    Args:
        pkg: WordPackage or python-docx Document (duck-typed)
        style_name: Name of the style to get formatting for
    """
    # WordPackage path (pure OOXML)
    if hasattr(pkg, "document_xml"):
        if not pkg.has_part("/word/styles.xml"):
            raise ValueError(f"No styles.xml found, cannot get style: {style_name}")
        return _get_style_format_ooxml(pkg.styles_xml, style_name)

    # python-docx Document path (legacy)
    alignment_to_api = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }

    style = pkg.styles[style_name]
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
    """Add a tab stop to a paragraph.

    Args:
        paragraph: The paragraph to add the tab stop to
        position_inches: Position from left margin (must be positive)
        alignment: Tab alignment - left, center, right, decimal
        leader: Tab leader character - spaces, dots, heavy, middle_dot
            Aliases: dot->dots, space->spaces, mid_dot->middle_dot
    """
    # Validate position
    if position_inches <= 0:
        raise ValueError(f"position_inches must be positive, got {position_inches}")

    # Alignment with validation
    align_map = {
        "left": WD_TAB_ALIGNMENT.LEFT,
        "center": WD_TAB_ALIGNMENT.CENTER,
        "right": WD_TAB_ALIGNMENT.RIGHT,
        "decimal": WD_TAB_ALIGNMENT.DECIMAL,
    }
    alignment_lower = alignment.lower()
    if alignment_lower not in align_map:
        raise ValueError(
            f"Invalid alignment '{alignment}'. Valid: {list(align_map.keys())}"
        )

    # Leader with aliases and validation
    leader_aliases = {"dot": "dots", "space": "spaces", "mid_dot": "middle_dot"}
    leader_map = {
        "spaces": WD_TAB_LEADER.SPACES,
        "dots": WD_TAB_LEADER.DOTS,
        "heavy": WD_TAB_LEADER.HEAVY,
        "middle_dot": WD_TAB_LEADER.MIDDLE_DOT,
    }
    leader_lower = leader_aliases.get(leader.lower(), leader.lower())
    if leader_lower not in leader_map:
        raise ValueError(f"Invalid leader '{leader}'. Valid: {list(leader_map.keys())}")

    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(position_inches),
        align_map[alignment_lower],
        leader_map[leader_lower],
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
