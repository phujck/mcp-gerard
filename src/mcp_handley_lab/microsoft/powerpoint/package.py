"""PowerPoint package wrapper extending OpcPackage."""

from __future__ import annotations

import re
from pathlib import Path
from typing import BinaryIO

from lxml import etree

from mcp_handley_lab.microsoft.opc.package import OpcPackage
from mcp_handley_lab.microsoft.powerpoint.constants import CT, NSMAP, RT, qn


class PowerPointPackage(OpcPackage):
    """PowerPoint-specific OPC package wrapper.

    Provides convenience properties and methods for working with .pptx files.
    """

    def __init__(self) -> None:
        super().__init__()
        self._presentation_path: str | None = None
        self._slide_paths: list[tuple[int, str, str]] | None = None

    @property
    def presentation_path(self) -> str:
        """Get the path to presentation.xml via relationships."""
        if self._presentation_path is None:
            pkg_rels = self.get_pkg_rels()
            rid = pkg_rels.rId_for_reltype(RT.OFFICE_DOCUMENT)
            if rid is None:
                raise ValueError("No presentation relationship found in package")
            target = pkg_rels.target_for_rId(rid)
            self._presentation_path = target if target.startswith("/") else "/" + target
        return self._presentation_path

    @property
    def presentation_xml(self) -> etree._Element:
        """Get the presentation.xml root element."""
        return self.get_xml(self.presentation_path)

    def get_slide_paths(self) -> list[tuple[int, str, str]]:
        """Get ordered list of (slide_number, rId, partname) tuples.

        Slide order is determined by p:sldIdLst in presentation.xml.
        """
        if self._slide_paths is not None:
            return self._slide_paths

        pres = self.presentation_xml
        pres_rels = self.get_rels(self.presentation_path)

        # Build rId -> partname mapping
        rid_to_path: dict[str, str] = {}
        for rid, rel in pres_rels.items():
            if rel.reltype == RT.SLIDE:
                rid_to_path[rid] = self.resolve_rel_target(self.presentation_path, rid)

        # Get ordered slide IDs from presentation.xml
        sld_id_lst = pres.find(qn("p:sldIdLst"), NSMAP)
        if sld_id_lst is None:
            self._slide_paths = []
            return self._slide_paths

        result = []
        for idx, sld_id in enumerate(sld_id_lst.findall(qn("p:sldId"), NSMAP), start=1):
            rid = sld_id.get(qn("r:id"))
            if rid and rid in rid_to_path:
                result.append((idx, rid, rid_to_path[rid]))

        self._slide_paths = result
        return self._slide_paths

    def get_slide_xml(self, slide_num: int) -> etree._Element:
        """Get slide XML by 1-based slide number."""
        for num, _rid, partname in self.get_slide_paths():
            if num == slide_num:
                return self.get_xml(partname)
        raise KeyError(f"Slide {slide_num} not found")

    def get_slide_partname(self, slide_num: int) -> str:
        """Get slide partname by 1-based slide number."""
        for num, _rid, partname in self.get_slide_paths():
            if num == slide_num:
                return partname
        raise KeyError(f"Slide {slide_num} not found")

    def get_notes_xml(self, slide_num: int) -> etree._Element | None:
        """Get notes slide XML for a slide, or None if no notes exist."""
        slide_partname = self.get_slide_partname(slide_num)
        slide_rels = self.get_rels(slide_partname)

        rid = slide_rels.rId_for_reltype(RT.NOTES_SLIDE)
        if rid is None:
            return None

        notes_path = self.resolve_rel_target(slide_partname, rid)
        if notes_path not in self._bytes and notes_path not in self._xml:
            return None

        return self.get_xml(notes_path)

    def get_slide_layout_name(self, slide_num: int) -> str | None:
        """Get the layout name for a slide."""
        slide_partname = self.get_slide_partname(slide_num)
        slide_rels = self.get_rels(slide_partname)

        rid = slide_rels.rId_for_reltype(RT.SLIDE_LAYOUT)
        if rid is None:
            return None

        layout_path = self.resolve_rel_target(slide_partname, rid)
        layout_xml = self.get_xml(layout_path)

        # Layout name is in cSld@name attribute
        cSld = layout_xml.find(qn("p:cSld"), NSMAP)
        if cSld is not None:
            return cSld.get("name")
        return None

    def get_slide_dimensions(self) -> tuple[int, int]:
        """Get slide dimensions (width, height) in EMUs."""
        pres = self.presentation_xml
        sld_sz = pres.find(qn("p:sldSz"), NSMAP)
        if sld_sz is None:
            # Default PowerPoint dimensions
            return (9144000, 6858000)
        cx = int(sld_sz.get("cx", "9144000"))
        cy = int(sld_sz.get("cy", "6858000"))
        return (cx, cy)

    @classmethod
    def open(cls, file: str | Path | BinaryIO) -> PowerPointPackage:
        """Open an existing PowerPoint file."""
        pkg = cls()
        if isinstance(file, str | Path):
            with open(file, "rb") as f:
                pkg._load_from_stream(f)
        else:
            pkg._load_from_stream(file)
        return pkg

    @classmethod
    def new(cls) -> PowerPointPackage:
        """Create a new PowerPoint presentation.

        Uses a bundled template if available, otherwise creates a minimal valid
        presentation structure programmatically.
        """
        template_path = Path(__file__).parent / "template.pptx"
        if template_path.exists():
            return cls.open(template_path)

        # Create minimal valid presentation programmatically
        pkg = cls()
        pkg._init_minimal_presentation()
        return pkg

    # Standard layouts for new presentations: (name, type, placeholders)
    # Each placeholder: (type, idx, name, x, y, cx, cy)
    # Coordinates in EMU (914400 EMU = 1 inch)
    _STANDARD_LAYOUTS: list[
        tuple[str, str, list[tuple[str, str, str, str, str, str, str]]]
    ] = [
        (
            "Title Slide",
            "title",
            [
                ("ctrTitle", "0", "Title 1", "685800", "2130425", "7772400", "1470025"),
                (
                    "subTitle",
                    "1",
                    "Subtitle 2",
                    "1371600",
                    "3886200",
                    "6400800",
                    "1752600",
                ),
            ],
        ),
        (
            "Title and Content",
            "obj",
            [
                ("title", "0", "Title 1", "457200", "274638", "8229600", "1143000"),
                ("body", "1", "Content 2", "457200", "1600200", "8229600", "4525963"),
            ],
        ),
        (
            "Title Only",
            "titleOnly",
            [
                ("title", "0", "Title 1", "457200", "274638", "8229600", "1143000"),
            ],
        ),
        ("Blank", "blank", []),
    ]

    def _init_minimal_presentation(self) -> None:
        """Initialize a minimal valid presentation structure.

        Creates all required parts: presentation, slide master, slide layouts, theme.
        """
        layouts = self._STANDARD_LAYOUTS

        # Register content types for presentation parts
        self._content_types["/ppt/presentation.xml"] = CT.PML_PRESENTATION_MAIN
        self._content_types["/ppt/slideMasters/slideMaster1.xml"] = CT.PML_SLIDE_MASTER
        for i in range(len(layouts)):
            self._content_types[f"/ppt/slideLayouts/slideLayout{i + 1}.xml"] = (
                CT.PML_SLIDE_LAYOUT
            )
        self._content_types["/ppt/theme/theme1.xml"] = CT.THEME

        # Package relationships
        pkg_rels = self.get_pkg_rels()
        pkg_rels.get_or_add(RT.OFFICE_DOCUMENT, "/ppt/presentation.xml")

        # Create presentation.xml
        pres = etree.Element(
            qn("p:presentation"),
            nsmap={"p": NSMAP["p"], "a": NSMAP["a"], "r": NSMAP["r"]},
        )
        # Slide master ID list (required)
        sld_master_id_lst = etree.SubElement(pres, qn("p:sldMasterIdLst"))
        sld_master_id = etree.SubElement(sld_master_id_lst, qn("p:sldMasterId"))
        sld_master_id.set("id", "2147483648")
        sld_master_id.set(qn("r:id"), "rId1")
        # Empty slide ID list
        etree.SubElement(pres, qn("p:sldIdLst"))
        # Slide size
        etree.SubElement(
            pres, qn("p:sldSz"), cx="9144000", cy="6858000", type="screen4x3"
        )
        etree.SubElement(pres, qn("p:notesSz"), cx="6858000", cy="9144000")
        self.set_xml("/ppt/presentation.xml", pres)

        # Presentation relationships
        pres_rels = self.get_rels("/ppt/presentation.xml")
        pres_rels.get_or_add(RT.SLIDE_MASTER, "/ppt/slideMasters/slideMaster1.xml")
        pres_rels.get_or_add(RT.THEME, "/ppt/theme/theme1.xml")

        # Create slide master
        master = etree.Element(
            qn("p:sldMaster"),
            nsmap={"p": NSMAP["p"], "a": NSMAP["a"], "r": NSMAP["r"]},
        )
        cSld = etree.SubElement(master, qn("p:cSld"))
        spTree = etree.SubElement(cSld, qn("p:spTree"))
        # Non-visual group shape properties (required for spTree)
        nvGrpSpPr = etree.SubElement(spTree, qn("p:nvGrpSpPr"))
        cNvPr = etree.SubElement(nvGrpSpPr, qn("p:cNvPr"))
        cNvPr.set("id", "1")
        cNvPr.set("name", "")
        etree.SubElement(nvGrpSpPr, qn("p:cNvGrpSpPr"))
        etree.SubElement(nvGrpSpPr, qn("p:nvPr"))
        grpSpPr = etree.SubElement(spTree, qn("p:grpSpPr"))
        xfrm = etree.SubElement(grpSpPr, qn("a:xfrm"))
        etree.SubElement(xfrm, qn("a:off"), x="0", y="0")
        etree.SubElement(xfrm, qn("a:ext"), cx="0", cy="0")
        etree.SubElement(xfrm, qn("a:chOff"), x="0", y="0")
        etree.SubElement(xfrm, qn("a:chExt"), cx="0", cy="0")
        # Color map
        clr_map = etree.SubElement(master, qn("p:clrMap"))
        for attr, val in [
            ("bg1", "lt1"),
            ("tx1", "dk1"),
            ("bg2", "lt2"),
            ("tx2", "dk2"),
            ("accent1", "accent1"),
            ("accent2", "accent2"),
            ("accent3", "accent3"),
            ("accent4", "accent4"),
            ("accent5", "accent5"),
            ("accent6", "accent6"),
            ("hlink", "hlink"),
            ("folHlink", "folHlink"),
        ]:
            clr_map.set(attr, val)
        # Slide layout ID list
        sld_layout_id_lst = etree.SubElement(master, qn("p:sldLayoutIdLst"))
        for i in range(len(layouts)):
            sld_layout_id = etree.SubElement(sld_layout_id_lst, qn("p:sldLayoutId"))
            sld_layout_id.set("id", str(2147483649 + i))
            sld_layout_id.set(qn("r:id"), f"rId{i + 1}")
        self.set_xml("/ppt/slideMasters/slideMaster1.xml", master)

        # Slide master relationships
        master_rels = self.get_rels("/ppt/slideMasters/slideMaster1.xml")
        for i in range(len(layouts)):
            master_rels.get_or_add(
                RT.SLIDE_LAYOUT, f"/ppt/slideLayouts/slideLayout{i + 1}.xml"
            )
        master_rels.get_or_add(RT.THEME, "/ppt/theme/theme1.xml")

        # Create slide layouts
        for i, (layout_name, layout_type, placeholders) in enumerate(layouts):
            layout_path = f"/ppt/slideLayouts/slideLayout{i + 1}.xml"
            layout_xml = self._build_layout_xml(layout_name, layout_type, placeholders)
            self.set_xml(layout_path, layout_xml)
            layout_rels = self.get_rels(layout_path)
            layout_rels.get_or_add(
                RT.SLIDE_MASTER, "/ppt/slideMasters/slideMaster1.xml"
            )

        # Create minimal theme
        theme = etree.Element(
            qn("a:theme"),
            nsmap={"a": NSMAP["a"]},
        )
        theme.set("name", "Office Theme")
        # Theme elements (required)
        theme_elements = etree.SubElement(theme, qn("a:themeElements"))
        # Color scheme
        clr_scheme = etree.SubElement(theme_elements, qn("a:clrScheme"))
        clr_scheme.set("name", "Office")
        for name, rgb in [
            ("dk1", "000000"),
            ("lt1", "FFFFFF"),
            ("dk2", "44546A"),
            ("lt2", "E7E6E6"),
            ("accent1", "4472C4"),
            ("accent2", "ED7D31"),
            ("accent3", "A5A5A5"),
            ("accent4", "FFC000"),
            ("accent5", "5B9BD5"),
            ("accent6", "70AD47"),
            ("hlink", "0563C1"),
            ("folHlink", "954F72"),
        ]:
            el = etree.SubElement(clr_scheme, qn(f"a:{name}"))
            srgb = etree.SubElement(el, qn("a:srgbClr"))
            srgb.set("val", rgb)
        # Font scheme
        font_scheme = etree.SubElement(theme_elements, qn("a:fontScheme"))
        font_scheme.set("name", "Office")
        for kind in ("majorFont", "minorFont"):
            font = etree.SubElement(font_scheme, qn(f"a:{kind}"))
            latin = etree.SubElement(font, qn("a:latin"))
            latin.set("typeface", "Calibri")
            ea = etree.SubElement(font, qn("a:ea"))
            ea.set("typeface", "")
            cs = etree.SubElement(font, qn("a:cs"))
            cs.set("typeface", "")
        # Format scheme
        fmt_scheme = etree.SubElement(theme_elements, qn("a:fmtScheme"))
        fmt_scheme.set("name", "Office")
        # Fill style list
        fill_lst = etree.SubElement(fmt_scheme, qn("a:fillStyleLst"))
        for _ in range(3):
            etree.SubElement(fill_lst, qn("a:solidFill"))
        # Line style list
        ln_lst = etree.SubElement(fmt_scheme, qn("a:lnStyleLst"))
        for _ in range(3):
            etree.SubElement(ln_lst, qn("a:ln"))
        # Effect style list
        effect_lst = etree.SubElement(fmt_scheme, qn("a:effectStyleLst"))
        for _ in range(3):
            etree.SubElement(effect_lst, qn("a:effectStyle"))
        # Background fill style list
        bg_fill_lst = etree.SubElement(fmt_scheme, qn("a:bgFillStyleLst"))
        for _ in range(3):
            etree.SubElement(bg_fill_lst, qn("a:solidFill"))
        self.set_xml("/ppt/theme/theme1.xml", theme)

        self._presentation_path = "/ppt/presentation.xml"

    @staticmethod
    def _build_layout_xml(
        name: str,
        layout_type: str,
        placeholders: list[tuple[str, str, str, str, str, str, str]],
    ) -> etree._Element:
        """Build a slide layout XML element with optional placeholder shapes."""
        layout = etree.Element(
            qn("p:sldLayout"),
            nsmap={"p": NSMAP["p"], "a": NSMAP["a"], "r": NSMAP["r"]},
        )
        layout.set("type", layout_type)
        cSld = etree.SubElement(layout, qn("p:cSld"))
        cSld.set("name", name)
        spTree = etree.SubElement(cSld, qn("p:spTree"))

        # Required group shape properties
        nvGrpSpPr = etree.SubElement(spTree, qn("p:nvGrpSpPr"))
        cNvPr = etree.SubElement(nvGrpSpPr, qn("p:cNvPr"))
        cNvPr.set("id", "1")
        cNvPr.set("name", "")
        etree.SubElement(nvGrpSpPr, qn("p:cNvGrpSpPr"))
        etree.SubElement(nvGrpSpPr, qn("p:nvPr"))
        grpSpPr = etree.SubElement(spTree, qn("p:grpSpPr"))
        xfrm = etree.SubElement(grpSpPr, qn("a:xfrm"))
        etree.SubElement(xfrm, qn("a:off"), x="0", y="0")
        etree.SubElement(xfrm, qn("a:ext"), cx="0", cy="0")
        etree.SubElement(xfrm, qn("a:chOff"), x="0", y="0")
        etree.SubElement(xfrm, qn("a:chExt"), cx="0", cy="0")

        # Add placeholder shapes
        for shape_id, (ph_type, ph_idx, ph_name, x, y, cx, cy) in enumerate(
            placeholders, start=2
        ):
            sp = etree.SubElement(spTree, qn("p:sp"))
            nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
            sp_cNvPr = etree.SubElement(nvSpPr, qn("p:cNvPr"))
            sp_cNvPr.set("id", str(shape_id))
            sp_cNvPr.set("name", ph_name)
            cNvSpPr = etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
            sp_locks = etree.SubElement(cNvSpPr, qn("a:spLocks"))
            sp_locks.set("noGrp", "1")
            nvPr = etree.SubElement(nvSpPr, qn("p:nvPr"))
            ph = etree.SubElement(nvPr, qn("p:ph"))
            ph.set("type", ph_type)
            ph.set("idx", ph_idx)
            spPr = etree.SubElement(sp, qn("p:spPr"))
            sp_xfrm = etree.SubElement(spPr, qn("a:xfrm"))
            etree.SubElement(sp_xfrm, qn("a:off"), x=x, y=y)
            etree.SubElement(sp_xfrm, qn("a:ext"), cx=cx, cy=cy)

        etree.SubElement(layout, qn("p:clrMapOvr"))
        return layout

    def invalidate_caches(self) -> None:
        """Clear cached values after modifications."""
        self._slide_paths = None

    def next_partname(self, prefix: str, ext: str) -> str:
        """Get next available partname, avoiding collisions.

        Scans existing partnames to find the maximum number and returns prefix + (max+1) + ext.
        This is essential for avoiding collisions after slide/notes deletions.

        Args:
            prefix: Path prefix (e.g., "/ppt/slides/slide")
            ext: File extension (e.g., ".xml")

        Returns:
            Next available partname (e.g., "/ppt/slides/slide4.xml")
        """
        pattern = re.compile(rf"{re.escape(prefix)}(\d+){re.escape(ext)}$")
        max_num = 0
        for partname in self._bytes.keys() | self._xml.keys():
            match = pattern.search(partname)
            if match:
                max_num = max(max_num, int(match.group(1)))
        return f"{prefix}{max_num + 1}{ext}"
