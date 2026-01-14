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
        """Create a new PowerPoint presentation from template.

        Uses a minimal valid template to avoid repair issues.
        """
        template_path = Path(__file__).parent / "template.pptx"
        if template_path.exists():
            return cls.open(template_path)

        # Fallback: create minimal presentation
        # Note: This minimal version may trigger repairs in PowerPoint
        pkg = cls()
        pkg._init_minimal_presentation()
        return pkg

    def _init_minimal_presentation(self) -> None:
        """Initialize a minimal valid presentation structure."""
        # Content types
        self._content_types["pptx"] = CT.PML_PRESENTATION_MAIN
        self._content_types["/ppt/presentation.xml"] = CT.PML_PRESENTATION_MAIN

        # Package relationships
        pkg_rels = self.get_pkg_rels()
        pkg_rels.get_or_add(RT.OFFICE_DOCUMENT, "/ppt/presentation.xml")

        # Minimal presentation.xml
        pres = etree.Element(
            qn("p:presentation"),
            nsmap={
                "p": NSMAP["p"],
                "a": NSMAP["a"],
                "r": NSMAP["r"],
            },
        )
        etree.SubElement(pres, qn("p:sldMasterIdLst"))
        etree.SubElement(pres, qn("p:sldIdLst"))
        etree.SubElement(
            pres,
            qn("p:sldSz"),
            cx="9144000",
            cy="6858000",
            type="screen4x3",
        )
        etree.SubElement(
            pres,
            qn("p:notesSz"),
            cx="6858000",
            cy="9144000",
        )

        self.set_xml("/ppt/presentation.xml", pres)
        self._presentation_path = "/ppt/presentation.xml"

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
