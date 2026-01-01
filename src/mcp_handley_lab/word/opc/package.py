"""WordPackage - general OPC package for Word documents."""

from __future__ import annotations

import mimetypes
import os
import tempfile
import zipfile
from collections.abc import Iterator
from io import BytesIO
from pathlib import Path
from posixpath import normpath
from typing import BinaryIO

from lxml import etree

from mcp_handley_lab.word.opc.constants import CT, RT, qn
from mcp_handley_lab.word.opc.content_types import ContentTypeMap
from mcp_handley_lab.word.opc.relationships import Relationships


class WordPackage:
    """In-memory representation of a .docx package.

    Provides general part API:
    - get_xml(partname) / set_xml(partname, root) - XML parts
    - get_bytes(partname) / set_bytes(partname, data) - binary parts
    - get_rels(partname) - relationships for any part
    - relate_to(partname, target, reltype) - add relationship from any part

    Uses dirty tracking to minimize serialization on save.
    """

    def __init__(self) -> None:
        self._bytes: dict[str, bytes] = {}  # Original bytes from ZIP
        self._xml: dict[str, etree._Element] = {}  # Parsed XML (lazy)
        self._dirty_xml: set[str] = set()  # Partnames with modified XML
        self._dirty_bytes: set[str] = set()  # Partnames with new/modified bytes
        self._content_types = ContentTypeMap()
        self._pkg_rels = Relationships()  # /_rels/.rels
        self._part_rels: dict[str, Relationships] = {}  # partname -> Relationships
        self._dirty_rels: set[str] = set()  # Partnames with modified relationships

    # === General Part API ===

    def get_xml(self, partname: str) -> etree._Element:
        """Get parsed XML for any part. Caches the parse."""
        if partname in self._xml:
            return self._xml[partname]
        self._xml[partname] = etree.fromstring(self._bytes[partname])
        return self._xml[partname]

    def set_xml(
        self, partname: str, root: etree._Element, content_type: str = ""
    ) -> None:
        """Set XML for a part. Marks as dirty for save."""
        self._xml[partname] = root
        self._dirty_xml.add(partname)
        if partname not in self._bytes:
            self._bytes[partname] = b""  # Placeholder for new parts
        if content_type:
            self._content_types[partname] = content_type

    def get_bytes(self, partname: str) -> bytes:
        """Get raw bytes for any part."""
        return self._bytes[partname]

    def set_bytes(self, partname: str, data: bytes, content_type: str = "") -> None:
        """Set bytes for a part. Marks as dirty."""
        self._bytes[partname] = data
        self._dirty_bytes.add(partname)
        if content_type:
            self._content_types[partname] = content_type

    def mark_xml_dirty(self, partname: str) -> None:
        """Mark an XML part as modified so it will be serialized on save.

        Use after in-place modification of elements returned by get_xml().
        """
        if partname not in self._xml:
            raise ValueError(f"Part not parsed as XML: {partname}")
        self._dirty_xml.add(partname)

    def has_part(self, partname: str) -> bool:
        """Check if part exists."""
        return partname in self._bytes

    def iter_partnames(self) -> Iterator[str]:
        """Iterate all partnames."""
        return iter(self._bytes.keys())

    def drop_part(self, partname: str) -> None:
        """Remove a part and its relationships."""
        self._bytes.pop(partname, None)
        self._xml.pop(partname, None)
        self._part_rels.pop(partname, None)
        self._dirty_xml.discard(partname)
        self._dirty_bytes.discard(partname)
        self._dirty_rels.discard(partname)
        self._content_types._overrides.pop(partname, None)

    # === Relationships API ===

    def get_rels(self, partname: str) -> Relationships:
        """Get relationships for any part. Creates empty if not exists."""
        from posixpath import dirname

        if partname not in self._part_rels:
            rels_path = self._rels_path_for(partname)
            if rels_path in self._bytes:
                base_uri = dirname(partname)
                self._part_rels[partname] = Relationships.from_xml(
                    self._bytes[rels_path], base_uri
                )
            else:
                self._part_rels[partname] = Relationships(dirname(partname))
        return self._part_rels[partname]

    def relate_to(
        self,
        source_partname: str,
        target: str,
        reltype: str,
        *,
        is_external: bool = False,
    ) -> str:
        """Add relationship from source part to target. Returns rId."""
        rels = self.get_rels(source_partname)
        rId = rels.get_or_add(reltype, target, is_external)
        self._dirty_rels.add(source_partname)
        return rId

    def resolve_rel_target(self, partname: str, rId: str) -> str:
        """Resolve rId to absolute partname.

        Handles relative paths (e.g., "media/image1.png" from /word/document.xml).
        Returns normalized absolute partname starting with "/".
        """
        from posixpath import dirname

        rels = self.get_rels(partname)
        rel = rels.get(rId)
        if not rel:
            raise KeyError(f"Relationship {rId} not found in {partname}")
        if rel.is_external:
            return rel.target  # External URL, return as-is
        # Internal: resolve relative to source part's directory
        base_uri = dirname(partname)
        resolved = normpath(f"{base_uri}/{rel.target}")
        # Ensure leading slash for absolute OPC partname
        if not resolved.startswith("/"):
            resolved = "/" + resolved
        return resolved

    # === Convenience Properties ===

    @property
    def document_xml(self) -> etree._Element:
        return self.get_xml("/word/document.xml")

    @property
    def styles_xml(self) -> etree._Element | None:
        return (
            self.get_xml("/word/styles.xml")
            if self.has_part("/word/styles.xml")
            else None
        )

    @property
    def numbering_xml(self) -> etree._Element | None:
        return (
            self.get_xml("/word/numbering.xml")
            if self.has_part("/word/numbering.xml")
            else None
        )

    @property
    def settings_xml(self) -> etree._Element | None:
        return (
            self.get_xml("/word/settings.xml")
            if self.has_part("/word/settings.xml")
            else None
        )

    @property
    def comments_xml(self) -> etree._Element | None:
        return (
            self.get_xml("/word/comments.xml")
            if self.has_part("/word/comments.xml")
            else None
        )

    @property
    def footnotes_xml(self) -> etree._Element | None:
        return (
            self.get_xml("/word/footnotes.xml")
            if self.has_part("/word/footnotes.xml")
            else None
        )

    @property
    def endnotes_xml(self) -> etree._Element | None:
        return (
            self.get_xml("/word/endnotes.xml")
            if self.has_part("/word/endnotes.xml")
            else None
        )

    @property
    def body(self) -> etree._Element:
        return self.document_xml.find(qn("w:body"))

    # === Loading ===

    @classmethod
    def open(cls, file: str | Path | BinaryIO) -> WordPackage:
        """Open a .docx file."""
        pkg = cls()
        if isinstance(file, str | Path):
            with open(file, "rb") as f:
                pkg._load_from_stream(f)
        else:
            pkg._load_from_stream(file)
        return pkg

    @classmethod
    def new(cls) -> WordPackage:
        """Create a new empty Word document."""
        pkg = cls()
        pkg._create_minimal_document()
        return pkg

    def _load_from_stream(self, stream: BinaryIO) -> None:
        """Load package contents from stream."""
        with zipfile.ZipFile(stream, "r") as z:
            for name in z.namelist():
                partname = f"/{name}" if not name.startswith("/") else name
                self._bytes[partname] = z.read(name)

        # Parse content types
        if "/[Content_Types].xml" in self._bytes:
            self._content_types = ContentTypeMap.from_xml(
                self._bytes["/[Content_Types].xml"]
            )

        # Parse package relationships
        if "/_rels/.rels" in self._bytes:
            self._pkg_rels = Relationships.from_xml(self._bytes["/_rels/.rels"])

    def _create_minimal_document(self) -> None:
        """Create minimal valid Word document structure.

        Creates document.xml, styles.xml, settings.xml, and core.xml
        for a document that reliably opens in Microsoft Word.
        """
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        # Document XML
        doc_nsmap = {None: w_ns, "r": r_ns}
        document = etree.Element(qn("w:document"), nsmap=doc_nsmap)
        body = etree.SubElement(document, qn("w:body"))
        p = etree.SubElement(body, qn("w:p"))
        etree.SubElement(p, qn("w:r"))

        # Section properties with page size and margins
        sectPr = body.find(qn("w:sectPr"))
        if sectPr is None:
            sectPr = etree.SubElement(body, qn("w:sectPr"))
        pgSz = etree.SubElement(sectPr, qn("w:pgSz"))
        pgSz.set(qn("w:w"), "12240")  # 8.5 inches in twips
        pgSz.set(qn("w:h"), "15840")  # 11 inches in twips
        pgMar = etree.SubElement(sectPr, qn("w:pgMar"))
        pgMar.set(qn("w:top"), "1440")  # 1 inch
        pgMar.set(qn("w:right"), "1440")
        pgMar.set(qn("w:bottom"), "1440")
        pgMar.set(qn("w:left"), "1440")
        pgMar.set(qn("w:header"), "720")  # 0.5 inch
        pgMar.set(qn("w:footer"), "720")
        pgMar.set(qn("w:gutter"), "0")

        self._xml["/word/document.xml"] = document
        self._bytes["/word/document.xml"] = b""
        self._dirty_xml.add("/word/document.xml")
        self._content_types["/word/document.xml"] = CT.WML_DOCUMENT_MAIN

        # Styles XML (minimal)
        styles_nsmap = {None: w_ns}
        styles = etree.Element(qn("w:styles"), nsmap=styles_nsmap)
        # Add Normal style
        normal = etree.SubElement(styles, qn("w:style"))
        normal.set(qn("w:type"), "paragraph")
        normal.set(qn("w:styleId"), "Normal")
        normal.set(qn("w:default"), "1")
        name_el = etree.SubElement(normal, qn("w:name"))
        name_el.set(qn("w:val"), "Normal")
        self._xml["/word/styles.xml"] = styles
        self._bytes["/word/styles.xml"] = b""
        self._dirty_xml.add("/word/styles.xml")
        self._content_types["/word/styles.xml"] = CT.WML_STYLES

        # Settings XML (minimal)
        settings = etree.Element(qn("w:settings"), nsmap=styles_nsmap)
        self._xml["/word/settings.xml"] = settings
        self._bytes["/word/settings.xml"] = b""
        self._dirty_xml.add("/word/settings.xml")
        self._content_types["/word/settings.xml"] = CT.WML_SETTINGS

        # Core properties (docProps/core.xml)
        ns_cp = (
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        )
        ns_dc = "http://purl.org/dc/elements/1.1/"
        ns_dcterms = "http://purl.org/dc/terms/"
        ns_xsi = "http://www.w3.org/2001/XMLSchema-instance"
        core_nsmap = {"cp": ns_cp, "dc": ns_dc, "dcterms": ns_dcterms, "xsi": ns_xsi}
        core = etree.Element(f"{{{ns_cp}}}coreProperties", nsmap=core_nsmap)
        etree.SubElement(core, f"{{{ns_dc}}}title")
        etree.SubElement(core, f"{{{ns_dc}}}creator")
        etree.SubElement(core, f"{{{ns_cp}}}revision").text = "1"
        self._xml["/docProps/core.xml"] = core
        self._bytes["/docProps/core.xml"] = b""
        self._dirty_xml.add("/docProps/core.xml")
        self._content_types["/docProps/core.xml"] = CT.OPC_CORE_PROPERTIES

        # Package relationships
        self._pkg_rels.add(RT.OFFICE_DOCUMENT, "word/document.xml")
        self._pkg_rels.add(RT.CORE_PROPERTIES, "docProps/core.xml")

        # Document relationships to styles and settings
        doc_rels = self.get_rels("/word/document.xml")
        doc_rels.add(RT.STYLES, "styles.xml")
        doc_rels.add(RT.SETTINGS, "settings.xml")
        self._dirty_rels.add("/word/document.xml")

    # === Saving ===

    def save(self, file: str | Path | BinaryIO) -> None:
        """Save package to .docx file. Uses atomic write for file paths."""
        stream = BytesIO()

        with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as z:
            # Write content types
            z.writestr("[Content_Types].xml", self._content_types.to_xml())

            # Write package relationships
            if self._pkg_rels:
                z.writestr("_rels/.rels", self._pkg_rels.to_xml())

            # Collect all .rels paths we'll write from _part_rels
            rels_to_write = {self._rels_path_for(pn) for pn in self._part_rels}

            # Write all parts
            for partname in self._bytes:
                # Skip content types and pkg rels (already written)
                if partname in ("/[Content_Types].xml", "/_rels/.rels"):
                    continue
                # Skip original .rels if we have modified version
                if partname in rels_to_write:
                    continue

                membername = partname.lstrip("/")

                # Only serialize XML if dirty; otherwise keep original bytes
                if partname in self._dirty_xml:
                    xml_bytes = etree.tostring(
                        self._xml[partname],
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    z.writestr(membername, xml_bytes)
                else:
                    z.writestr(membername, self._bytes[partname])

            # Write part relationships (only if dirty or new)
            for partname, rels in self._part_rels.items():
                if not rels:
                    continue
                rels_path = self._rels_path_for(partname)
                if partname in self._dirty_rels or rels_path not in self._bytes:
                    # Dirty or new: serialize
                    z.writestr(rels_path.lstrip("/"), rels.to_xml())
                elif rels_path in self._bytes:
                    # Unchanged: write original bytes
                    z.writestr(rels_path.lstrip("/"), self._bytes[rels_path])

        # Atomic write for file paths
        if isinstance(file, str | Path):
            file_path = Path(file)
            fd, temp_path = tempfile.mkstemp(suffix=".docx", dir=file_path.parent)
            try:
                os.write(fd, stream.getvalue())
                os.close(fd)
                os.replace(temp_path, file_path)
            except Exception:
                os.close(fd)
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
                raise
        else:
            file.write(stream.getvalue())

    @staticmethod
    def _rels_path_for(partname: str) -> str:
        """Compute .rels path: /word/document.xml -> /word/_rels/document.xml.rels"""
        # Use posixpath for OPC paths (always forward slashes regardless of platform)
        from posixpath import basename, dirname

        parent = dirname(partname)
        name = basename(partname)
        return f"{parent}/_rels/{name}.rels"

    # === Image Helpers ===

    def add_image(self, image_bytes: bytes, ext: str = "png") -> str:
        """Add image to package and return rId from document.xml."""
        ext = ext.lower().lstrip(".")
        content_type, _ = mimetypes.guess_type(f"file.{ext}")
        if not content_type:
            content_type = f"image/{ext}"

        # Find next image number
        n = 1
        while self.has_part(f"/word/media/image{n}.{ext}"):
            n += 1

        partname = f"/word/media/image{n}.{ext}"
        self.set_bytes(partname, image_bytes)
        self._content_types.add_default(ext, content_type)

        # Add relationship from document.xml
        target = f"media/image{n}.{ext}"
        return self.relate_to("/word/document.xml", target, RT.IMAGE)
