"""Document properties operations.

Contains functions for:
- Core properties (title, author, created, modified, revision)
- Custom properties (user-defined metadata in docProps/custom.xml)
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from docx import Document

from mcp_handley_lab.word.models import CustomPropertyInfo, DocumentMeta

# =============================================================================
# Core Properties
# =============================================================================


def get_document_meta(doc: Document) -> DocumentMeta:
    """Extract document metadata from core properties and custom properties."""
    cp = doc.core_properties
    custom_props = get_custom_properties(doc)
    return DocumentMeta(
        title=cp.title or "",
        author=cp.author or "",
        created=cp.created.isoformat() if cp.created else "",
        modified=cp.modified.isoformat() if cp.modified else "",
        revision=cp.revision or 0,
        sections=len(doc.sections),
        custom_properties=custom_props,
    )


def set_document_meta(doc: Document, **kwargs) -> None:
    """Update document core properties. Only updates non-None values."""
    cp = doc.core_properties
    for key, value in kwargs.items():
        if value is not None:
            setattr(cp, key, value)


# =============================================================================
# Custom Properties
# =============================================================================


def get_custom_properties(doc: Document) -> list[CustomPropertyInfo]:
    """Get all custom document properties from docProps/custom.xml."""
    props = []
    # Custom properties are in docProps/custom.xml
    for part in doc.part.package.iter_parts():
        if part.partname == "/docProps/custom.xml":
            root = etree.fromstring(part.blob)
            ns_custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

            for prop in root.findall(f"{{{ns_custom}}}property"):
                name = prop.get("name", "")
                # Find value element - could be vt:lpwstr, vt:i4, vt:bool, vt:filetime, etc.
                value = ""
                prop_type = "string"
                for child in prop:
                    local_tag = (
                        child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    )
                    if local_tag == "lpwstr":
                        value = child.text or ""
                        prop_type = "string"
                    elif local_tag == "i4":
                        value = child.text or "0"
                        prop_type = "int"
                    elif local_tag == "bool":
                        value = child.text or "false"
                        prop_type = "bool"
                    elif local_tag == "filetime":
                        value = child.text or ""
                        prop_type = "datetime"
                    elif local_tag == "r8":
                        value = child.text or "0.0"
                        prop_type = "float"
                    break
                props.append(CustomPropertyInfo(name=name, value=value, type=prop_type))
            break
    return props


def set_custom_property(
    doc: Document, name: str, value: str, prop_type: str = "string"
) -> None:
    """Set or update a custom document property.

    Args:
        doc: The Document object
        name: Property name (must be unique)
        value: Property value as string
        prop_type: One of "string", "int", "bool", "datetime", "float"
    """
    ns_custom = (
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    )
    ns_vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    # Type to element tag mapping
    type_map = {
        "string": f"{{{ns_vt}}}lpwstr",
        "int": f"{{{ns_vt}}}i4",
        "bool": f"{{{ns_vt}}}bool",
        "datetime": f"{{{ns_vt}}}filetime",
        "float": f"{{{ns_vt}}}r8",
    }

    value_tag = type_map.get(prop_type, type_map["string"])

    # Find or create custom.xml part
    custom_part = None
    for part in doc.part.package.iter_parts():
        if part.partname == "/docProps/custom.xml":
            custom_part = part
            break

    if custom_part is None:
        # Create new custom.xml using python-docx's internal mechanisms
        from docx.opc.packuri import PackURI
        from docx.opc.part import Part

        root = etree.Element(
            f"{{{ns_custom}}}Properties",
            nsmap={None: ns_custom, "vt": ns_vt},
        )
        # Add property element
        prop = etree.SubElement(root, f"{{{ns_custom}}}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", "2")
        prop.set("name", name)
        value_el = etree.SubElement(prop, value_tag)
        value_el.text = value

        # Save - need to add the part to the package
        xml_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
        # Create custom properties part using the package's load method
        content_type = (
            "application/vnd.openxmlformats-officedocument.custom-properties+xml"
        )
        part_uri = PackURI("/docProps/custom.xml")

        # Create the part and add to package properly
        custom_part = Part.load(part_uri, content_type, xml_bytes, doc.part.package)

        # Add relationship from package to this part
        doc.part.package.relate_to(
            custom_part,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
        )
    else:
        # Update existing custom.xml
        root = etree.fromstring(custom_part.blob)

        # Find property with this name
        existing = None
        for prop in root.findall(f"{{{ns_custom}}}property"):
            if prop.get("name") == name:
                existing = prop
                break

        if existing is not None:
            # Update existing property
            for child in list(existing):
                existing.remove(child)
            value_el = etree.SubElement(existing, value_tag)
            value_el.text = value
        else:
            # Add new property - find next pid
            max_pid = 1
            for prop in root.findall(f"{{{ns_custom}}}property"):
                pid = int(prop.get("pid", "1"))
                if pid > max_pid:
                    max_pid = pid
            new_pid = max_pid + 1

            prop = etree.SubElement(root, f"{{{ns_custom}}}property")
            prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            prop.set("pid", str(new_pid))
            prop.set("name", name)
            value_el = etree.SubElement(prop, value_tag)
            value_el.text = value

        # Save updated XML back to part
        custom_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def delete_custom_property(doc: Document, name: str) -> bool:
    """Delete a custom document property.

    Args:
        doc: The Document object
        name: Property name to delete

    Returns:
        True if property was found and deleted, False if not found
    """
    ns_custom = (
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    )

    for part in doc.part.package.iter_parts():
        if part.partname == "/docProps/custom.xml":
            root = etree.fromstring(part.blob)

            # Find property with this name
            for prop in root.findall(f"{{{ns_custom}}}property"):
                if prop.get("name") == name:
                    root.remove(prop)
                    part._blob = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8"
                    )
                    return True
            break
    return False
