"""Open Packaging Convention (OPC) layer for Office documents."""

from mcp_handley_lab.word.opc.constants import (
    CT,
    DEFAULT_CONTENT_TYPES,
    NSMAP,
    PFXMAP,
    RT,
    qn,
)
from mcp_handley_lab.word.opc.content_types import ContentTypeMap
from mcp_handley_lab.word.opc.package import WordPackage
from mcp_handley_lab.word.opc.relationships import Relationship, Relationships

__all__ = [
    "CT",
    "ContentTypeMap",
    "DEFAULT_CONTENT_TYPES",
    "NSMAP",
    "PFXMAP",
    "RT",
    "Relationship",
    "Relationships",
    "WordPackage",
    "qn",
]
