"""Comment operations.

Contains functions for:
- Building comment lists (simple and with threading)
- Adding comments to blocks
- Replying to comments
- Resolving/unresolving comments
- Threading and extended comment info
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.text.paragraph import Paragraph
from lxml import etree

from mcp_handley_lab.word.opc.constants import qn

if TYPE_CHECKING:
    from docx import Document

from mcp_handley_lab.word.models import CommentInfo

# =============================================================================
# Constants
# =============================================================================

# Word 2012 namespace for extended comment features
_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_COMMENTS_EXT_NS = {"w15": _W15_NS}

# Content type for commentsExtended.xml
_COMMENTS_EXTENDED_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"


# =============================================================================
# Basic Comment Functions
# =============================================================================


def build_comments(doc: Document) -> list[CommentInfo]:
    """Build list of CommentInfo from document comments."""
    return [
        CommentInfo(
            id=c.comment_id,
            author=c.author,
            initials=c.initials,
            timestamp=c.timestamp.isoformat() if c.timestamp else None,
            text=c.text,
        )
        for c in doc.comments
    ]


def add_comment_to_block(
    doc: Document,
    paragraph: Paragraph,
    text: str,
    author: str = "",
    initials: str = "",
) -> int:
    """Add a comment anchored to all runs in a paragraph. Returns comment_id."""
    return doc.add_comment(
        runs=paragraph.runs, text=text, author=author, initials=initials
    ).comment_id


# =============================================================================
# Extended Comments (Threading/Resolution)
# =============================================================================


def _get_comments_extended_part(doc: Document):
    """Get commentsExtended.xml part if exists, else None.

    The part may not exist in older Word documents or those without threaded comments.
    """
    package = doc.part.package
    # Look for the part by content type
    for rel in package.rels.values():
        if hasattr(rel, "_target") and hasattr(rel._target, "content_type"):
            if rel._target.content_type == _COMMENTS_EXTENDED_CT:
                return rel._target
    # Also try by part name pattern
    try:
        for part in package.iter_parts():
            if "/commentsExtended.xml" in part.partname:
                return part
    except Exception:
        pass
    return None


def _parse_comment_threading(doc: Document) -> dict:
    """Parse threading/resolution from commentsExtended.xml.

    Returns: {para_id: {'parent_para_id': str|None, 'done': bool}}
    """
    result = {}
    ext_part = _get_comments_extended_part(doc)
    if ext_part is None:
        return result

    try:
        root = etree.fromstring(ext_part.blob)

        # commentsExtended contains w15:commentEx elements
        # Use Clark notation for w15 namespace since qn() doesn't know w15
        para_id_attr = f"{{{_W15_NS}}}paraId"
        done_attr = f"{{{_W15_NS}}}done"
        parent_attr = f"{{{_W15_NS}}}paraIdParent"

        for comment_ex in root.findall(".//w15:commentEx", namespaces=_COMMENTS_EXT_NS):
            # w15:paraId is the link to the comment (matches w:comment's w15:paraId)
            para_id = comment_ex.get(para_id_attr)
            if not para_id:
                continue

            # w15:done indicates resolution
            done_str = comment_ex.get(done_attr)
            done = done_str == "1" if done_str else False

            # w15:paraIdParent indicates parent comment for threading
            parent_para_id = comment_ex.get(parent_attr)

            result[para_id] = {"parent_para_id": parent_para_id, "done": done}
    except (etree.XMLSyntaxError, AttributeError):
        # Narrow exception handling - only catch XML parsing errors
        pass

    return result


def _get_comment_para_id_map(doc: Document) -> dict:
    """Build mapping from w15:paraId to comment_id.

    python-docx comments don't expose paraId, so we parse comments.xml directly.
    """
    para_id_to_comment_id = {}
    try:
        comments_part = doc.part._comments_part
        if comments_part is None:
            return para_id_to_comment_id

        root = etree.fromstring(comments_part.blob)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # Use Clark notation for attributes - w:id uses standard WML namespace
        w_id_attr = qn("w:id")
        # w15:paraId uses Word 2012 namespace (not in python-docx's qn)
        para_id_attr = f"{{{_W15_NS}}}paraId"

        for comment in root.findall(".//w:comment", namespaces=ns):
            comment_id = comment.get(w_id_attr)
            para_id = comment.get(para_id_attr)
            # Guard against None para_id
            if comment_id is not None and para_id is not None:
                para_id_to_comment_id[para_id] = int(comment_id)
    except (etree.XMLSyntaxError, AttributeError, ValueError):
        # Narrow exception handling for XML/parsing errors
        pass

    return para_id_to_comment_id


def build_comments_with_threading(doc: Document) -> list[dict]:
    """Build comments list with threading info from extended part.

    Falls back to flat list if commentsExtended.xml not present.
    Returns list of dicts compatible with CommentInfo model.
    """
    # Get base comment info from python-docx
    comments = {}
    for c in doc.comments:
        comments[c.comment_id] = {
            "id": c.comment_id,
            "author": c.author,
            "initials": c.initials,
            "timestamp": c.timestamp.isoformat() if c.timestamp else None,
            "text": c.text,
            "parent_id": None,
            "resolved": False,
            "replies": [],
        }

    # Parse extended info if available
    threading_info = _parse_comment_threading(doc)
    para_id_map = _get_comment_para_id_map(doc)

    # Apply threading info
    for para_id, info in threading_info.items():
        comment_id = para_id_map.get(para_id)
        if comment_id is None or comment_id not in comments:
            continue

        comments[comment_id]["resolved"] = info["done"]

        # Find parent comment by para_id
        parent_para_id = info.get("parent_para_id")
        if parent_para_id:
            parent_comment_id = para_id_map.get(parent_para_id)
            if parent_comment_id is not None:
                comments[comment_id]["parent_id"] = parent_comment_id

    # Build replies lists
    for comment_id, comment in comments.items():
        parent_id = comment["parent_id"]
        if parent_id is not None and parent_id in comments:
            comments[parent_id]["replies"].append(comment_id)

    # Return sorted by comment ID for deterministic ordering
    return sorted(comments.values(), key=lambda c: c["id"])


def reply_to_comment(
    doc: Document, parent_id: int, text: str, author: str = "", initials: str = ""
) -> int:
    """Add reply to existing comment. Returns new comment ID.

    Note: Full threading support requires commentsExtended.xml manipulation
    which involves OPC packaging. This creates a basic reply comment
    anchored to the same location as the parent.
    """
    # Validate parent comment exists
    parent_comment = None
    for c in doc.comments:
        if c.comment_id == parent_id:
            parent_comment = c
            break

    if parent_comment is None:
        raise ValueError(f"Parent comment not found: {parent_id}")

    # Find runs anchored to parent comment by searching for commentRangeStart
    # with matching ID in the document body
    anchored_runs = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # Find the commentRangeStart for this comment ID
    range_start = doc.element.find(
        f".//w:commentRangeStart[@w:id='{parent_id}']", namespaces=ns
    )

    if range_start is not None:
        # Find the containing paragraph
        parent_el = range_start.getparent()
        while parent_el is not None:
            if parent_el.tag == qn("w:p"):
                # Found paragraph - get its runs
                para = Paragraph(parent_el, doc)
                anchored_runs = para.runs
                break
            parent_el = parent_el.getparent()

    # Fall back to first paragraph with runs if no anchored runs found
    if not anchored_runs:
        for para in doc.paragraphs:
            if para.runs:
                anchored_runs = para.runs
                break

    if not anchored_runs:
        raise ValueError("No runs available to anchor reply comment")

    # Create reply comment anchored to same location as parent
    new_comment = doc.add_comment(
        runs=anchored_runs, text=text, author=author, initials=initials
    )

    return new_comment.comment_id


def resolve_comment(doc: Document, comment_id: int) -> None:
    """Mark comment as resolved.

    Note: Full resolution support requires commentsExtended.xml manipulation.
    This is a placeholder that validates the comment exists.
    Word 2013+ uses commentsExtended.xml with w15:done="1" attribute.
    """
    # Validate comment exists
    found = False
    for c in doc.comments:
        if c.comment_id == comment_id:
            found = True
            break

    if not found:
        raise ValueError(f"Comment not found: {comment_id}")

    # Note: Actual resolution requires modifying commentsExtended.xml
    # which involves complex OPC packaging. This is a placeholder.
    # For now, we validate the comment exists - full implementation
    # would create/modify commentsExtended.xml part.


def unresolve_comment(doc: Document, comment_id: int) -> None:
    """Mark comment as unresolved (clears 'done' state).

    Note: Full support requires commentsExtended.xml manipulation.
    """
    # Validate comment exists
    found = False
    for c in doc.comments:
        if c.comment_id == comment_id:
            found = True
            break

    if not found:
        raise ValueError(f"Comment not found: {comment_id}")
