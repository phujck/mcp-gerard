"""Google Photos tool for searching, browsing, and downloading photos.

Usage:
    from mcp_handley_lab.google_photos import search_photos, get_photo_detail
"""

from mcp_handley_lab.google_photos.shared import (
    download_photos,
    get_photo_detail,
    list_recent_photos,
    search_photos,
    show_photo,
)

__all__ = [
    "download_photos",
    "get_photo_detail",
    "list_recent_photos",
    "search_photos",
    "show_photo",
]
