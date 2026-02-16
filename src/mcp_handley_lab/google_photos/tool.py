"""Google Photos MCP tool for searching, browsing, and downloading photos.

Uses reverse-engineered batchexecute API with session cookies.
Session must be refreshed externally by gphotos-refresh-session.
"""

from typing import Literal

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.google_photos.shared import GooglePhotosResult

mcp = FastMCP("Google Photos Tool")


@mcp.tool(
    description="""Interact with Google Photos via reverse-engineered batchexecute API.
Requires session cookies at ~/.local/share/gphotos/session.json
(refreshed externally by gphotos-refresh-session).

Actions:
- search: Text search for photos (e.g. "whiteboard", "cat"). Returns media keys + URLs + timestamps.
  Required: query. Optional: limit (default 50).
- list_recent: List recent photos by date range, optionally filtered by camera.
  Optional: days (default 14), camera_make, camera_model, limit (default 50).
- detail: Full metadata for one photo (dimensions, camera EXIF, filename, download URL).
  Required: media_key.
- show: Display a photo visually (returns image content Claude can see). Skips videos.
  Required: media_key.
- download: Download photo(s) to disk. Skips videos.
  Required: media_keys (list). Optional: output_dir (default /tmp/photos-live).
"""
)
def fetch(
    action: Literal["search", "list_recent", "detail", "show", "download"] = Field(
        ...,
        description="Operation to perform.",
    ),
    query: str = Field(default="", description="Search query (for 'search')."),
    media_key: str = Field(
        default="", description="Single media key (for 'detail'/'show')."
    ),
    media_keys: list[str] = Field(
        default_factory=list, description="Media keys (for 'download')."
    ),
    output_dir: str = Field(
        default="/tmp/photos-live", description="Download directory (for 'download')."
    ),
    days: int = Field(default=14, description="Days to look back (for 'list_recent')."),
    camera_make: str = Field(
        default="", description="Filter by camera make (for 'list_recent')."
    ),
    camera_model: str = Field(
        default="", description="Filter by camera model (for 'list_recent')."
    ),
    limit: int = Field(
        default=50, description="Max results (for 'search'/'list_recent')."
    ),
):
    """Dispatch to the appropriate Google Photos operation."""
    from mcp_handley_lab.google_photos.shared import (
        download_photos,
        get_photo_detail,
        list_recent_photos,
        search_photos,
        show_photo,
    )

    if action == "search":
        if not query:
            raise ValueError("'query' is required for search action")
        return GooglePhotosResult(photos=search_photos(query, limit))
    elif action == "list_recent":
        return GooglePhotosResult(
            photos=list_recent_photos(days, camera_make, camera_model, limit)
        )
    elif action == "detail":
        if not media_key:
            raise ValueError("'media_key' is required for detail action")
        return GooglePhotosResult(detail=get_photo_detail(media_key))
    elif action == "show":
        if not media_key:
            raise ValueError("'media_key' is required for show action")
        return show_photo(media_key)
    elif action == "download":
        if not media_keys:
            raise ValueError("'media_keys' is required for download action")
        return GooglePhotosResult(download=download_photos(media_keys, output_dir))
    raise ValueError(f"Unknown action: {action}")
