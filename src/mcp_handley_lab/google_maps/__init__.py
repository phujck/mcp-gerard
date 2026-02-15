"""Google Maps MCP tool for directions and routing.

Usage:
    from mcp_handley_lab.google_maps import get_directions

    result = get_directions(
        origin="Cambridge, UK",
        destination="London, UK",
        mode="transit",
        departure_time="tomorrow 9am",
    )
"""

from mcp_handley_lab.google_maps.shared import get_directions

__all__ = ["get_directions"]
