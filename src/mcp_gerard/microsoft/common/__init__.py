"""Common utilities shared across Microsoft Office formats (Word, Excel, PowerPoint)."""

from mcp_gerard.microsoft.common.constants import EMU_PER_INCH, EMU_PER_PT
from mcp_gerard.microsoft.common.properties import (
    delete_custom_property,
    get_core_properties,
    get_custom_properties,
    set_core_properties,
    set_custom_property,
)
from mcp_gerard.microsoft.common.render import (
    convert_to_pdf,
    render_pages_to_images,
)

__all__ = [
    "EMU_PER_INCH",
    "EMU_PER_PT",
    "convert_to_pdf",
    "delete_custom_property",
    "get_core_properties",
    "get_custom_properties",
    "render_pages_to_images",
    "set_core_properties",
    "set_custom_property",
]
