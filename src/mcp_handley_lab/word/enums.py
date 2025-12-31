"""Word document enums for internal operations.

These enums provide values for internal API operations. Note:
- OOXML uses STRING values (e.g., w:jc="center"), not these integers
- Numeric values may not exactly match python-docx equivalents
- These are for internal use; python-docx enums are still used where
  python-docx APIs are called until full migration is complete
- String mapping helpers (STR_TO_*, *_TO_STR) are the primary interface
  for OOXML serialization
"""

from enum import IntEnum


class WdStyleType(IntEnum):
    """Style types matching OOXML w:type values."""

    PARAGRAPH = 1
    CHARACTER = 2
    TABLE = 3
    LIST = 4


class WdAlignParagraph(IntEnum):
    """Paragraph alignment matching OOXML w:jc values."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3


class WdOrient(IntEnum):
    """Page orientation matching OOXML w:orient values."""

    PORTRAIT = 0
    LANDSCAPE = 1


class WdSection(IntEnum):
    """Section break types matching OOXML w:type values."""

    CONTINUOUS = 0
    NEW_COLUMN = 1
    NEW_PAGE = 2
    EVEN_PAGE = 3
    ODD_PAGE = 4


class WdTabAlignment(IntEnum):
    """Tab stop alignment matching OOXML w:val values."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    DECIMAL = 3
    BAR = 4


class WdTabLeader(IntEnum):
    """Tab stop leader matching OOXML w:leader values."""

    SPACES = 0
    DOTS = 1
    HEAVY = 4
    MIDDLE_DOT = 5


class WdColorIndex(IntEnum):
    """Highlight color index matching OOXML w:val values."""

    BLACK = 1
    BLUE = 2
    TURQUOISE = 3
    BRIGHT_GREEN = 4
    PINK = 5
    RED = 6
    YELLOW = 7
    WHITE = 8
    DARK_BLUE = 9
    DARK_RED = 13
    DARK_YELLOW = 14
    GRAY_50 = 15
    GRAY_25 = 16


class WdBreakType(IntEnum):
    """Break types for w:br elements."""

    PAGE = 1
    COLUMN = 2
    LINE = 3


# String mappings for API compatibility
STYLE_TYPE_TO_STR = {
    WdStyleType.PARAGRAPH: "paragraph",
    WdStyleType.CHARACTER: "character",
    WdStyleType.TABLE: "table",
    WdStyleType.LIST: "list",
}

STR_TO_STYLE_TYPE = {v: k for k, v in STYLE_TYPE_TO_STR.items()}

ALIGN_TO_STR = {
    WdAlignParagraph.LEFT: "left",
    WdAlignParagraph.CENTER: "center",
    WdAlignParagraph.RIGHT: "right",
    WdAlignParagraph.JUSTIFY: "justify",
}

STR_TO_ALIGN = {v: k for k, v in ALIGN_TO_STR.items()}

ORIENT_TO_STR = {
    WdOrient.PORTRAIT: "portrait",
    WdOrient.LANDSCAPE: "landscape",
}

STR_TO_ORIENT = {v: k for k, v in ORIENT_TO_STR.items()}

SECTION_TO_STR = {
    WdSection.CONTINUOUS: "continuous",
    WdSection.NEW_COLUMN: "new_column",
    WdSection.NEW_PAGE: "new_page",
    WdSection.EVEN_PAGE: "even_page",
    WdSection.ODD_PAGE: "odd_page",
}

STR_TO_SECTION = {v: k for k, v in SECTION_TO_STR.items()}

TAB_ALIGN_TO_STR = {
    WdTabAlignment.LEFT: "left",
    WdTabAlignment.CENTER: "center",
    WdTabAlignment.RIGHT: "right",
    WdTabAlignment.DECIMAL: "decimal",
    WdTabAlignment.BAR: "bar",
}

STR_TO_TAB_ALIGN = {v: k for k, v in TAB_ALIGN_TO_STR.items()}

TAB_LEADER_TO_STR = {
    WdTabLeader.SPACES: "spaces",
    WdTabLeader.DOTS: "dots",
    WdTabLeader.HEAVY: "heavy",
    WdTabLeader.MIDDLE_DOT: "middle_dot",
}

STR_TO_TAB_LEADER = {v: k for k, v in TAB_LEADER_TO_STR.items()}

COLOR_TO_STR = {
    WdColorIndex.BLACK: "black",
    WdColorIndex.BLUE: "blue",
    WdColorIndex.TURQUOISE: "cyan",
    WdColorIndex.BRIGHT_GREEN: "green",
    WdColorIndex.PINK: "pink",
    WdColorIndex.RED: "red",
    WdColorIndex.YELLOW: "yellow",
    WdColorIndex.WHITE: "white",
    WdColorIndex.DARK_BLUE: "dark_blue",
    WdColorIndex.DARK_RED: "dark_red",
    WdColorIndex.DARK_YELLOW: "dark_yellow",
    WdColorIndex.GRAY_50: "dark_gray",
    WdColorIndex.GRAY_25: "gray",
}

STR_TO_COLOR = {v: k for k, v in COLOR_TO_STR.items()}
