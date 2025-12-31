"""Shared types: Length units and colors."""

from __future__ import annotations


class Length(int):
    """Length value stored as EMU (English Metric Units).

    1 inch = 914,400 EMU
    1 pt = 12,700 EMU
    1 cm = 360,000 EMU
    1 twip = 635 EMU (1/20 of a point)
    """

    _EMUS_PER_INCH = 914400
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PT = 12700
    _EMUS_PER_TWIP = 635

    @property
    def cm(self) -> float:
        return self / self._EMUS_PER_CM

    @property
    def inches(self) -> float:
        return self / self._EMUS_PER_INCH

    @property
    def mm(self) -> float:
        return self / self._EMUS_PER_MM

    @property
    def pt(self) -> float:
        return self / self._EMUS_PER_PT

    @property
    def twips(self) -> int:
        return int(self / self._EMUS_PER_TWIP)

    @property
    def emu(self) -> int:
        return int(self)


class Inches(Length):
    """Length specified in inches."""

    def __new__(cls, inches: float) -> Inches:
        return Length.__new__(cls, int(inches * Length._EMUS_PER_INCH))


class Cm(Length):
    """Length specified in centimeters."""

    def __new__(cls, cm: float) -> Cm:
        return Length.__new__(cls, int(cm * Length._EMUS_PER_CM))


class Mm(Length):
    """Length specified in millimeters."""

    def __new__(cls, mm: float) -> Mm:
        return Length.__new__(cls, int(mm * Length._EMUS_PER_MM))


class Pt(Length):
    """Length specified in points (1/72 inch)."""

    def __new__(cls, points: float) -> Pt:
        return Length.__new__(cls, int(points * Length._EMUS_PER_PT))


class Emu(Length):
    """Length specified in EMUs (English Metric Units)."""

    def __new__(cls, emu: int) -> Emu:
        return Length.__new__(cls, emu)


class Twips(Length):
    """Length specified in twips (1/20 of a point)."""

    def __new__(cls, twips: int) -> Twips:
        return Length.__new__(cls, int(twips * Length._EMUS_PER_TWIP))


class RGBColor(tuple):
    """Immutable RGB color value."""

    def __new__(cls, r: int, g: int, b: int) -> RGBColor:
        return tuple.__new__(cls, (r, g, b))

    def __str__(self) -> str:
        return f"{self[0]:02X}{self[1]:02X}{self[2]:02X}"

    @classmethod
    def from_string(cls, rgb_hex: str) -> RGBColor:
        """Parse hex string like 'FF0000' or '#FF0000'."""
        s = rgb_hex.lstrip("#")
        return cls(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
