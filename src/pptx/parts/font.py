"""Font part, for embedded font files in a .pptx package."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part

if TYPE_CHECKING:
    from pptx.package import Package


class FontPart(Part):
    """A font file (.ttf/.otf) embedded in the PPTX package.

    Stored at /ppt/fonts/fontN.fntdata with content type application/x-fontdata.
    """

    partname_template = "/ppt/fonts/font%d.fntdata"

    @classmethod
    def new(cls, font_bytes: bytes, package: Package) -> FontPart:
        """Return a new |FontPart| containing *font_bytes*, added to *package*."""
        partname = package.next_partname(cls.partname_template)
        return cls(partname, CT.X_FONTDATA, package, font_bytes)
