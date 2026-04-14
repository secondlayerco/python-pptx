"""Unit test suite for `pptx.parts.font` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.package import Package
from pptx.parts.font import FontPart

from ..unitutil.mock import initializer_mock, instance_mock


class DescribeFontPart:
    """Unit-test suite for `pptx.parts.font.FontPart` objects."""

    def it_can_construct_from_font_bytes(self, request):
        _init_ = initializer_mock(request, FontPart)
        package_ = instance_mock(request, Package)
        package_.next_partname.return_value = PackURI("/ppt/fonts/font1.fntdata")
        font_bytes = b"\x00\x01\x00\x00fake-font-data"

        font_part = FontPart.new(font_bytes, package_)

        package_.next_partname.assert_called_once_with("/ppt/fonts/font%d.fntdata")
        _init_.assert_called_once_with(
            font_part,
            PackURI("/ppt/fonts/font1.fntdata"),
            CT.X_FONTDATA,
            package_,
            font_bytes,
        )
        assert isinstance(font_part, FontPart)

    def it_stores_the_font_blob(self):
        font_bytes = b"\x00\x01\x00\x00fake-font-data"
        font_part = FontPart(PackURI("/ppt/fonts/font1.fntdata"), CT.X_FONTDATA, None, font_bytes)

        assert font_part.blob == font_bytes

    def it_has_the_correct_content_type(self):
        font_part = FontPart(
            PackURI("/ppt/fonts/font1.fntdata"), CT.X_FONTDATA, None, b"bytes"
        )

        assert font_part.content_type == CT.X_FONTDATA

    def it_has_the_correct_partname(self):
        partname = PackURI("/ppt/fonts/font1.fntdata")
        font_part = FontPart(partname, CT.X_FONTDATA, None, b"bytes")

        assert font_part.partname == partname
