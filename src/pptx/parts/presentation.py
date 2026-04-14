"""Presentation part, the main part in a .pptx package."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterable

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.eot import ttf_to_eot
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.parts.font import FontPart
from pptx.parts.slide import NotesMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster


class PresentationPart(XmlPart):
    """Top level class in object model.

    Represents the contents of the /ppt directory of a .pptx file.
    """

    def add_slide(self, slide_layout: SlideLayout):
        """Return (rId, slide) pair of a newly created blank slide.

        New slide inherits appearance from `slide_layout`.
        """
        partname = self._next_slide_partname
        slide_layout_part = slide_layout.part
        slide_part = SlidePart.new(partname, self.package, slide_layout_part)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    @property
    def core_properties(self) -> CorePropertiesPart:
        """A |CoreProperties| object for the presentation.

        Provides read/write access to the Dublin Core properties of this presentation.
        """
        return self.package.core_properties

    def get_slide(self, slide_id: int) -> Slide | None:
        """Return optional related |Slide| object identified by `slide_id`.

        Returns |None| if no slide with `slide_id` is related to this presentation.
        """
        for sldId in self._element.sldIdLst:
            if sldId.id == slide_id:
                return self.related_part(sldId.rId).slide
        return None

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """
        Return the |NotesMaster| object for this presentation. If the
        presentation does not have a notes master, one is created from
        a default template. The same single instance is returned on each
        call.
        """
        return self.notes_master_part.notes_master

    @lazyproperty
    def notes_master_part(self) -> NotesMasterPart:
        """Return the |NotesMasterPart| object for this presentation.

        If the presentation does not have a notes master, one is created from a default template.
        The same single instance is returned on each call.
        """
        try:
            return self.part_related_by(RT.NOTES_MASTER)
        except KeyError:
            notes_master_part = NotesMasterPart.create_default(self.package)
            self.relate_to(notes_master_part, RT.NOTES_MASTER)
            return notes_master_part

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    def related_slide(self, rId: str) -> Slide:
        """Return |Slide| object for related |SlidePart| related by `rId`."""
        return self.related_part(rId).slide

    def related_slide_master(self, rId: str) -> SlideMaster:
        """Return |SlideMaster| object for |SlideMasterPart| related by `rId`."""
        return self.related_part(rId).slide_master

    def rename_slide_parts(self, rIds: Iterable[str]):
        """Assign incrementing partnames to the slide parts identified by `rIds`.

        Partnames are like `/ppt/slides/slide9.xml` and are assigned in the order their id appears
        in the `rIds` sequence. The name portion is always `slide`. The number part forms a
        continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...). The extension is always
        `.xml`.
        """
        for idx, rId in enumerate(rIds):
            slide_part = self.related_part(rId)
            slide_part.partname = PackURI("/ppt/slides/slide%d.xml" % (idx + 1))

    def embed_font(
        self, typeface: str, font_bytes: bytes, bold: bool = False, italic: bool = False
    ) -> None:
        """Embed a font file into the presentation.

        Args:
            typeface: The font family name (must match what's used in run.font.name).
            font_bytes: Raw TTF/OTF file bytes.
            bold: True if this is the bold variant.
            italic: True if this is the italic variant.
        """
        # 1. Ensure embedTrueTypeFonts is enabled on the presentation element
        self._element.set("embedTrueTypeFonts", "1")

        # 2. Convert TTF/OTF to EOT format and create font part
        eot_bytes = ttf_to_eot(font_bytes)
        font_part = FontPart.new(eot_bytes, self.package)
        rId = self.relate_to(font_part, RT.FONT)

        # 3. Get or create <p:embeddedFontLst>
        embeddedFontLst = self._element.get_or_add_embeddedFontLst()

        # 4. Find or create <p:embeddedFont> for this typeface
        embeddedFont = self._get_or_add_embedded_font(embeddedFontLst, typeface)

        # 5. Add the style variant sub-element with the rId
        style_tag = self._font_style_tag(bold, italic)
        style_elm = OxmlElement(style_tag)
        style_elm.set(qn("r:id"), rId)
        embeddedFont.append(style_elm)

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this presentation package to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        self.package.save(path_or_stream)

    def slide_id(self, slide_part):
        """Return the slide-id associated with `slide_part`."""
        for sldId in self._element.sldIdLst:
            if self.related_part(sldId.rId) is slide_part:
                return sldId.id
        raise ValueError("matching slide_part not found")

    @staticmethod
    def _font_style_tag(bold: bool, italic: bool) -> str:
        """Return the element tag for the font style variant."""
        if bold and italic:
            return "p:boldItalic"
        if bold:
            return "p:bold"
        if italic:
            return "p:italic"
        return "p:regular"

    @staticmethod
    def _get_or_add_embedded_font(lst, typeface: str):
        """Find existing <p:embeddedFont> for typeface, or create one."""
        for ef in lst.findall(qn("p:embeddedFont")):
            font_elm = ef.find(qn("p:font"))
            if font_elm is not None and font_elm.get("typeface") == typeface:
                return ef
        # Create new
        ef = OxmlElement("p:embeddedFont")
        font_elm = OxmlElement("p:font")
        font_elm.set("typeface", typeface)
        font_elm.set("pitchFamily", "2")
        font_elm.set("charset", "0")
        ef.append(font_elm)
        lst.append(ef)
        return ef

    @property
    def _next_slide_partname(self):
        """Return |PackURI| instance containing next available slide partname."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        partname_str = "/ppt/slides/slide%d.xml" % (len(sldIdLst) + 1)
        return PackURI(partname_str)
