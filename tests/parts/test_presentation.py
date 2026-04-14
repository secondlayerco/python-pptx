"""Unit-test suite for `pptx.parts.presentation` module."""

from __future__ import annotations

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.package import Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.font import FontPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import NotesMasterPart, SlideMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster

from ..unitutil.cxml import element
from ..unitutil.mock import (
    call,
    class_mock,
    function_mock,
    instance_mock,
    method_mock,
    property_mock,
)


class DescribePresentationPart(object):
    """Unit-test suite for `pptx.parts.presentation.PresentationPart` objects."""

    def it_provides_access_to_its_presentation(self, request):
        prs_ = instance_mock(request, Presentation)
        Presentation_ = class_mock(
            request, "pptx.parts.presentation.Presentation", return_value=prs_
        )
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, None, prs_elm)

        prs = prs_part.presentation

        Presentation_.assert_called_once_with(prs_elm, prs_part)
        assert prs is prs_

    def it_provides_access_to_its_core_properties(self, request, package_):
        core_properties_ = instance_mock(request, CorePropertiesPart)
        package_.core_properties = core_properties_
        prs_part = PresentationPart(None, None, package_, None)

        assert prs_part.core_properties is core_properties_

    def it_provides_access_to_an_existing_notes_master_part(
        self, notes_master_part_, part_related_by_
    ):
        """This is the first of a two-part test to cover the existing notes master case.

        The notes master not-present case follows.
        """
        prs_part = PresentationPart(None, None, None, None)
        part_related_by_.return_value = notes_master_part_

        notes_master_part = prs_part.notes_master_part

        prs_part.part_related_by.assert_called_once_with(prs_part, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def but_it_adds_a_notes_master_part_when_needed(
        self, request, package_, notes_master_part_, part_related_by_, relate_to_
    ):
        """This is the second of a two-part test to cover notes-master-not-present case.

        The notes master present case is just above.
        """
        NotesMasterPart_ = class_mock(request, "pptx.parts.presentation.NotesMasterPart")
        NotesMasterPart_.create_default.return_value = notes_master_part_
        part_related_by_.side_effect = KeyError
        prs_part = PresentationPart(None, None, package_, None)

        notes_master_part = prs_part.notes_master_part

        NotesMasterPart_.create_default.assert_called_once_with(package_)
        relate_to_.assert_called_once_with(prs_part, notes_master_part_, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def it_provides_access_to_its_notes_master(self, request, notes_master_part_):
        notes_master_ = instance_mock(request, NotesMaster)
        property_mock(
            request,
            PresentationPart,
            "notes_master_part",
            return_value=notes_master_part_,
        )
        notes_master_part_.notes_master = notes_master_
        prs_part = PresentationPart(None, None, None, None)

        assert prs_part.notes_master is notes_master_

    def it_provides_access_to_a_related_slide(self, request, slide_, related_part_):
        slide_part_ = instance_mock(request, SlidePart, slide=slide_)
        related_part_.return_value = slide_part_
        prs_part = PresentationPart(None, None, None, None)

        slide = prs_part.related_slide("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide is slide_

    def it_provides_access_to_a_related_master(self, request, slide_master_, related_part_):
        slide_master_part_ = instance_mock(request, SlideMasterPart, slide_master=slide_master_)
        related_part_.return_value = slide_master_part_
        prs_part = PresentationPart(None, None, None, None)

        slide_master = prs_part.related_slide_master("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide_master is slide_master_

    def it_can_rename_related_slide_parts(self, request, related_part_):
        rIds = tuple("rId%d" % n for n in range(5, 0, -1))
        slide_parts = tuple(instance_mock(request, SlidePart) for _ in range(5))
        related_part_.side_effect = iter(slide_parts)
        prs_part = PresentationPart(None, None, None, None)

        prs_part.rename_slide_parts(rIds)

        assert related_part_.call_args_list == [call(prs_part, rId) for rId in rIds]
        assert [s.partname for s in slide_parts] == [
            PackURI("/ppt/slides/slide%d.xml" % (i + 1)) for i in range(len(rIds))
        ]

    def it_can_save_the_package_to_a_file(self, package_):
        PresentationPart(None, None, package_, None).save("prs.pptx")
        package_.save.assert_called_once_with("prs.pptx")

    def it_can_add_a_new_slide(self, request, package_, slide_part_, slide_, relate_to_):
        slide_layout_ = instance_mock(request, SlideLayout)
        partname = PackURI("/ppt/slides/slide9.xml")
        property_mock(request, PresentationPart, "_next_slide_partname", return_value=partname)
        SlidePart_ = class_mock(request, "pptx.parts.presentation.SlidePart")
        SlidePart_.new.return_value = slide_part_
        relate_to_.return_value = "rId42"
        slide_layout_part_ = slide_layout_.part
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, package_, None)

        rId, slide = prs_part.add_slide(slide_layout_)

        SlidePart_.new.assert_called_once_with(partname, package_, slide_layout_part_)
        prs_part.relate_to.assert_called_once_with(prs_part, slide_part_, RT.SLIDE)
        assert rId == "rId42"
        assert slide is slide_

    def it_finds_the_slide_id_of_a_slide_part(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.side_effect = iter((None, slide_part_, None))
        prs_part = PresentationPart(None, None, None, prs_elm)

        _slide_id = prs_part.slide_id(slide_part_)

        assert related_part_.call_args_list == [
            call(prs_part, "a"),
            call(prs_part, "b"),
        ]
        assert _slide_id == 257

    def it_raises_on_slide_id_not_found(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.return_value = "not the slide you're looking for"
        prs_part = PresentationPart(None, None, None, prs_elm)

        with pytest.raises(ValueError):
            prs_part.slide_id(slide_part_)

    @pytest.mark.parametrize("is_present", (True, False))
    def it_finds_a_slide_by_slide_id(self, is_present, slide_, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        slide_id = 257 if is_present else 666
        expected_value = slide_ if is_present else None
        related_part_.return_value = slide_part_
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, None, prs_elm)

        slide = prs_part.get_slide(slide_id)

        assert slide == expected_value

    def it_knows_the_next_slide_partname_to_help(self):
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId,p:sldId)")
        prs_part = PresentationPart(None, None, None, prs_elm)

        assert prs_part._next_slide_partname == PackURI("/ppt/slides/slide3.xml")

    def it_can_embed_a_font(self, request, package_, relate_to_):
        font_bytes = b"\x00\x01\x00\x00fake-font-data"
        eot_bytes = b"fake-eot-data"
        ttf_to_eot_ = function_mock(
            request, "pptx.parts.presentation.ttf_to_eot", return_value=eot_bytes
        )
        font_part_ = instance_mock(request, FontPart)
        FontPart_ = class_mock(
            request, "pptx.parts.presentation.FontPart", return_value=font_part_
        )
        FontPart_.new.return_value = font_part_
        relate_to_.return_value = "rId10"
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, package_, prs_elm)

        prs_part.embed_font("Beth Ellen Regular", font_bytes)

        ttf_to_eot_.assert_called_once_with(font_bytes)
        FontPart_.new.assert_called_once_with(eot_bytes, package_)
        relate_to_.assert_called_once_with(prs_part, font_part_, RT.FONT)
        # verify embedTrueTypeFonts attribute is set
        assert prs_elm.get("embedTrueTypeFonts") == "1"
        # verify XML structure
        embeddedFontLst = prs_elm.embeddedFontLst
        assert embeddedFontLst is not None
        ef_list = embeddedFontLst.findall(qn("p:embeddedFont"))
        assert len(ef_list) == 1
        font_elm = ef_list[0].find(qn("p:font"))
        assert font_elm is not None
        assert font_elm.get("typeface") == "Beth Ellen Regular"
        regular_elm = ef_list[0].find(qn("p:regular"))
        assert regular_elm is not None
        assert regular_elm.get(qn("r:id")) == "rId10"

    def it_can_embed_bold_and_italic_variants(self, request, package_, relate_to_):
        function_mock(
            request, "pptx.parts.presentation.ttf_to_eot", return_value=b"eot"
        )
        font_part_ = instance_mock(request, FontPart)
        FontPart_ = class_mock(
            request, "pptx.parts.presentation.FontPart", return_value=font_part_
        )
        FontPart_.new.return_value = font_part_
        relate_to_.side_effect = ["rId10", "rId11", "rId12", "rId13"]
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, package_, prs_elm)

        prs_part.embed_font("MyFont", b"regular", bold=False, italic=False)
        prs_part.embed_font("MyFont", b"bold", bold=True, italic=False)
        prs_part.embed_font("MyFont", b"italic", bold=False, italic=True)
        prs_part.embed_font("MyFont", b"bolditalic", bold=True, italic=True)

        # all should be under the same <p:embeddedFont>
        embeddedFontLst = prs_elm.embeddedFontLst
        ef_list = embeddedFontLst.findall(qn("p:embeddedFont"))
        assert len(ef_list) == 1
        ef = ef_list[0]
        assert ef.find(qn("p:regular")).get(qn("r:id")) == "rId10"
        assert ef.find(qn("p:bold")).get(qn("r:id")) == "rId11"
        assert ef.find(qn("p:italic")).get(qn("r:id")) == "rId12"
        assert ef.find(qn("p:boldItalic")).get(qn("r:id")) == "rId13"

    def it_creates_separate_entries_for_different_typefaces(self, request, package_, relate_to_):
        function_mock(
            request, "pptx.parts.presentation.ttf_to_eot", return_value=b"eot"
        )
        font_part_ = instance_mock(request, FontPart)
        FontPart_ = class_mock(
            request, "pptx.parts.presentation.FontPart", return_value=font_part_
        )
        FontPart_.new.return_value = font_part_
        relate_to_.side_effect = ["rId10", "rId11"]
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, package_, prs_elm)

        prs_part.embed_font("FontA", b"a")
        prs_part.embed_font("FontB", b"b")

        ef_list = prs_elm.embeddedFontLst.findall(qn("p:embeddedFont"))
        assert len(ef_list) == 2
        assert ef_list[0].find(qn("p:font")).get("typeface") == "FontA"
        assert ef_list[1].find(qn("p:font")).get("typeface") == "FontB"

    # fixture components ---------------------------------------------

    @pytest.fixture
    def notes_master_part_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request):
        return method_mock(request, PresentationPart, "part_related_by")

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, PresentationPart, "relate_to")

    @pytest.fixture
    def related_part_(self, request):
        return method_mock(request, PresentationPart, "related_part")

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)
