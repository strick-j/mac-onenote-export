"""Tests for onenote_export.parser.one_store module."""

from onenote_export.parser.one_store import (
    ExtractedObject,
    ExtractedPage,
    ExtractedSection,
    _clean_text,
    _extract_guid,
    _parse_int,
)


class TestExtractGuid:
    """Tests for _extract_guid."""

    def test_standard_identity_string(self):
        identity = "<ExtendedGUID> (abc-def-123, 42)"
        assert _extract_guid(identity) == "abc-def-123"

    def test_identity_with_extra_spaces(self):
        identity = "<ExtendedGUID> ( some-guid , 99)"
        assert _extract_guid(identity) == "some-guid"

    def test_empty_string(self):
        assert _extract_guid("") == ""

    def test_no_parentheses(self):
        assert _extract_guid("no-parens-here") == ""

    def test_missing_comma(self):
        assert _extract_guid("(no-comma)") == ""

    def test_complex_guid(self):
        identity = "<ExtendedGUID> ({12345678-abcd-ef01-2345-6789abcdef01}, 138)"
        assert _extract_guid(identity) == "{12345678-abcd-ef01-2345-6789abcdef01}"


class TestCleanText:
    """Tests for _clean_text in one_store module."""

    def test_removes_null_bytes(self):
        assert _clean_text("hello\x00world") == "helloworld"

    def test_strips_whitespace(self):
        assert _clean_text("  hello  ") == "hello"

    def test_empty_string(self):
        assert _clean_text("") == ""

    def test_only_null_bytes(self):
        assert _clean_text("\x00\x00\x00") == ""

    def test_mixed_content(self):
        assert _clean_text("  test\x00data  ") == "testdata"


class TestParseInt:
    """Tests for _parse_int."""

    def test_int_value(self):
        assert _parse_int(42) == 42

    def test_zero(self):
        assert _parse_int(0) == 0

    def test_bytes_value(self):
        raw = (100).to_bytes(4, "little")
        assert _parse_int(raw) == 100

    def test_bytes_short(self):
        raw = (5).to_bytes(2, "little")
        assert _parse_int(raw) == 5

    def test_string_with_number(self):
        assert _parse_int("level: 3") == 3

    def test_string_no_number(self):
        assert _parse_int("none") == 0

    def test_none(self):
        assert _parse_int(None) == 0

    def test_empty_string(self):
        assert _parse_int("") == 0

    def test_negative_int(self):
        assert _parse_int(-1) == -1


class TestExtractedDataclasses:
    """Tests for dataclass constructors."""

    def test_extracted_object_defaults(self):
        obj = ExtractedObject(obj_type="test", identity="id-1")
        assert obj.obj_type == "test"
        assert obj.identity == "id-1"
        assert obj.properties == {}

    def test_extracted_object_with_properties(self):
        props = {"Bold": True, "Font": "Arial"}
        obj = ExtractedObject(obj_type="text", identity="id-2", properties=props)
        assert obj.properties["Bold"] is True
        assert obj.properties["Font"] == "Arial"

    def test_extracted_page_defaults(self):
        page = ExtractedPage()
        assert page.title == ""
        assert page.level == 0
        assert page.author == ""
        assert page.objects == []

    def test_extracted_page_with_values(self):
        page = ExtractedPage(title="My Page", level=2, author="Test User")
        assert page.title == "My Page"
        assert page.level == 2
        assert page.author == "Test User"

    def test_extracted_section_defaults(self):
        section = ExtractedSection()
        assert section.file_path == ""
        assert section.display_name == ""
        assert section.pages == []
        assert section.file_data == {}
        assert section.paragraph_styles == {}

    def test_extracted_section_with_pages(self):
        page = ExtractedPage(title="Test")
        section = ExtractedSection(
            file_path="/test.one",
            display_name="Test Section",
            pages=[page],
        )
        assert section.file_path == "/test.one"
        assert section.display_name == "Test Section"
        assert len(section.pages) == 1
