"""Tests for onenote_export.parser.content_extractor module."""

from onenote_export.parser.content_extractor import (
    _as_bool,
    _build_top_level_oe_ids,
    _clean_text,
    _decode_text_value,
    _deduplicate_objects,
    _dedup_elements,
    _detect_image_format,
    _extract_rich_text,
    _looks_garbled,
    _parse_byte_prop_as_int,
    _parse_font_size,
    _parse_int_prop,
    _section_name_from_path,
)
from onenote_export.model.content import RichText, TextRun
from onenote_export.parser.one_store import ExtractedObject


class TestDecodeTextValue:
    """Tests for _decode_text_value."""

    def test_plain_string(self):
        assert _decode_text_value("Hello world") == "Hello world"

    def test_empty_string(self):
        assert _decode_text_value("") == ""

    def test_none_value(self):
        assert _decode_text_value(None) == ""

    def test_bytes_unicode(self):
        text = "Hello"
        raw = text.encode("utf-16-le")
        result = _decode_text_value(raw, encoding="unicode")
        assert result == "Hello"

    def test_bytes_ascii(self):
        raw = b"Hello\x00"
        result = _decode_text_value(raw, encoding="ascii")
        assert result == "Hello"

    def test_hex_string_ascii(self):
        raw = "48656c6c6f"  # "Hello" in hex
        result = _decode_text_value(raw, encoding="ascii")
        assert result == "Hello"

    def test_string_with_null_bytes(self):
        result = _decode_text_value("Hello\x00World")
        assert "\x00" not in result


class TestLooksGarbled:
    """Tests for _looks_garbled."""

    def test_normal_text(self):
        assert _looks_garbled("Hello world") is False

    def test_empty_text(self):
        assert _looks_garbled("") is False

    def test_short_text(self):
        assert _looks_garbled("ab") is False

    def test_garbled_text(self):
        # Simulate CJK characters from misinterpreted ASCII
        garbled = "\u4e48\u5f00\u53d1"
        assert _looks_garbled(garbled) is True


class TestCleanText:
    """Tests for _clean_text."""

    def test_removes_null_bytes(self):
        assert _clean_text("hello\x00world") == "helloworld"

    def test_strips_whitespace(self):
        assert _clean_text("  hello  ") == "hello"

    def test_empty_string(self):
        assert _clean_text("") == ""


class TestAsBool:
    """Tests for _as_bool."""

    def test_true_bool(self):
        assert _as_bool(True) is True

    def test_false_bool(self):
        assert _as_bool(False) is False

    def test_string_true(self):
        assert _as_bool("true") is True
        assert _as_bool("True") is True
        assert _as_bool("1") is True
        assert _as_bool("yes") is True

    def test_string_false(self):
        assert _as_bool("false") is False
        assert _as_bool("no") is False
        assert _as_bool("") is False

    def test_integer(self):
        assert _as_bool(1) is True
        assert _as_bool(0) is False

    def test_none(self):
        assert _as_bool(None) is False


class TestParseIntProp:
    """Tests for _parse_int_prop."""

    def test_int_value(self):
        assert _parse_int_prop(42) == 42

    def test_zero(self):
        assert _parse_int_prop(0) == 0

    def test_bytes_value(self):
        raw = (100).to_bytes(4, "little")
        assert _parse_int_prop(raw) == 100

    def test_string_with_number(self):
        assert _parse_int_prop("size: 12pt") == 12

    def test_string_no_number(self):
        assert _parse_int_prop("none") == 0

    def test_none(self):
        assert _parse_int_prop(None) == 0


class TestParseFontSize:
    """Tests for _parse_font_size."""

    def test_int_value(self):
        assert _parse_font_size(11) == 11

    def test_bytes_value(self):
        raw = (14).to_bytes(2, "little")
        assert _parse_font_size(raw) == 14

    def test_string_value(self):
        assert _parse_font_size("12") == 12

    def test_zero(self):
        assert _parse_font_size(0) == 0


class TestDetectImageFormat:
    """Tests for _detect_image_format."""

    def test_png(self):
        data = b"\x89PNG\r\n\x1a\n" + b"\x00" * 100
        assert _detect_image_format(data) == "png"

    def test_jpeg(self):
        data = b"\xff\xd8\xff\xe0" + b"\x00" * 100
        assert _detect_image_format(data) == "jpeg"

    def test_gif87a(self):
        data = b"GIF87a" + b"\x00" * 100
        assert _detect_image_format(data) == "gif"

    def test_gif89a(self):
        data = b"GIF89a" + b"\x00" * 100
        assert _detect_image_format(data) == "gif"

    def test_bmp(self):
        data = b"BM" + b"\x00" * 100
        assert _detect_image_format(data) == "bmp"

    def test_webp(self):
        data = b"RIFF\x00\x00\x00\x00WEBP" + b"\x00" * 100
        assert _detect_image_format(data) == "webp"

    def test_unknown_format(self):
        data = b"\x00\x01\x02\x03" + b"\x00" * 100
        assert _detect_image_format(data) == ""

    def test_empty_data(self):
        assert _detect_image_format(b"") == ""

    def test_too_short(self):
        assert _detect_image_format(b"\x89PN") == ""


class TestDeduplicateObjects:
    """Tests for _deduplicate_objects."""

    def test_fewer_than_4_objects_returned_unchanged(self):
        objs = [
            ExtractedObject(obj_type="jcidRichTextOENode", identity="1",
                            properties={"RichEditTextUnicode": "Hello"}),
            ExtractedObject(obj_type="jcidRichTextOENode", identity="2",
                            properties={"RichEditTextUnicode": "World"}),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 2

    def test_no_content_fingerprints_returns_unchanged(self):
        """Objects with no content (styles, outlines) are returned as-is."""
        objs = [
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="1"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="2"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="3"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="4"),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 4

    def test_single_content_fp_returns_unchanged(self):
        """Only one content fingerprint means no duplicates possible."""
        objs = [
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="1"),
            ExtractedObject(obj_type="jcidRichTextOENode", identity="2",
                            properties={"RichEditTextUnicode": "Hello"}),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="3"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="4"),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 4

    def test_duplicates_removed(self):
        """Duplicate content objects should be deduplicated."""
        objs = [
            ExtractedObject(obj_type="jcidRichTextOENode", identity="1",
                            properties={"RichEditTextUnicode": "Hello"}),
            ExtractedObject(obj_type="jcidRichTextOENode", identity="2",
                            properties={"RichEditTextUnicode": "World"}),
            ExtractedObject(obj_type="jcidRichTextOENode", identity="3",
                            properties={"RichEditTextUnicode": "Hello"}),
            ExtractedObject(obj_type="jcidRichTextOENode", identity="4",
                            properties={"RichEditTextUnicode": "World"}),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 2

    def test_empty_list(self):
        assert _deduplicate_objects([]) == []


class TestBuildTopLevelOeIds:
    """Tests for _build_top_level_oe_ids."""

    def test_collects_outline_node_children(self):
        objs = [
            ExtractedObject(
                obj_type="jcidOutlineNode", identity="outline-1",
                properties={"ElementChildNodesOfVersionHistory": ["oe-1", "oe-2"]},
            ),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == {"oe-1", "oe-2"}

    def test_string_ref_normalized_to_list(self):
        """A single string ref should be handled as a list."""
        objs = [
            ExtractedObject(
                obj_type="jcidOutlineNode", identity="outline-1",
                properties={"ElementChildNodesOfVersionHistory": "oe-single"},
            ),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == {"oe-single"}

    def test_no_outline_nodes(self):
        objs = [
            ExtractedObject(obj_type="jcidRichTextOENode", identity="1"),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == set()

    def test_empty_refs(self):
        objs = [
            ExtractedObject(
                obj_type="jcidOutlineNode", identity="outline-1",
                properties={"ElementChildNodesOfVersionHistory": []},
            ),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == set()


class TestDecodeTextEdgeCases:
    """Additional edge case tests for _decode_text_value."""

    def test_hex_string_unicode_encoding(self):
        """Hex string decoded as UTF-16LE."""
        # "Hi" in UTF-16LE = 48 00 69 00
        result = _decode_text_value("48006900", encoding="unicode")
        assert result == "Hi"

    def test_garbled_ascii_re_encoding(self):
        """Garbled CJK text re-encoded from UTF-16LE to ASCII."""
        # Create text that looks garbled (> 30% non-ASCII)
        ascii_text = "Hello World"
        # Encode as UTF-16LE, then decode incorrectly as UTF-16LE to get garbled
        garbled = ascii_text.encode("ascii").decode("utf-16-le", errors="replace")
        result = _decode_text_value(garbled, encoding="ascii")
        # Should attempt to re-encode and recover
        assert isinstance(result, str)

    def test_bytes_utf16_decode_error_falls_back_to_latin1(self):
        """Invalid UTF-16LE bytes fall back to latin-1 decoding."""
        # Odd-length bytes can't be valid UTF-16LE
        raw = b"\xff\xfe\x80"
        result = _decode_text_value(raw, encoding="unicode")
        assert isinstance(result, str)

    def test_integer_value(self):
        """Non-string, non-bytes values are converted via str()."""
        result = _decode_text_value(42)
        assert result == "42"

    def test_whitespace_only_string(self):
        result = _decode_text_value("   ")
        assert result == ""


class TestParseBytePropAsInt:
    """Tests for _parse_byte_prop_as_int."""

    def test_int_value(self):
        assert _parse_byte_prop_as_int(36) == 36

    def test_bytes_value(self):
        raw = (36).to_bytes(2, "little")
        assert _parse_byte_prop_as_int(raw) == 36

    def test_repr_string(self):
        """Parse repr() style byte string like b'$\\x00'."""
        assert _parse_byte_prop_as_int("b'$\\x00'") == 36  # ord('$') = 36

    def test_repr_string_hex(self):
        assert _parse_byte_prop_as_int("b'\\x04\\x00'") == 4

    def test_empty_string(self):
        assert _parse_byte_prop_as_int("") == 0

    def test_none(self):
        assert _parse_byte_prop_as_int(None) == 0

    def test_invalid_repr(self):
        assert _parse_byte_prop_as_int("b'invalid") == 0

    def test_short_bytes(self):
        """Single byte should return 0 (need at least 2)."""
        assert _parse_byte_prop_as_int(b"\x05") == 0


class TestSectionNameFromPath:
    """Tests for _section_name_from_path."""

    def test_simple_filename(self):
        assert _section_name_from_path("/path/to/Notes.one") == "Notes"

    def test_filename_with_date(self):
        result = _section_name_from_path("/path/to/ADI (On 2-25-26).one")
        assert result == "ADI"

    def test_filename_with_date_dash_suffix(self):
        result = _section_name_from_path("/path/to/Section (On 2-25-26 - 3).one")
        assert result == "Section"

    def test_empty_after_cleaning(self):
        """If name becomes empty after cleaning, return 'Untitled'."""
        result = _section_name_from_path("/path/to/.one")
        assert result == "Untitled"


class TestParseFontSizeEdgeCases:
    """Additional edge case tests for _parse_font_size."""

    def test_none_value(self):
        assert _parse_font_size(None) == 0

    def test_string_no_number(self):
        assert _parse_font_size("normal") == 0

    def test_single_byte(self):
        raw = b"\x0e"  # 14
        assert _parse_font_size(raw) == 14

    def test_empty_bytes(self):
        assert _parse_font_size(b"") == 0


class TestDedupElements:
    """Tests for _dedup_elements."""

    def test_removes_duplicate_rich_text(self):
        run1 = TextRun(text="Hello world")
        elem1 = RichText(runs=[run1])
        elem2 = RichText(runs=[run1])
        result = _dedup_elements([elem1, elem2])
        assert len(result) == 1

    def test_keeps_same_text_different_list_type(self):
        """Same text in different list types is not a duplicate."""
        run = TextRun(text="Item")
        bullet = RichText(runs=[run], list_type="unordered")
        numbered = RichText(runs=[run], list_type="ordered")
        result = _dedup_elements([bullet, numbered])
        assert len(result) == 2

    def test_empty_list(self):
        assert _dedup_elements([]) == []


class TestExtractRichText:
    """Tests for _extract_rich_text."""

    def test_returns_none_for_empty_text(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode", identity="1",
            properties={},
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is None

    def test_returns_none_for_whitespace_only(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode", identity="1",
            properties={"RichEditTextUnicode": "   "},
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is None

    def test_extracts_text_with_formatting(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode", identity="1",
            properties={"RichEditTextUnicode": "Hello"},
        )
        style = {"Bold": True, "Italic": True}
        result = _extract_rich_text(obj, style, None)
        assert result is not None
        assert result.runs[0].text == "Hello"
        assert result.runs[0].bold is True
        assert result.runs[0].italic is True

    def test_extracts_text_from_ascii(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode", identity="1",
            properties={"TextExtendedAscii": "World"},
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is not None
        assert result.runs[0].text == "World"
