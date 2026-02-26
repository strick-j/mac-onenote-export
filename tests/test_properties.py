"""Tests for onenote_export.parser.properties module."""

from onenote_export.parser.properties import (
    BOLD,
    CACHED_TITLE_STRING,
    CONTENT_CHILD_NODES,
    ELEMENT_CHILD_NODES,
    EMBEDDED_FILE_CONTAINER,
    FONT,
    FONT_COLOR,
    FONT_SIZE,
    HIGHLIGHT,
    IMAGE_ALT_TEXT,
    IMAGE_FILENAME,
    ITALIC,
    JCIDType,
    NUMBER_LIST_FORMAT,
    PICTURE_CONTAINER,
    PICTURE_HEIGHT,
    PICTURE_WIDTH,
    PROPERTY_NAMES,
    PropertyType,
    RICH_EDIT_TEXT_UNICODE,
    ROW_COUNT,
    COLUMN_COUNT,
    STRIKETHROUGH,
    SUBSCRIPT,
    SUPERSCRIPT,
    TABLE_BORDERS_VISIBLE,
    TABLE_COLUMN_WIDTHS,
    TEXT_EXTENDED_ASCII,
    UNDERLINE,
    WZ_HYPERLINK_URL,
    property_index,
    property_type,
)


class TestPropertyType:
    """Tests for PropertyType enum."""

    def test_no_data(self):
        assert PropertyType.NO_DATA == 0x01

    def test_bool(self):
        assert PropertyType.BOOL == 0x02

    def test_one_byte(self):
        assert PropertyType.ONE_BYTE == 0x03

    def test_two_bytes(self):
        assert PropertyType.TWO_BYTES == 0x04

    def test_four_bytes(self):
        assert PropertyType.FOUR_BYTES == 0x05

    def test_eight_bytes(self):
        assert PropertyType.EIGHT_BYTES == 0x06

    def test_four_bytes_length_data(self):
        assert PropertyType.FOUR_BYTES_OF_LENGTH_FOLLOWED_BY_DATA == 0x07

    def test_object_id(self):
        assert PropertyType.OBJECT_ID == 0x08

    def test_array_of_object_ids(self):
        assert PropertyType.ARRAY_OF_OBJECT_IDS == 0x09

    def test_object_space_id(self):
        assert PropertyType.OBJECT_SPACE_ID == 0x0A

    def test_array_of_object_space_ids(self):
        assert PropertyType.ARRAY_OF_OBJECT_SPACE_IDS == 0x0B

    def test_context_id(self):
        assert PropertyType.CONTEXT_ID == 0x0C

    def test_array_of_context_ids(self):
        assert PropertyType.ARRAY_OF_CONTEXT_IDS == 0x0D

    def test_array_of_property_values(self):
        assert PropertyType.ARRAY_OF_PROPERTY_VALUES == 0x10

    def test_enum_member_count(self):
        assert len(PropertyType) == 14


class TestJCIDType:
    """Tests for JCIDType enum."""

    def test_page_node(self):
        assert JCIDType.PAGE_NODE == 0x0006000B

    def test_outline_node(self):
        assert JCIDType.OUTLINE_NODE == 0x0006000C

    def test_outline_element_node(self):
        assert JCIDType.OUTLINE_ELEMENT_NODE == 0x0006000D

    def test_rich_text_oe_node(self):
        assert JCIDType.RICH_TEXT_OE_NODE == 0x0006000E

    def test_image_node(self):
        assert JCIDType.IMAGE_NODE == 0x00060011

    def test_table_node(self):
        assert JCIDType.TABLE_NODE == 0x00060022

    def test_table_row_node(self):
        assert JCIDType.TABLE_ROW_NODE == 0x00060023

    def test_table_cell_node(self):
        assert JCIDType.TABLE_CELL_NODE == 0x00060024

    def test_embedded_file_node(self):
        assert JCIDType.EMBEDDED_FILE_NODE == 0x00060035

    def test_section_node(self):
        assert JCIDType.SECTION_NODE == 0x00060007

    def test_page_series_node(self):
        assert JCIDType.PAGE_SERIES_NODE == 0x00060008

    def test_page_meta_data(self):
        assert JCIDType.PAGE_META_DATA == 0x00020030

    def test_section_meta_data(self):
        assert JCIDType.SECTION_META_DATA == 0x00020031

    def test_number_list_node(self):
        assert JCIDType.NUMBER_LIST_NODE == 0x00060012

    def test_title_node(self):
        assert JCIDType.TITLE_NODE == 0x0006002C

    def test_paragraph_style_object(self):
        assert JCIDType.PARAGRAPH_STYLE_OBJECT == 0x0012004D

    def test_enum_member_count(self):
        assert len(JCIDType) == 26


class TestPropertyTypeExtraction:
    """Tests for property_type() helper function."""

    def test_bold_is_bool(self):
        # BOLD = 0x08001C04 -> type bits 26-30 = 0x02 (BOOL)
        assert property_type(BOLD) == PropertyType.BOOL

    def test_italic_is_bool(self):
        assert property_type(ITALIC) == PropertyType.BOOL

    def test_font_is_four_bytes_length_data(self):
        # FONT = 0x1C001C0A -> type bits = 0x07
        assert property_type(FONT) == PropertyType.FOUR_BYTES_OF_LENGTH_FOLLOWED_BY_DATA

    def test_font_size_is_four_bytes(self):
        # FONT_SIZE = 0x10001C0B -> type bits = 0x04
        assert property_type(FONT_SIZE) == PropertyType.TWO_BYTES

    def test_font_color_is_four_bytes(self):
        # FONT_COLOR = 0x14001C0C -> type bits = 0x05
        assert property_type(FONT_COLOR) == PropertyType.FOUR_BYTES

    def test_content_child_nodes_is_array_of_object_ids(self):
        # CONTENT_CHILD_NODES = 0x24001C1F -> type bits = 0x09
        assert property_type(CONTENT_CHILD_NODES) == PropertyType.ARRAY_OF_OBJECT_IDS

    def test_picture_container_is_object_id(self):
        # PICTURE_CONTAINER = 0x20001C3F -> type bits = 0x08
        assert property_type(PICTURE_CONTAINER) == PropertyType.OBJECT_ID

    def test_rich_edit_text_unicode(self):
        assert property_type(RICH_EDIT_TEXT_UNICODE) == PropertyType.FOUR_BYTES_OF_LENGTH_FOLLOWED_BY_DATA

    def test_row_count_is_four_bytes(self):
        assert property_type(ROW_COUNT) == PropertyType.FOUR_BYTES


class TestPropertyIndexExtraction:
    """Tests for property_index() helper function."""

    def test_bold_index(self):
        # BOLD = 0x08001C04 -> index bits 0-25 = 0x001C04
        assert property_index(BOLD) == 0x001C04

    def test_font_index(self):
        assert property_index(FONT) == 0x001C0A

    def test_row_count_index(self):
        assert property_index(ROW_COUNT) == 0x001D57

    def test_zero_property(self):
        assert property_index(0) == 0

    def test_round_trip(self):
        """Verify type and index can reconstruct the original property ID."""
        for prop_id in [BOLD, FONT, ROW_COUNT, PICTURE_CONTAINER]:
            ptype = property_type(prop_id)
            pidx = property_index(prop_id)
            reconstructed = (ptype << 26) | pidx
            assert reconstructed == prop_id


class TestPropertyNames:
    """Tests for PROPERTY_NAMES dict."""

    def test_bold_name(self):
        assert PROPERTY_NAMES[BOLD] == "Bold"

    def test_italic_name(self):
        assert PROPERTY_NAMES[ITALIC] == "Italic"

    def test_underline_name(self):
        assert PROPERTY_NAMES[UNDERLINE] == "Underline"

    def test_strikethrough_name(self):
        assert PROPERTY_NAMES[STRIKETHROUGH] == "Strikethrough"

    def test_font_name(self):
        assert PROPERTY_NAMES[FONT] == "Font"

    def test_font_size_name(self):
        assert PROPERTY_NAMES[FONT_SIZE] == "FontSize"

    def test_rich_edit_text_unicode_name(self):
        assert PROPERTY_NAMES[RICH_EDIT_TEXT_UNICODE] == "RichEditTextUnicode"

    def test_text_extended_ascii_name(self):
        assert PROPERTY_NAMES[TEXT_EXTENDED_ASCII] == "TextExtendedAscii"

    def test_picture_container_name(self):
        assert PROPERTY_NAMES[PICTURE_CONTAINER] == "PictureContainer"

    def test_image_filename_name(self):
        assert PROPERTY_NAMES[IMAGE_FILENAME] == "ImageFilename"

    def test_row_count_name(self):
        assert PROPERTY_NAMES[ROW_COUNT] == "RowCount"

    def test_column_count_name(self):
        assert PROPERTY_NAMES[COLUMN_COUNT] == "ColumnCount"

    def test_hyperlink_url_name(self):
        assert PROPERTY_NAMES[WZ_HYPERLINK_URL] == "WzHyperlinkUrl"

    def test_has_expected_entry_count(self):
        # Verify the dict is not empty and has a reasonable number of entries
        assert len(PROPERTY_NAMES) >= 40

    def test_all_values_are_strings(self):
        for key, val in PROPERTY_NAMES.items():
            assert isinstance(key, int), f"Key {key} should be int"
            assert isinstance(val, str), f"Value for {key} should be str"
            assert val, f"Name for {key} should not be empty"


class TestPropertyConstants:
    """Tests for property constant values."""

    def test_formatting_constants_exist(self):
        assert BOLD == 0x08001C04
        assert ITALIC == 0x08001C05
        assert UNDERLINE == 0x08001C06
        assert STRIKETHROUGH == 0x08001C07
        assert SUPERSCRIPT == 0x08001C08
        assert SUBSCRIPT == 0x08001C09

    def test_content_constants_exist(self):
        assert RICH_EDIT_TEXT_UNICODE == 0x1C001C22
        assert TEXT_EXTENDED_ASCII == 0x1C003498

    def test_table_constants_exist(self):
        assert ROW_COUNT == 0x14001D57
        assert COLUMN_COUNT == 0x14001D58
        assert TABLE_BORDERS_VISIBLE == 0x08001D5E
        assert TABLE_COLUMN_WIDTHS == 0x1C001D66

    def test_image_constants_exist(self):
        assert PICTURE_CONTAINER == 0x20001C3F
        assert IMAGE_FILENAME == 0x1C001DD7
        assert IMAGE_ALT_TEXT == 0x1C001E58
        assert PICTURE_WIDTH == 0x140034CD
        assert PICTURE_HEIGHT == 0x140034CE

    def test_embedded_file_constants_exist(self):
        assert EMBEDDED_FILE_CONTAINER == 0x20001D9B

    def test_layout_constants_exist(self):
        assert CONTENT_CHILD_NODES == 0x24001C1F
        assert ELEMENT_CHILD_NODES == 0x24001C20
        assert NUMBER_LIST_FORMAT == 0x1C001C1A

    def test_misc_constants_exist(self):
        assert HIGHLIGHT == 0x14001C0D
        assert FONT_COLOR == 0x14001C0C
        assert CACHED_TITLE_STRING == 0x1C001CF3
