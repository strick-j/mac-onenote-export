"""Extracts structured content from pyOneNote parsed objects.

Bridges the parser output (ExtractedSection) to the high-level
content model (Section with Pages and ContentElements).
"""

import logging
import re
from pathlib import Path

from onenote_export.model.content import (
    ContentElement,
    EmbeddedFile,
    ImageElement,
    RichText,
    TableElement,
    TextRun,
)
from onenote_export.model.page import Page
from onenote_export.model.section import Section
from onenote_export.parser.one_store import (
    ExtractedObject,
    ExtractedPage,
    ExtractedSection,
)

logger = logging.getLogger(__name__)

# JCID type constants
_RICH_TEXT = "jcidRichTextOENode"
_IMAGE_NODE = "jcidImageNode"
_TABLE_NODE = "jcidTableNode"
_EMBEDDED_FILE = "jcidEmbeddedFileNode"
_OUTLINE_ELEMENT = "jcidOutlineElementNode"
_NUMBER_LIST = "jcidNumberListNode"
_STYLE_CONTAINER = "jcidPersistablePropertyContainerForTOCSection"


def _deduplicate_objects(
    objects: list[ExtractedObject],
) -> list[ExtractedObject]:
    """Remove duplicate content objects caused by OneNote revision history.

    OneNote stores full copies of all page content for each revision.
    This creates repeated blocks of identical text/images. We detect
    the repeat pattern and keep only one copy.
    """
    if len(objects) < 4:
        return objects

    # Build a fingerprint sequence for content-bearing objects
    fingerprints: list[str] = []
    for obj in objects:
        fp = _object_fingerprint(obj)
        fingerprints.append(fp)

    # Find the repeat pattern: look for the first content fingerprint
    # appearing again later in the sequence
    content_fps = [
        (i, fp) for i, fp in enumerate(fingerprints)
        if fp and objects[i].obj_type in (_RICH_TEXT, _IMAGE_NODE, _EMBEDDED_FILE)
    ]

    if len(content_fps) < 2:
        return objects

    # Find where the first content element repeats
    first_fp = content_fps[0][1]
    first_idx = content_fps[0][0]
    repeat_idx = None

    for i, fp in content_fps[1:]:
        if fp == first_fp:
            repeat_idx = i
            break

    if repeat_idx is None:
        return objects

    # Take objects up to the repeat point, plus any non-content objects
    # (styles, outline elements) that follow
    result: list[ExtractedObject] = []
    seen_content: set[str] = set()

    for i, obj in enumerate(objects):
        fp = fingerprints[i]

        # Non-content objects (styles, outlines) - always include
        if not fp:
            result.append(obj)
            continue

        # Content object - include only the first occurrence
        if fp not in seen_content:
            seen_content.add(fp)
            result.append(obj)

    return result


def _object_fingerprint(obj: ExtractedObject) -> str:
    """Create a content-based fingerprint for deduplication."""
    if obj.obj_type == _RICH_TEXT:
        text = str(obj.properties.get("RichEditTextUnicode", ""))
        ascii_text = str(obj.properties.get("TextExtendedAscii", ""))
        return f"text:{text}:{ascii_text}" if (text or ascii_text) else ""
    elif obj.obj_type == _IMAGE_NODE:
        filename = str(obj.properties.get("ImageFilename", ""))
        alt = str(obj.properties.get("ImageAltText", ""))
        return f"img:{filename}:{alt}" if (filename or alt) else ""
    elif obj.obj_type == _EMBEDDED_FILE:
        name = str(obj.properties.get("EmbeddedFileName", ""))
        return f"file:{name}" if name else ""
    return ""


def extract_section(parsed: ExtractedSection) -> Section:
    """Convert an ExtractedSection into a high-level Section model."""
    section = Section(
        name=parsed.display_name or _section_name_from_path(parsed.file_path),
        file_path=parsed.file_path,
    )

    for extracted_page in parsed.pages:
        page = _build_page(extracted_page, parsed.file_data)
        section.pages.append(page)

    return section


def _build_page(
    extracted: ExtractedPage, file_data: dict[str, bytes]
) -> Page:
    """Build a Page model from an ExtractedPage."""
    page = Page(
        title=extracted.title or "Untitled",
        level=extracted.level,
        author=extracted.author,
    )

    # Deduplicate objects: OneNote revisions repeat the full content.
    # Remove objects that are exact duplicates (same type + same text content).
    deduped_objects = _deduplicate_objects(extracted.objects)

    # Process objects in order, building content elements
    # Track the current style context for formatting
    current_style: dict[str, object] = {}
    has_list = False

    for obj in deduped_objects:
        if obj.obj_type == _STYLE_CONTAINER:
            # Style container defines formatting for subsequent text
            current_style = dict(obj.properties)
            continue

        if obj.obj_type == _NUMBER_LIST:
            has_list = True
            continue

        if obj.obj_type == _OUTLINE_ELEMENT:
            # Outline elements contain child level info
            # The actual content is in the next RichTextOENode
            has_list = bool(obj.properties.get("ListNodes"))
            continue

        if obj.obj_type == _RICH_TEXT:
            element = _extract_rich_text(obj, current_style, has_list)
            if element:
                page.elements.append(element)
            has_list = False
            continue

        if obj.obj_type == _IMAGE_NODE:
            element = _extract_image(obj, file_data)
            if element:
                page.elements.append(element)
            continue

        if obj.obj_type == _TABLE_NODE:
            element = _extract_table(obj)
            if element:
                page.elements.append(element)
            continue

        if obj.obj_type == _EMBEDDED_FILE:
            element = _extract_embedded_file(obj, file_data)
            if element:
                page.elements.append(element)
            continue

    return page


def _extract_rich_text(
    obj: ExtractedObject,
    style: dict[str, object],
    is_list_item: bool,
) -> RichText | None:
    """Extract rich text from a RichTextOENode."""
    props = obj.properties

    # Get text content - try RichEditTextUnicode first, then TextExtendedAscii
    text = ""
    raw_unicode = props.get("RichEditTextUnicode", "")
    raw_ascii = props.get("TextExtendedAscii", "")

    if raw_unicode:
        text = _decode_text_value(raw_unicode, encoding="unicode")
    elif raw_ascii:
        text = _decode_text_value(raw_ascii, encoding="ascii")

    if not text or not text.strip():
        return None

    # Get formatting from the style context
    bold = _as_bool(style.get("Bold", False))
    italic = _as_bool(style.get("Italic", False))
    underline = _as_bool(style.get("Underline", False))
    strikethrough = _as_bool(style.get("Strikethrough", False))
    superscript = _as_bool(style.get("Superscript", False))
    subscript = _as_bool(style.get("Subscript", False))
    font = _clean_text(str(style.get("Font", "")))
    font_size = _parse_font_size(style.get("FontSize", 0))

    # Check for hyperlink
    hyperlink_url = _clean_text(str(props.get("WzHyperlinkUrl", "")))

    # Check if this is title text
    is_title = _as_bool(props.get("IsTitleText", False))

    # Build text run
    run = TextRun(
        text=text,
        bold=bold,
        italic=italic,
        underline=underline,
        strikethrough=strikethrough,
        superscript=superscript,
        subscript=subscript,
        font=font,
        font_size=font_size,
        hyperlink_url=hyperlink_url,
    )

    indent_level = 1 if is_list_item else 0

    return RichText(
        runs=[run],
        indent_level=indent_level,
        is_title=is_title,
    )


def _extract_image(
    obj: ExtractedObject, file_data: dict[str, bytes]
) -> ImageElement | None:
    """Extract image from an ImageNode."""
    props = obj.properties

    filename = _clean_text(str(props.get("ImageFilename", "")))
    alt_text = _clean_text(str(props.get("ImageAltText", "")))
    width = _parse_int_prop(props.get("PictureWidth", 0))
    height = _parse_int_prop(props.get("PictureHeight", 0))

    # Try to find image data
    data = b""
    pic_container = props.get("PictureContainer")
    if isinstance(pic_container, bytes):
        data = pic_container
    elif isinstance(pic_container, list) and file_data:
        # Reference to file data store
        for key, content in file_data.items():
            fmt = _detect_image_format(content)
            if fmt:
                data = content
                break

    if not data and not filename:
        return None

    fmt = _detect_image_format(data) if data else ""

    return ImageElement(
        data=data,
        filename=filename or f"image.{fmt or 'bin'}",
        alt_text=alt_text,
        width=width,
        height=height,
        format=fmt,
    )


def _extract_table(obj: ExtractedObject) -> TableElement | None:
    """Extract table structure from a TableNode."""
    props = obj.properties
    row_count = _parse_int_prop(props.get("RowCount", 0))
    col_count = _parse_int_prop(props.get("ColumnCount", 0))
    borders = _as_bool(props.get("TableBordersVisible", True))

    if row_count == 0 or col_count == 0:
        return None

    return TableElement(
        rows=[],
        borders_visible=borders,
    )


def _extract_embedded_file(
    obj: ExtractedObject, file_data: dict[str, bytes]
) -> EmbeddedFile | None:
    """Extract embedded file from an EmbeddedFileNode."""
    props = obj.properties
    filename = _clean_text(str(props.get("EmbeddedFileName", "")))
    source_path = _clean_text(str(props.get("SourceFilepath", "")))

    data = b""
    container = props.get("EmbeddedFileContainer")
    if isinstance(container, bytes):
        data = container

    if not filename and not data:
        return None

    return EmbeddedFile(
        data=data,
        filename=filename,
        source_path=source_path,
    )


def _decode_text_value(value: object, encoding: str = "unicode") -> str:
    """Decode text from pyOneNote property values.

    pyOneNote returns text in various formats:
    - Direct string for RichEditTextUnicode
    - Hex string for TextExtendedAscii
    - Garbled UTF-16 decoded strings for TextExtendedAscii (needs re-encoding)
    - bytes for raw data
    """
    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned:
            return ""

        # Check if it's a hex string (common for TextExtendedAscii)
        if all(c in "0123456789abcdefABCDEF" for c in cleaned):
            try:
                raw = bytes.fromhex(cleaned)
                if encoding == "ascii":
                    return raw.decode("ascii", errors="replace").rstrip("\x00")
                else:
                    return raw.decode("utf-16-le", errors="replace").rstrip("\x00")
            except (ValueError, UnicodeDecodeError):
                pass

        # Check for garbled text: pyOneNote sometimes decodes ASCII bytes
        # as UTF-16LE, producing CJK/symbol characters. Detect and fix this.
        if encoding == "ascii" and _looks_garbled(cleaned):
            try:
                raw = cleaned.encode("utf-16-le")
                return raw.decode("ascii", errors="replace").rstrip("\x00")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass

        return _clean_text(cleaned)

    if isinstance(value, bytes):
        if encoding == "ascii":
            return value.decode("ascii", errors="replace").rstrip("\x00")
        try:
            return value.decode("utf-16-le").rstrip("\x00")
        except UnicodeDecodeError:
            return value.decode("latin-1").rstrip("\x00")

    return str(value) if value else ""


def _looks_garbled(text: str) -> bool:
    """Detect if text looks like ASCII bytes misinterpreted as UTF-16LE.

    This happens when pyOneNote decodes TextExtendedAscii raw bytes as
    UTF-16 strings instead of ASCII. The result contains CJK characters,
    unusual symbols, and zero-width spaces for normal English text.
    """
    if not text:
        return False
    # Count characters outside normal ASCII+extended range
    non_ascii = sum(1 for c in text if ord(c) > 0xFF)
    # If more than 30% of characters are non-ASCII, it's likely garbled
    return len(text) > 2 and non_ascii / len(text) > 0.3


def _section_name_from_path(file_path: str) -> str:
    """Extract a clean section name from a file path."""
    name = Path(file_path).stem
    name = re.sub(r"\s*\(On\s+\d+-\d+-\d+(?:\s*-\s*\d+)?\)", "", name)
    name = re.sub(r"\.one$", "", name, flags=re.IGNORECASE)
    return name.strip() or "Untitled"


def _clean_text(text: str) -> str:
    """Clean text by removing null bytes and extra whitespace."""
    return text.replace("\x00", "").strip()


def _as_bool(value: object) -> bool:
    """Convert a property value to bool."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() in ("true", "1", "yes")
    return bool(value)


def _parse_int_prop(value: object) -> int:
    """Parse an integer property."""
    if isinstance(value, int):
        return value
    if isinstance(value, bytes) and len(value) >= 2:
        return int.from_bytes(value[:4].ljust(4, b"\x00"), "little")
    if isinstance(value, str):
        match = re.search(r"\d+", value)
        if match:
            return int(match.group())
    return 0


def _parse_font_size(value: object) -> int:
    """Parse font size from pyOneNote format."""
    if isinstance(value, int):
        return value
    if isinstance(value, bytes):
        return int.from_bytes(value[:2].ljust(2, b"\x00"), "little")
    if isinstance(value, str):
        match = re.search(r"\d+", value)
        if match:
            return int(match.group())
    return 0


def _detect_image_format(data: bytes) -> str:
    """Detect image format from magic bytes."""
    if not data or len(data) < 4:
        return ""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:2] == b"BM":
        return "bmp"
    if data[:4] == b"RIFF" and len(data) > 8 and data[8:12] == b"WEBP":
        return "webp"
    return ""
