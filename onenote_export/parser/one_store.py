"""High-level OneNote parser using pyOneNote as the binary parsing engine.

Extracts structured content (text, images, formatting) from .one files
and organizes it by page.
"""

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

from pyOneNote.Header import Header
from pyOneNote.OneDocument import OneDocment

logger = logging.getLogger(__name__)

# JCID type names from the OneNote spec
_PAGE_META = "jcidPageMetaData"
_SECTION_NODE = "jcidSectionNode"
_SECTION_META = "jcidSectionMetaData"
_PAGE_SERIES = "jcidPageSeriesNode"
_PAGE_MANIFEST = "jcidPageManifestNode"
_PAGE_NODE = "jcidPageNode"
_TITLE_NODE = "jcidTitleNode"
_OUTLINE_NODE = "jcidOutlineNode"
_OUTLINE_ELEMENT = "jcidOutlineElementNode"
_RICH_TEXT = "jcidRichTextOENode"
_IMAGE_NODE = "jcidImageNode"
_TABLE_NODE = "jcidTableNode"
_TABLE_ROW = "jcidTableRowNode"
_TABLE_CELL = "jcidTableCellNode"
_EMBEDDED_FILE = "jcidEmbeddedFileNode"
_NUMBER_LIST = "jcidNumberListNode"
_STYLE_CONTAINER = "jcidPersistablePropertyContainerForTOCSection"
_REVISION_META = "jcidRevisionMetaData"


@dataclass
class ExtractedProperty:
    """A single property from a OneNote object."""
    name: str
    value: object  # str, bytes, int, bool, list, etc.


@dataclass
class ExtractedObject:
    """A parsed object from the OneNote file."""
    obj_type: str
    identity: str
    properties: dict[str, object] = field(default_factory=dict)


@dataclass
class ExtractedPage:
    """A page with its title and content objects."""
    title: str = ""
    level: int = 0
    author: str = ""
    creation_time: str = ""
    last_modified: str = ""
    objects: list[ExtractedObject] = field(default_factory=list)


@dataclass
class ExtractedSection:
    """All pages extracted from a single .one file."""
    file_path: str = ""
    display_name: str = ""
    pages: list[ExtractedPage] = field(default_factory=list)
    file_data: dict[str, bytes] = field(default_factory=dict)


class OneStoreParser:
    """Parses a MS-ONESTORE (.one) file using pyOneNote."""

    def __init__(self, file_path: str | Path) -> None:
        self.file_path = Path(file_path)

    def parse(self) -> ExtractedSection:
        """Parse the .one file and return structured content."""
        section = ExtractedSection(file_path=str(self.file_path))

        with open(self.file_path, "rb") as f:
            doc = OneDocment(f)

        # Validate it's a .one file
        if doc.header.guidFileType != Header.ONE_UUID:
            raise ValueError(f"{self.file_path} is not a .one file")

        # Get all properties (objects with their property sets)
        raw_props = doc.get_properties()

        # Get embedded files
        raw_files = doc.get_files()
        for guid, finfo in raw_files.items():
            content = finfo.get("content", b"")
            if content:
                section.file_data[guid] = content

        # Convert raw properties to ExtractedObjects
        all_objects = []
        for raw in raw_props:
            obj = ExtractedObject(
                obj_type=raw["type"],
                identity=raw["identity"],
                properties=dict(raw["val"]),
            )
            all_objects.append(obj)

        # Build page structure
        section.pages = self._build_pages(all_objects)
        section.display_name = self._extract_section_name(all_objects)

        return section

    def _extract_section_name(self, objects: list[ExtractedObject]) -> str:
        """Extract section display name from section metadata."""
        for obj in objects:
            if obj.obj_type == _SECTION_META:
                name = obj.properties.get("SectionDisplayName", "")
                if name:
                    return str(name).strip()
        return ""

    def _build_pages(
        self, objects: list[ExtractedObject]
    ) -> list[ExtractedPage]:
        """Group objects into pages based on the document structure.

        OneNote stores multiple revisions per page, each sharing a GUID.
        Content objects (text, images, etc.) share the GUID of their
        owning page node.  Orphan page metadata entries (from older
        revisions) have a different GUID with no associated content.

        Strategy:
        1. Identify "content GUIDs" — GUIDs that own a jcidPageNode
           (these are the actual page revisions with content objects).
        2. Build one page per content GUID using its metadata + content.
        3. If a content GUID has no metadata, fall back to matching
           orphan metadata by title.
        """
        pages: list[ExtractedPage] = []

        # Classify every object by type, indexed by GUID
        guid_objects: dict[str, list[ExtractedObject]] = {}
        page_metas: list[ExtractedObject] = []
        page_node_guids: list[str] = []

        for obj in objects:
            guid = _extract_guid(obj.identity)
            guid_objects.setdefault(guid, []).append(obj)

            if obj.obj_type == _PAGE_META:
                page_metas.append(obj)
            elif obj.obj_type == _PAGE_NODE:
                if guid not in page_node_guids:
                    page_node_guids.append(guid)

        # No page metadata at all — single unnamed page
        if not page_metas:
            all_content = [
                o for o in objects
                if o.obj_type in (
                    _RICH_TEXT, _IMAGE_NODE, _TABLE_NODE,
                    _TABLE_ROW, _TABLE_CELL, _EMBEDDED_FILE,
                    _OUTLINE_ELEMENT, _OUTLINE_NODE, _NUMBER_LIST,
                )
            ]
            if all_content:
                pages.append(ExtractedPage(objects=all_content))
            return pages

        # Build a lookup: GUID -> page metadata
        meta_by_guid: dict[str, ExtractedObject] = {}
        for meta in page_metas:
            guid = _extract_guid(meta.identity)
            # Later entries (newer revisions) overwrite earlier ones
            meta_by_guid[guid] = meta

        # Orphan metas: metadata GUIDs with no page node (old revisions)
        orphan_metas: dict[str, ExtractedObject] = {
            g: m for g, m in meta_by_guid.items()
            if g not in page_node_guids
        }

        # Content types we care about
        _CONTENT_TYPES = {
            _RICH_TEXT, _IMAGE_NODE, _TABLE_NODE,
            _TABLE_ROW, _TABLE_CELL, _EMBEDDED_FILE,
            _OUTLINE_ELEMENT, _OUTLINE_NODE, _NUMBER_LIST,
        }

        # Build one page per content GUID (GUID that has a PageNode)
        seen_titles: dict[str, int] = {}
        for content_guid in page_node_guids:
            objs = guid_objects.get(content_guid, [])

            # Find metadata — prefer same GUID, fall back to orphan
            meta = meta_by_guid.get(content_guid)
            if not meta:
                # Match orphan metadata by scanning (first available)
                for og, om in list(orphan_metas.items()):
                    meta = om
                    del orphan_metas[og]
                    break

            title = ""
            level = 0
            creation = ""
            if meta:
                title = _clean_text(
                    str(meta.properties.get("CachedTitleString", ""))
                )
                level = _parse_int(meta.properties.get("PageLevel", 0))
                creation = str(
                    meta.properties.get("TopologyCreationTimeStamp", "")
                )

            # Extract author from the page node
            author = ""
            last_modified = ""
            for o in objs:
                if o.obj_type == _PAGE_NODE:
                    author = _clean_text(
                        str(o.properties.get("Author", ""))
                    )
                    last_modified = str(
                        o.properties.get("LastModifiedTime", "")
                    )
                    break

            # Collect content objects for this GUID only
            content = [o for o in objs if o.obj_type in _CONTENT_TYPES]

            page = ExtractedPage(
                title=title or "Untitled",
                level=level,
                author=author,
                creation_time=creation,
                last_modified=last_modified,
                objects=content,
            )

            # Deduplicate by title — keep the version with more content
            key = title.lower().strip()
            if key in seen_titles:
                idx = seen_titles[key]
                if len(content) > len(pages[idx].objects):
                    pages[idx] = page
            else:
                seen_titles[key] = len(pages)
                pages.append(page)

        return pages


def _extract_guid(identity_str: str) -> str:
    """Extract the GUID from an ExtendedGUID identity string.

    Input format: '<ExtendedGUID> (guid-string, n)'
    Returns just the guid-string part.
    """
    match = re.search(r"\(([^,]+),", identity_str)
    if match:
        return match.group(1).strip()
    return ""


def _clean_text(text: str) -> str:
    """Clean up text by stripping null bytes and extra whitespace."""
    text = text.replace("\x00", "").strip()
    return text


def _parse_int(value: object) -> int:
    """Parse an integer from various formats pyOneNote returns."""
    if isinstance(value, int):
        return value
    if isinstance(value, bytes):
        try:
            return int.from_bytes(value[:4], "little")
        except Exception:
            return 0
    if isinstance(value, str):
        # Try to extract numeric value
        match = re.search(r"\d+", value)
        if match:
            return int(match.group())
    return 0
