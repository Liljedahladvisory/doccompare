"""Detect non-body OOXML changes that Word/PDF diffs can hide."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree
from loguru import logger

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _qn(tag: str) -> str:
    prefix, local = tag.split(":")
    ns = {"w": W_NS, "xml": XML_NS}[prefix]
    return f"{{{ns}}}{local}"


W_T = _qn("w:t")
W_INSTR_TEXT = _qn("w:instrText")
_PART_RE = re.compile(r"^word/(header|footer)(\d+)\.xml$")
_PAGE_FIELDS = {"PAGE", "NUMPAGES", "SECTIONPAGES"}


def detect_header_footer_changes(original: Path, modified: Path) -> list[dict]:
    """Return human-readable changes in headers/footers, including page fields."""
    original_parts = _read_header_footer_parts(Path(original))
    modified_parts = _read_header_footer_parts(Path(modified))
    changes = []

    for key in sorted(set(original_parts) | set(modified_parts)):
        old = original_parts.get(key)
        new = modified_parts.get(key)
        if old and new and old["signature"] == new["signature"]:
            continue

        kind, index = key
        label = _part_label(kind, index)
        summary = _describe_change(label, old, new)
        if summary:
            changes.append({
                "kind": kind,
                "index": index,
                "part": label,
                "change_type": _change_type(old, new),
                "summary": summary,
            })

    return changes


def _read_header_footer_parts(docx_path: Path) -> dict[tuple[str, int], dict]:
    parts = {}
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            for name in zf.namelist():
                match = _PART_RE.match(name)
                if not match:
                    continue
                kind, index = match.group(1), int(match.group(2))
                parts[(kind, index)] = _extract_part_signature(
                    zf.read(name), kind, index)
    except Exception as e:
        logger.warning(f"Could not inspect headers/footers in {docx_path.name}: {e}")
    return parts


def _extract_part_signature(xml_bytes: bytes, kind: str, index: int) -> dict:
    tree = etree.fromstring(xml_bytes)
    text = _normalize_text(" ".join(
        t.text for t in tree.iter(W_T) if t.text
    ))
    fields = tuple(sorted(set(
        _normalize_field(f.text) for f in tree.iter(W_INSTR_TEXT) if f.text
    )))
    return {
        "kind": kind,
        "index": index,
        "text": text,
        "fields": fields,
        "signature": (text, fields),
    }


def _normalize_text(text: str) -> str:
    return " ".join(text.split())


def _normalize_field(field: str) -> str:
    return re.sub(r"\s+", " ", field.strip().upper())


def _part_label(kind: str, index: int) -> str:
    name = "Header" if kind == "header" else "Footer"
    return f"{name} {index}"


def _describe_change(label: str, old: dict | None, new: dict | None) -> str:
    if old is None and new is None:
        return ""
    if old is None:
        details = _describe_added_content(new)
        return f"{label} added{': ' + details if details else ''}"
    if new is None:
        return f"{label} removed"

    old_fields = set(old["fields"])
    new_fields = set(new["fields"])
    added_fields = sorted(new_fields - old_fields)
    removed_fields = sorted(old_fields - new_fields)
    details = []

    added_page_fields = [f for f in added_fields if _is_page_field(f)]
    removed_page_fields = [f for f in removed_fields if _is_page_field(f)]
    if added_page_fields:
        details.append(_field_phrase(added_page_fields, "added"))
    if removed_page_fields:
        details.append(_field_phrase(removed_page_fields, "removed"))

    other_added = [f for f in added_fields if f not in added_page_fields]
    other_removed = [f for f in removed_fields if f not in removed_page_fields]
    if other_added:
        details.append("field added: " + ", ".join(other_added))
    if other_removed:
        details.append("field removed: " + ", ".join(other_removed))

    if old["text"] != new["text"]:
        details.append(_text_change(old["text"], new["text"]))

    return f"{label} changed: {'; '.join(details) if details else 'content changed'}"


def _change_type(old: dict | None, new: dict | None) -> str:
    if old is None:
        return "added"
    if new is None:
        return "deleted"
    return "modified"


def _describe_added_content(part: dict | None) -> str:
    if not part:
        return ""
    details = []
    page_fields = [f for f in part["fields"] if _is_page_field(f)]
    if page_fields:
        details.append(_field_phrase(page_fields, "added"))
    other_fields = [f for f in part["fields"] if f not in page_fields]
    if other_fields:
        details.append("field added: " + ", ".join(other_fields))
    if part["text"]:
        details.append(f'text "{_shorten(part["text"])}"')
    return "; ".join(details)


def _field_phrase(fields: list[str], action: str) -> str:
    names = [_friendly_field_name(f) for f in fields]
    return ", ".join(names) + f" {action}"


def _friendly_field_name(field: str) -> str:
    token = field.split()[0] if field else field
    return {
        "PAGE": "page number field",
        "NUMPAGES": "total pages field",
        "SECTIONPAGES": "section pages field",
    }.get(token, field)


def _is_page_field(field: str) -> bool:
    token = field.split()[0] if field else field
    return token in _PAGE_FIELDS


def _text_change(old_text: str, new_text: str) -> str:
    if old_text and new_text:
        return f'text changed from "{_shorten(old_text)}" to "{_shorten(new_text)}"'
    if new_text:
        return f'text added "{_shorten(new_text)}"'
    return f'text removed "{_shorten(old_text)}"'


def _shorten(text: str, limit: int = 90) -> str:
    if len(text) <= limit:
        return text
    return text[:limit - 1].rstrip() + "..."
