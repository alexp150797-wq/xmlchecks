# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, Any, Optional
import xml.etree.ElementTree as ET

try:
    import yaml  # type: ignore
except Exception:
    yaml = None  # type: ignore

DEFAULT_RULES = {
    "entry_tag": "ModelFile",
    "name_tag": "FileName",
    "checksum_tag": "FileChecksum",
    "format_tag": "FileFormat",
    "filter_format": "IFC",
}

def read_rules(path: Path) -> Dict[str, Any]:
    rules = DEFAULT_RULES.copy()
    try:
        if path and path.exists() and yaml is not None:
            with path.open("r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
                for k in DEFAULT_RULES:
                    if k in data and data[k] is not None:
                        rules[k] = data[k]
    except Exception:
        pass
    return rules

def _localname(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag

def _find_child_text(elem: ET.Element, wanted: str) -> Optional[str]:
    wanted = wanted.lower()
    for ch in list(elem):
        if _localname(ch.tag).lower() == wanted:
            txt = (ch.text or "").strip()
            return txt if txt else None
    return None

def extract_from_xml(xml_path: Path, rules: Dict[str, Any], case_sensitive: bool=True) -> Dict[str, Dict[str, Any]]:
    if not xml_path.exists():
        raise FileNotFoundError(f"XML не найден: {xml_path}")
    tree = ET.parse(str(xml_path))
    root = tree.getroot()

    entry_tag = (rules.get("entry_tag") or DEFAULT_RULES["entry_tag"])
    name_tag = (rules.get("name_tag") or DEFAULT_RULES["name_tag"])
    checksum_tag = (rules.get("checksum_tag") or DEFAULT_RULES["checksum_tag"])
    format_tag = (rules.get("format_tag") or DEFAULT_RULES["format_tag"])
    filter_format = rules.get("filter_format", DEFAULT_RULES["filter_format"])

    entries = [e for e in root.iter() if _localname(getattr(e, "tag", "")).lower() == str(entry_tag).lower()]
    result: Dict[str, Dict[str, Any]] = {}

    for e in entries:
        name = _find_child_text(e, name_tag)
        if not name:
            continue
        fmt = _find_child_text(e, format_tag)
        crc = _find_child_text(e, checksum_tag)
        if crc:
            crc = crc.strip().upper()

        if filter_format and (fmt or "").strip().upper() != str(filter_format).upper():
            continue

        key = name if case_sensitive else name.lower()
        if key not in result:
            result[key] = {"crc_hex": crc, "format": fmt}

    return result
