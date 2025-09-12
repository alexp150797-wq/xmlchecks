# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
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
    "filter_format": ["IFC", "PDF"],
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

def extract_from_xml(xml_path: Path, rules: Dict[str, Any], case_sensitive: bool=True) -> Tuple[Dict[str, Dict[str, Any]], List[Dict[str, Any]]]:
    if not xml_path.exists():
        raise FileNotFoundError(f"XML не найден: {xml_path}")
    tree = ET.parse(str(xml_path))
    root = tree.getroot()

    entry_tag = (rules.get("entry_tag") or DEFAULT_RULES["entry_tag"])
    name_tag = (rules.get("name_tag") or DEFAULT_RULES["name_tag"])
    checksum_tag = (rules.get("checksum_tag") or DEFAULT_RULES["checksum_tag"])
    format_tag = (rules.get("format_tag") or DEFAULT_RULES["format_tag"])
    filter_format = rules.get("filter_format", DEFAULT_RULES["filter_format"])
    if isinstance(filter_format, str):
        filter_set = {filter_format.strip().upper()} if filter_format else set()
    else:
        try:
            filter_set = {str(x).strip().upper() for x in filter_format if x}
        except Exception:
            filter_set = set()

    entries = [e for e in root.iter() if _localname(getattr(e, "tag", "")).lower() == str(entry_tag).lower()]
    result_ifc: Dict[str, Dict[str, Any]] = {}
    result_pdf: List[Dict[str, Any]] = []

    for e in entries:
        name = _find_child_text(e, name_tag)
        if not name:
            continue
        fmt = _find_child_text(e, format_tag)
        crc = _find_child_text(e, checksum_tag)
        if crc:
            crc = crc.strip().upper()
        fmt_upper = (fmt or "").strip().upper()

        if (not filter_set) or (fmt_upper in filter_set):
            key = name if case_sensitive else name.lower()
            if key not in result_ifc:
                result_ifc[key] = {"crc_hex": crc, "format": fmt}

        for ch in list(e):
            if _localname(ch.tag).lower() != "signfile":
                continue
            s_name = _find_child_text(ch, name_tag)
            if not s_name:
                continue
            s_fmt = _find_child_text(ch, format_tag)
            s_crc = _find_child_text(ch, checksum_tag)
            if s_crc:
                s_crc = s_crc.strip().upper()
            s_fmt_upper = (s_fmt or "").strip().upper()
            if (not filter_set) or (s_fmt_upper in filter_set):
                result_pdf.append({
                    "name": s_name if case_sensitive else s_name.lower(),
                    "format": s_fmt,
                    "crc_hex": s_crc,
                })

    return result_ifc, result_pdf
