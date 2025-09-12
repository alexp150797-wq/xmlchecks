# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Optional

from .crc import compute_crc32

def build_report_pdf_pairs(pairs: Dict[str, Dict[str, Optional[str]]], pdf_files: List[Path]) -> List[Dict]:
    rows: List[Dict] = []
    index = {p.name: p for p in pdf_files}

    for iul_name, meta in pairs.items():
        base_name = meta.get("base_name")
        iul_crc = (meta.get("iul_crc") or "").upper() or None
        base_crc = (meta.get("base_crc") or "").upper() or None

        iul_path = index.get(iul_name)
        base_path = index.get(base_name) if base_name else None

        status: List[str] = []
        details: List[str] = []

        if not iul_path:
            status.append("MISSING_IUL")
            details.append("Не найден PDF ИУЛ")
        if not base_path:
            status.append("MISSING_BASE")
            details.append("Не найден PDF осн.")

        if iul_path and iul_crc:
            actual_iul = f"{compute_crc32(iul_path):08X}"
            if actual_iul != iul_crc:
                status.append("CRC_MISMATCH")
                details.append(f"CRC ИУЛ: XML={iul_crc}, факт={actual_iul}")
        elif iul_path and not iul_crc:
            actual_iul = f"{compute_crc32(iul_path):08X}"
            status.append("CRC_MISMATCH")
            details.append(f"CRC ИУЛ отсутствует в XML, факт={actual_iul}")

        if base_path and base_crc:
            actual_base = f"{compute_crc32(base_path):08X}"
            if actual_base != base_crc:
                if "CRC_MISMATCH" not in status:
                    status.append("CRC_MISMATCH")
                details.append(f"CRC осн.: XML={base_crc}, факт={actual_base}")
        elif base_path and not base_crc:
            actual_base = f"{compute_crc32(base_path):08X}"
            if "CRC_MISMATCH" not in status:
                status.append("CRC_MISMATCH")
            details.append(f"CRC осн. отсутствует в XML, факт={actual_base}")

        if not status:
            status.append("OK")

        rows.append({
            "PDF ИУЛ": iul_name,
            "CRC ИУЛ": iul_crc,
            "PDF осн.": base_name,
            "CRC осн.": base_crc,
            "Статус": ";".join(status),
            "Подробности": "; ".join(details) if details else None,
        })

    return rows

