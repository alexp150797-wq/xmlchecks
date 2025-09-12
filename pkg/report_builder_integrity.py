# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, List

from .iul_reader import IulEntry


def build_integrity_report(
    xml_map: Dict[str, dict],
    iul_map: Dict[str, IulEntry],
    ifc_files: List[Path],
    *,
    case_sensitive: bool = True,
) -> List[Dict]:
    """Проверка наличия записей во всех источниках (IFC, PDF, XML)."""
    names_ifc = {
        f.name if case_sensitive else f.name.lower() for f in ifc_files
    }
    names_pdf = {
        k if case_sensitive else k.lower() for k in iul_map.keys()
    }
    names_xml = {
        k if case_sensitive else k.lower() for k in xml_map.keys()
    }

    all_names = sorted(names_ifc | names_pdf | names_xml)
    rows: List[Dict] = []
    for name in all_names:
        has_ifc = name in names_ifc
        has_pdf = name in names_pdf
        has_xml = name in names_xml

        status: List[str] = []
        details: List[str] = []
        if not has_ifc:
            status.append("MISSING_IFC")
            details.append("IFC отсутствует")
        if not has_pdf:
            status.append("MISSING_PDF")
            details.append("PDF отсутствует")
        if not has_xml:
            status.append("MISSING_XML")
            details.append("XML отсутствует")
        if not status:
            status.append("OK")

        rows.append(
            {
                "IFC": name if has_ifc else None,
                "PDF": name if has_pdf else None,
                "XML": name if has_xml else None,
                "has_ifc": has_ifc,
                "has_pdf": has_pdf,
                "has_xml": has_xml,
                "Статус": ";".join(status),
                "Подробности": "; ".join(details) if details else None,
            }
        )

    return rows
