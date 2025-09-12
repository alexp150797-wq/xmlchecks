# -*- coding: utf-8 -*-
"""XLSX writer producing a single workbook with multiple report sheets."""
from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Optional

from openpyxl import Workbook

from .xlsx_writer import HEADERS as XML_HEADERS
from .xlsx_writer_iul import HEADERS as IUL_HEADERS
from .xlsx_writer_pdf_xml import HEADERS as PDF_XML_HEADERS
from .xlsx_utils import autosize, apply_borders, style_sheet, add_summary_sheet


def _add_sheet(
    wb: Workbook,
    title: str,
    headers: List[str],
    rows: List[Dict],
    left_cols: tuple[int, ...],
    yes_no_cols: tuple[str, ...],
    status_col: str,
    summary_title: str,
) -> Dict[str, int]:
    ws = wb.create_sheet(title)
    ws.append(headers)
    for r in rows:
        row_data = []
        for h in headers:
            if h == "Рекомендации":
                row_data.append(r.get("recommendation"))
            else:
                row_data.append(r.get(h))
        ws.append(row_data)
    style_sheet(ws, left_cols=left_cols, yes_no_cols=yes_no_cols, status_col=status_col)
    autosize(ws)
    apply_borders(ws)
    return add_summary_sheet(wb, rows, title=summary_title)


def write_combined_xlsx(
    xml_rows: Optional[List[Dict]],
    iul_rows: Optional[List[Dict]],
    pdf_xml_rows: Optional[List[Dict]],
    out_path: Path,
) -> Dict[str, Dict[str, int]]:
    """Write selected reports to a single XLSX workbook.

    Returns a mapping of report keys to statistics dictionaries.
    """
    wb = Workbook()
    wb.remove(wb.active)
    stats: Dict[str, Dict[str, int]] = {}

    if xml_rows:
        stats["xml"] = _add_sheet(
            wb,
            "XML - IFC",
            XML_HEADERS,
            xml_rows,
            left_cols=(1,2,7,8,9),
            yes_no_cols=("E","F"),
            status_col="G",
            summary_title="Итого XML",
        )

    if iul_rows:
        stats["iul"] = _add_sheet(
            wb,
            "ИУЛ - IFC",
            IUL_HEADERS,
            iul_rows,
            left_cols=(1,2,3,6,7,16,17),
            yes_no_cols=("J","K","L","M","N"),
            status_col="O",
            summary_title="Итого ИУЛ",
        )

    if pdf_xml_rows:
        stats["pdf_xml"] = _add_sheet(
            wb,
            "PDF-XML Report",
            PDF_XML_HEADERS,
            pdf_xml_rows,
            left_cols=(1,2,7,8),
            yes_no_cols=("E","F"),
            status_col="G",
            summary_title="Summary PDF-XML",
        )

    wb.save(out_path)
    return stats
