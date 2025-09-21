# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List, Dict

from openpyxl import Workbook
from .xlsx_utils import autosize, apply_borders, style_sheet, add_summary_sheet

HEADERS = [
    "Имя файла IFC",
    "Имя файла IFC из XML",
    "CRC-32 XML",
    "CRC-32 PDF",
    "Имя совпадает",
    "CRC совпадает",
    "Статус",
    "Подробности",
]

def write_xlsx_pdf_xml(rows: List[Dict], out_path: Path) -> tuple[int, dict]:
    wb = Workbook()
    ws = wb.active; ws.title = "PDF-XML Report"

    ws.append(HEADERS)
    for r in rows:
        ws.append([
            r.get("Имя файла IFC"),
            r.get("Имя файла IFC из XML"),
            r.get("CRC-32 XML"),
            r.get("CRC-32 PDF"),
            r.get("Имя совпадает"),
            r.get("CRC совпадает"),
            r.get("Статус"),
            r.get("Подробности"),
        ])

    style_sheet(ws, left_cols=(1,2,7,8), yes_no_cols=("E","F"), status_col="G")
    autosize(ws)
    apply_borders(ws)

    stats = add_summary_sheet(wb, rows, title="Summary PDF-XML")

    wb.save(out_path)
    exit_code = 0 if stats["errors"] == 0 else 1
    return exit_code, stats
