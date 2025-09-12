# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List, Dict

from openpyxl import Workbook

from .xlsx_utils import _autosize, _borders, _style, create_summary_sheet

HEADERS = [
    "Имя файла",
    "Файл из XML",
    "CRC-32 XML",
    "CRC-32 IFC",
    "Имя совпадает",
    "CRC совпадает",
    "Статус",
    "Подробности",
]

def write_xlsx(rows: List[Dict], out_path: Path) -> tuple[int, dict]:
    wb = Workbook()
    ws = wb.active; ws.title = "XML Report"

    ws.append(HEADERS)
    for r in rows:
        ws.append([
            r.get("Имя файла"),
            r.get("Файл из XML"),
            r.get("CRC-32 XML"),
            r.get("CRC-32 IFC"),
            r.get("Имя совпадает"),
            r.get("CRC совпадает"),
            r.get("Статус"),
            r.get("Подробности"),
        ])

    _style(ws, left_cols=(1, 2, 7, 8), yes_no_cols=("E", "F"), status_col="G")
    _autosize(ws)
    _borders(ws)

    stats = create_summary_sheet(wb, rows)

    wb.save(out_path)
    exit_code = 0 if stats["errors"] == 0 else 1
    return exit_code, stats
