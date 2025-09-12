# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List, Dict

from openpyxl import Workbook

from .xlsx_utils import _autosize, _borders, _style, create_summary_sheet

HEADERS = [
    "Имя файла",
    "Имя PDF",
    "Файл из ИУЛ",
    "CRC-32 ИУЛ",
    "CRC-32 IFC",
    "Дата/время ИУЛ",
    "Дата/время IFC",
    "Размер ИУЛ, байт",
    "Размер IFC, байт",
    "Имя совпадает",
    "CRC совпадает",
    "Дата/время совпадает",
    "Размер совпадает",
    "Имя PDF соответствует правилу",
    "Статус",
    "Подробности",
]

def write_xlsx_iul(rows: List[Dict], out_path: Path) -> tuple[int, dict]:
    wb = Workbook()
    ws = wb.active; ws.title = "IUL Report"

    ws.append(HEADERS)
    for r in rows:
        ws.append([
            r.get("Имя файла"),
            r.get("Имя PDF"),
            r.get("Файл из ИУЛ"),
            r.get("CRC-32 ИУЛ"),
            r.get("CRC-32 IFC"),
            r.get("Дата/время ИУЛ"),
            r.get("Дата/время IFC"),
            r.get("Размер ИУЛ, байт"),
            r.get("Размер IFC, байт"),
            r.get("Имя совпадает"),
            r.get("CRC совпадает"),
            r.get("Дата/время совпадает"),
            r.get("Размер совпадает"),
            r.get("Имя PDF соответствует правилу"),
            r.get("Статус"),
            r.get("Подробности"),
        ])

    _style(
        ws,
        left_cols=(1, 2, 3, 6, 7, 16),
        yes_no_cols=("J", "K", "L", "M", "N"),
        status_col="O",
    )
    _autosize(ws)
    _borders(ws)

    stats = create_summary_sheet(wb, rows)

    wb.save(out_path)
    exit_code = 0 if stats["errors"] == 0 else 1
    return exit_code, stats
