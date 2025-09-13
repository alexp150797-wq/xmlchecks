# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List, Dict

from openpyxl import Workbook
from .xlsx_utils import autosize, apply_borders, style_sheet, add_summary_sheet

BASE_HEADERS = [
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
    "Статус",
    "Подробности",
    "Рекомендации",
]

PDF_NAME_COL = "Имя PDF соответствует шаблону"


def get_headers(include_pdf_name_col: bool) -> List[str]:
    headers = BASE_HEADERS.copy()
    if include_pdf_name_col:
        headers.insert(13, PDF_NAME_COL)
    return headers


def write_xlsx_iul(
    rows: List[Dict],
    out_path: Path,
    include_pdf_name_col: bool = True,
) -> tuple[int, dict]:
    wb = Workbook()
    ws = wb.active
    ws.title = "ИУЛ - IFC"

    headers = get_headers(include_pdf_name_col)
    ws.append(headers)
    for r in rows:
        row_data = [
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
        ]
        if include_pdf_name_col:
            row_data.append(r.get(PDF_NAME_COL))
        row_data.extend([
            r.get("Статус"),
            r.get("Подробности"),
            r.get("recommendation"),
        ])
        ws.append(row_data)

    if include_pdf_name_col:
        yes_no_cols = ("J", "K", "L", "M", "N")
        status_col = "O"
        left_cols = (1, 2, 3, 6, 7, 16, 17)
    else:
        yes_no_cols = ("J", "K", "L", "M")
        status_col = "N"
        left_cols = (1, 2, 3, 6, 7, 15, 16)

    style_sheet(
        ws,
        left_cols=left_cols,
        yes_no_cols=yes_no_cols,
        status_col=status_col,
    )
    autosize(ws)
    apply_borders(ws)

    stats = add_summary_sheet(wb, rows, title="Итого ИУЛ")

    wb.save(out_path)
    exit_code = 0 if stats["errors"] == 0 else 1
    return exit_code, stats
