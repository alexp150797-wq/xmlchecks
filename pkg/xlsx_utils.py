# -*- coding: utf-8 -*-
"""Utility helpers shared by XLSX report writers."""
from __future__ import annotations
from typing import Iterable, List, Dict
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

GREEN = PatternFill(start_color="E7F7E7", end_color="E7F7E7", fill_type="solid")
RED = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
GRAY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

def autosize(ws) -> None:
    """Auto fit column width based on cell contents."""
    max_width: Dict[int, int] = {}
    for row in ws.iter_rows(values_only=True):
        for idx, val in enumerate(row, start=1):
            s = "" if val is None else str(val)
            width = max(3, min(120, int(len(s) * 1.1) + 2))
            if idx not in max_width or width > max_width[idx]:
                max_width[idx] = width
    for idx, w in max_width.items():
        ws.column_dimensions[get_column_letter(idx)].width = w

def apply_borders(ws) -> None:
    """Apply thin grey border to all cells."""
    thin = Side(border_style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border

def style_sheet(ws, left_cols: Iterable[int], yes_no_cols: Iterable[str], status_col: str) -> None:
    """Apply common styling and conditional formatting.

    Parameters
    ----------
    left_cols: Iterable[int]
        Columns that should be left-aligned.
    yes_no_cols: Iterable[str]
        Columns containing "Да"/"Нет" values.
    status_col: str
        Column with overall status ("OK" or other).
    """
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = GRAY

    left_cols_set = set(left_cols)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in left_cols_set:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    for col in yes_no_cols:
        ws.conditional_formatting.add(f"{col}2:{col}{ws.max_row}", CellIsRule(operator="equal", formula=['"Да"'], fill=GREEN))
        ws.conditional_formatting.add(f"{col}2:{col}{ws.max_row}", CellIsRule(operator="equal", formula=['"Нет"'], fill=RED))

    ws.conditional_formatting.add(f"{status_col}2:{status_col}{ws.max_row}", CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN))
    ws.conditional_formatting.add(f"{status_col}2:{status_col}{ws.max_row}", CellIsRule(operator="notEqual", formula=['"OK"'], fill=RED))

def add_summary_sheet(wb, rows: List[Dict], status_key: str = "Статус") -> Dict[str, int]:
    """Create a Summary sheet with statistics.

    Returns dict with keys ``total``, ``ok`` and ``errors``.
    """
    sm = wb.create_sheet("Summary")
    total = len(rows)
    ok = sum(1 for r in rows if r.get(status_key) == "OK")
    errors = total - ok
    sm.append(["Метрика", "Значение"])
    sm.append(["Всего строк", total])
    sm.append(["OK", ok])
    sm.append(["Ошибки (все не-OK)", errors])
    sm.auto_filter.ref = sm.dimensions
    sm.freeze_panes = "A2"
    for c in sm[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = GRAY
    autosize(sm)
    apply_borders(sm)
    return {"total": total, "ok": ok, "errors": errors}
