# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List, Dict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

GREEN = PatternFill(start_color="E7F7E7", end_color="E7F7E7", fill_type="solid")
RED = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
GRAY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

HEADERS = [
    "Имя файла",
    "Файл из XML",
    "CRC-32 XML",
    "CRC-32 IFC",
    "Имя совпадает",
    "CRC совпадает",
    "Статус",
    "Подробности",
    "Рекомендация",
]

def _autosize(ws):
    max_width = {}
    for row in ws.iter_rows(values_only=True):
        for idx, val in enumerate(row, start=1):
            s = "" if val is None else str(val)
            width = max(3, min(120, int(len(s) * 1.1) + 2))
            if idx not in max_width or width > max_width[idx]:
                max_width[idx] = width
    for idx, w in max_width.items():
        ws.column_dimensions[get_column_letter(idx)].width = w

def _borders(ws):
    thin = Side(border_style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border

def _style(ws):
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = GRAY

    left_cols = (1, 2, 7, 8, 9)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in left_cols:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    for col in ("E","F"):
        ws.conditional_formatting.add(f"{col}2:{col}{ws.max_row}", CellIsRule(operator='equal', formula=['"Да"'], fill=GREEN))
        ws.conditional_formatting.add(f"{col}2:{col}{ws.max_row}", CellIsRule(operator='equal', formula=['"Нет"'], fill=RED))

    ws.conditional_formatting.add(f"G2:G{ws.max_row}", CellIsRule(operator='equal', formula=['"OK"'], fill=GREEN))
    ws.conditional_formatting.add(f"G2:G{ws.max_row}", CellIsRule(operator='notEqual', formula=['"OK"'], fill=RED))

def write_xlsx(rows: List[Dict], out_path: Path) -> tuple[int, dict]:
    wb = Workbook()
    ws = wb.active
    ws.title = "XML Report"

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
            r.get("Рекомендация"),
        ])

    _style(ws)
    _autosize(ws)
    _borders(ws)

    sm = wb.create_sheet("Summary")
    total = len(rows)
    ok = sum(1 for r in rows if r["Статус"] == "OK")
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
    # автоширина и рамки у Summary
    maxw = {}
    for row in sm.iter_rows(values_only=True):
        for idx, val in enumerate(row, start=1):
            s = "" if val is None else str(val)
            w = max(3, min(80, int(len(s)*1.1)+2))
            if idx not in maxw or w>maxw[idx]:
                maxw[idx]=w
    from openpyxl.utils import get_column_letter
    for idx, w in maxw.items():
        sm.column_dimensions[get_column_letter(idx)].width = w
    thin = Side(border_style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in sm.iter_rows(min_row=1, max_row=sm.max_row, min_col=1, max_col=sm.max_column):
        for cell in row:
            cell.border = border

    wb.save(out_path)
    exit_code = 0 if errors==0 else 1
    return exit_code, {"total": total, "ok": ok, "errors": errors}
