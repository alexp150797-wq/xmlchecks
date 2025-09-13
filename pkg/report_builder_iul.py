# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, List
import time
from .crc import compute_crc32
from .iul_reader import IulEntry, pdf_name_ok_lenient, pdf_name_ok_strict
from .utils import tri, recommendation


RECOMMENDATIONS = {
    "OK": "Действий не требуется",
    "ERROR_IFC_EXTRA": "Удалите лишний файл или добавьте запись в ИУЛ",
    "ERROR_IUL_EXTRA": "Удалите лишнюю запись из ИУЛ или добавьте соответствующий файл",
    "CRC_MISMATCH": "Проверьте корректность файлов и пересоздайте CRC",
    "NAME_MISMATCH": "Переименуйте файл или обновите запись в ИУЛ",
    "SIZE_MISMATCH": "Проверьте размер файла и обновите информацию в ИУЛ",
    "DT_MISMATCH": "Обновите дату/время в ИУЛ или замените файл",
    "PDF_NAME_MISMATCH": "Переименуйте PDF согласно требуемому правилу (ожидаемое имя: {expected})",
}

def _fmt_mtime(ts: float) -> str:
    t = time.localtime(ts)
    return f"{t.tm_mday:02d}.{t.tm_mon:02d}.{t.tm_year:04d} {t.tm_hour:02d}:{t.tm_min:02d}"

PDF_NAME_COL = "Имя PDF соответствует шаблону"


def build_report_iul(
    iul_map: Dict[str, IulEntry],
    ifc_files: List[Path],
    strict_pdf_name: bool = False,
    include_pdf_name_col: bool | None = None,
) -> List[Dict]:
    if include_pdf_name_col is None:
        include_pdf_name_col = strict_pdf_name
    rows: List[Dict] = []
    used = set()

    iul_crc_index: Dict[str, List[str]] = {}
    for k, e in iul_map.items():
        if e.crc_hex:
            iul_crc_index.setdefault(e.crc_hex.upper(), []).append(k)

    for f in ifc_files:
        base = f.name
        e = iul_map.get(base)
        actual_crc_hex = f"{compute_crc32(f):08X}"
        actual_size = f.stat().st_size
        actual_dt = _fmt_mtime(f.stat().st_mtime)

        name_match = None
        crc_match = None
        size_match = None
        dt_match = None
        pdf_name_ok = None
        status: List[str] = []
        details: List[str] = []

        if e is None:
            hits = iul_crc_index.get(actual_crc_hex, [])
            if len(hits) == 1:
                e = iul_map[hits[0]]
                used.add(hits[0])
                name_match = (e.basename == base)
                if not name_match:
                    status.append("NAME_MISMATCH")
                    details.append(
                        f"Сопоставлено по CRC-32, имя различается; ожидается {base}"
                    )
                crc_match = True
            elif len(hits) > 1:
                status.append("ERROR_IFC_EXTRA")
                details.append(f"Найдено несколько записей в ИУЛ с тем же CRC ({actual_crc_hex})")
            else:
                status.append("ERROR_IFC_EXTRA")
                details.append(f"Файл есть, но отсутствует запись в ИУЛ; ожидается запись для {base}")
        else:
            used.add(base)
            name_match = True
            if e.crc_hex:
                crc_match = (e.crc_hex.upper() == actual_crc_hex)
                if not crc_match:
                    status.append("CRC_MISMATCH")
                    details.append(f"CRC-32 не совпадает: ИУЛ={e.crc_hex.upper()}, IFC={actual_crc_hex}")
            else:
                details.append(f"В ИУЛ отсутствует CRC-32; ожидается {actual_crc_hex}")

        if e:
            if e.size_bytes is not None:
                size_match = (e.size_bytes == actual_size)
                if not size_match:
                    status.append("SIZE_MISMATCH")
                    details.append(f"Размер не совпадает: ИУЛ={e.size_bytes}, IFC={actual_size}")
            else:
                details.append(f"В ИУЛ отсутствует размер файла; ожидается {actual_size}")

            if e.dt_str:
                dt_match = (e.dt_str == actual_dt)
                if not dt_match:
                    status.append("DT_MISMATCH")
                    details.append(f"Дата/время не совпадает: ИУЛ={e.dt_str}, IFC={actual_dt}")
            else:
                details.append(f"В ИУЛ отсутствует дата/время; ожидается {actual_dt}")

            if e.source_pdf:
                pdf_name_ok = pdf_name_ok_strict(base, e.source_pdf) if strict_pdf_name else pdf_name_ok_lenient(base, e.source_pdf)
                if not pdf_name_ok:
                    status.append("PDF_NAME_MISMATCH")
                    rule = "строгому шаблону (имяIFC_УЛ.pdf)" if strict_pdf_name else "мягкому шаблону (содержит ИУЛ/УЛ и имя IFC)"
                    expected_pdf = f"{Path(base).stem}_УЛ.pdf"
                    details.append(
                        f"Имя PDF не соответствует правилу: {e.source_pdf}; ожидается {expected_pdf} [{rule}]"
                    )

        if not status and e is not None:
            status.append("OK")

        expected_pdf_name = f"{Path(base).stem}_УЛ.pdf"
        row = {
            "Имя файла": base,
            "Имя PDF": (e.source_pdf if e else None),
            "Файл из ИУЛ": (e.basename if e else None),
            "CRC-32 ИУЛ": (e.crc_hex.upper() if (e and e.crc_hex) else None),
            "CRC-32 IFC": actual_crc_hex,
            "Дата/время ИУЛ": (e.dt_str if e else None),
            "Дата/время IFC": actual_dt,
            "Размер ИУЛ, байт": (e.size_bytes if e else None),
            "Размер IFC, байт": actual_size,
            "Имя совпадает": tri(name_match),
            "CRC совпадает": tri(crc_match),
            "Дата/время совпадает": tri(dt_match),
            "Размер совпадает": tri(size_match),
            "Статус": ";".join(status) if status else "—",
            "Подробности": "; ".join(details) if details else None,
            "recommendation": None,
        }
        rec = recommendation(status, RECOMMENDATIONS)
        if rec:
            rec = rec.format(expected=expected_pdf_name)
        row["recommendation"] = rec
        if include_pdf_name_col:
            row[PDF_NAME_COL] = tri(pdf_name_ok)
        rows.append(row)

    for k, e in iul_map.items():
        if k in used:
            continue
        row = {
            "Имя файла": None,
            "Имя PDF": e.source_pdf,
            "Файл из ИУЛ": e.basename,
            "CRC-32 ИУЛ": (e.crc_hex.upper() if e.crc_hex else None),
            "CRC-32 IFC": None,
            "Дата/время ИУЛ": e.dt_str,
            "Дата/время IFC": None,
            "Размер ИУЛ, байт": e.size_bytes,
            "Размер IFC, байт": None,
            "Имя совпадает": "—",
            "CRC совпадает": "—",
            "Дата/время совпадает": "—",
            "Размер совпадает": "—",
            "Статус": "ERROR_IUL_EXTRA",
            "Подробности": f"Запись в ИУЛ есть, соответствующий файл не найден; ожидается файл {e.basename}",
            "recommendation": RECOMMENDATIONS.get("ERROR_IUL_EXTRA"),
        }
        if include_pdf_name_col:
            row[PDF_NAME_COL] = "—"
        rows.append(row)
    return rows
