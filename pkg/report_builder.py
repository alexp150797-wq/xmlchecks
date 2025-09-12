# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, List
from .crc import compute_crc32


RECOMMENDATIONS = {
    "OK": "Действий не требуется",
    "ERROR_IFC_EXTRA": "Удалите лишний файл или добавьте запись в XML",
    "ERROR_XML_EXTRA": "Удалите лишнюю запись из XML или добавьте соответствующий файл IFC",
    "CRC_MISMATCH": "Проверьте корректность файлов и пересоздайте CRC",
    "NAME_MISMATCH": "Переименуйте файл или исправьте запись в XML",
}

def _tri(v: bool | None) -> str:
    return "Да" if v is True else "Нет" if v is False else "—"


def _recommendation(status: List[str]) -> str | None:
    recs = [RECOMMENDATIONS.get(s) for s in status if RECOMMENDATIONS.get(s)]
    return "; ".join(recs) if recs else None

def build_report(xml_map: Dict[str, dict], ifc_files: List[Path], case_sensitive: bool=True) -> List[Dict]:
    """
    Сравнение XML↔IFC:
      - Имя (строгое сравнение)
      - CRC-32
    Сценарии:
      - IFC есть, записи в XML нет → ERROR_IFC_EXTRA
      - Запись в XML есть, IFC не найден → ERROR_XML_EXTRA
      - CRC разные → CRC_MISMATCH
      - Есть совпадение по CRC, но имя отличается → NAME_MISMATCH (в одну строку)
      - Всё ок → OK
    """
    rows: List[Dict] = []
    used_xml = set()

    # Индекс: CRC из XML -> список имён
    xml_crc_index: Dict[str, List[str]] = {}
    for name, meta in xml_map.items():
        crc = (meta.get("crc_hex") or "").upper()
        if crc:
            xml_crc_index.setdefault(crc, []).append(name)

    for f in ifc_files:
        base = f.name
        key = base if case_sensitive else base.lower()
        meta = xml_map.get(key)
        actual_crc_hex = f"{compute_crc32(f):08X}"

        name_match = None
        crc_match = None
        status: List[str] = []
        details: List[str] = []

        if meta is None:
            # пытаемся сопоставить по CRC
            hits = xml_crc_index.get(actual_crc_hex, [])
            if len(hits) == 1:
                xml_name = hits[0]
                used_xml.add(xml_name if case_sensitive else xml_name.lower())
                name_match = (xml_name == base)
                if not name_match:
                    status.append("NAME_MISMATCH")
                    details.append("Сопоставлено по CRC-32, имя различается")
                crc_match = True
            elif len(hits) > 1:
                status.append("ERROR_IFC_EXTRA")
                details.append(f"Найдено несколько записей в XML с тем же CRC ({actual_crc_hex})")
            else:
                status.append("ERROR_IFC_EXTRA")
                details.append("Файл есть, но отсутствует запись в XML")
        else:
            used_xml.add(key)
            name_match = True
            xml_crc = (meta.get("crc_hex") or "").upper() or None
            if xml_crc:
                crc_match = (xml_crc == actual_crc_hex)
                if not crc_match:
                    status.append("CRC_MISMATCH")
                    details.append(f"CRC-32 не совпадает: XML={xml_crc}, IFC={actual_crc_hex}")
            else:
                details.append("В XML отсутствует CRC-32")

        if not status and name_match is True and (crc_match is True or crc_match is None):
            status.append("OK")

        rows.append({
            "Имя файла": base,
            "Файл из XML": (base if meta else (hits[0] if 'hits' in locals() and hits else None)),
            "CRC-32 XML": ((meta.get('crc_hex') or '').upper() if meta else (None)),
            "CRC-32 IFC": actual_crc_hex,
            "Имя совпадает": _tri(name_match),
            "CRC совпадает": _tri(crc_match),
            "Статус": ";".join(status) if status else "—",
            "Подробности": "; ".join(details) if details else None,
            "recommendation": _recommendation(status),
        })

    # Лишние записи в XML
    for name, meta in xml_map.items():
        k = name if case_sensitive else name.lower()
        if k in used_xml:
            continue
        rows.append({
            "Имя файла": None,
            "Файл из XML": name,
            "CRC-32 XML": (meta.get("crc_hex") or "").upper() or None,
            "CRC-32 IFC": None,
            "Имя совпадает": "—",
            "CRC совпадает": "—",
            "Статус": "ERROR_XML_EXTRA",
            "Подробности": "Запись в XML есть, соответствующий файл не найден",
            "recommendation": RECOMMENDATIONS.get("ERROR_XML_EXTRA"),
        })

    return rows
