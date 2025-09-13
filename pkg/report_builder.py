# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Dict, List
from .crc import compute_crc32
from .utils import tri, recommendation


RECOMMENDATIONS = {
    "OK": "Действий не требуется",
    "ERROR_IFC_EXTRA": "Удалите лишний файл или добавьте запись в XML",
    "ERROR_XML_EXTRA": "Удалите лишнюю запись из XML или добавьте соответствующий файл IFC",
    "CRC_MISMATCH": "Проверьте корректность файлов и пересоздайте CRC",
    "NAME_MISMATCH": "Переименуйте файл или исправьте запись в XML",
}

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
        xml_crc_from_xml = None
        status: List[str] = []
        details: List[str] = []
        xml_name_from_xml = base if meta else None

        if meta is None:
            # пытаемся сопоставить по CRC
            hits = xml_crc_index.get(actual_crc_hex, [])
            if len(hits) == 1:
                xml_name = hits[0]
                xml_name_from_xml = xml_name
                meta_hit = xml_map.get(xml_name if case_sensitive else xml_name.lower())
                xml_crc_from_xml = ((meta_hit.get("crc_hex") or "").upper() if meta_hit else None)
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
            xml_crc_from_xml = (meta.get("crc_hex") or "").upper() or None
            if xml_crc_from_xml:
                crc_match = (xml_crc_from_xml == actual_crc_hex)
                if not crc_match:
                    status.append("CRC_MISMATCH")
                    details.append(f"CRC-32 не совпадает: XML={xml_crc_from_xml}, IFC={actual_crc_hex}")
            else:
                details.append("В XML отсутствует CRC-32")

        if not status and name_match is True and (crc_match is True or crc_match is None):
            status.append("OK")

        rows.append({
            "Имя файла IFC": base,
            "Имя файла IFC из XML": xml_name_from_xml,
            "CRC-32 XML": xml_crc_from_xml,
            "CRC-32 IFC": actual_crc_hex,
            "Имя совпадает": tri(name_match),
            "CRC совпадает": tri(crc_match),
            "Статус": ";".join(status) if status else "—",
            "Подробности": "; ".join(details) if details else None,
            "recommendation": recommendation(status, RECOMMENDATIONS),
        })

    # Лишние записи в XML
    for name, meta in xml_map.items():
        k = name if case_sensitive else name.lower()
        if k in used_xml:
            continue
        rows.append({
            "Имя файла IFC": None,
            "Имя файла IFC из XML": name,
            "CRC-32 XML": (meta.get("crc_hex") or "").upper() or None,
            "CRC-32 IFC": None,
            "Имя совпадает": "—",
            "CRC совпадает": "—",
            "Статус": "ERROR_XML_EXTRA",
            "Подробности": "Запись в XML есть, соответствующий файл не найден",
            "recommendation": RECOMMENDATIONS.get("ERROR_XML_EXTRA"),
        })

    return rows
