# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import zlib

def compute_crc32(path: Path, chunk_size: int = 1024 * 1024) -> int:
    """
    Вычисляет CRC-32 файла (unsigned), совпадает со значением, которое ждём в XML/ИУЛ.
    Возвращает int (0..2^32-1). Представление в hex: f"{crc:08X}".
    """
    crc = 0
    with path.open("rb") as f:
        while True:
            buf = f.read(chunk_size)
            if not buf:
                break
            crc = zlib.crc32(buf, crc)
    return crc & 0xFFFFFFFF
