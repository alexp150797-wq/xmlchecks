# -*- coding: utf-8 -*-
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Optional, Callable
import re

try:
    from PyPDF2 import PdfReader  # type: ignore
except Exception:
    PdfReader = None

try:
    import fitz  # PyMuPDF  # type: ignore
except Exception:
    fitz = None

try:
    import pytesseract  # type: ignore
    from PIL import Image  # type: ignore
except Exception:
    pytesseract = None
    Image = None

CRC_RE = re.compile(r"CRC[-\s_]*32\s*([0-9A-Fa-f]{8})")
IFC_RE = re.compile(r"([\w\-. ]+?\.ifc)", re.IGNORECASE)
DT_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2})")
SIZE_RE = re.compile(r"Размер\s+файла\D*(\d+)", re.IGNORECASE)
# Обособленное вхождение "УЛ" или "ИУЛ" (\s, _)
IUL_KEYWORD_RE = re.compile(r"(^|[\s_])(ИУЛ|УЛ)([\s_]|$)")

@dataclass
class IulEntry:
    basename: str
    crc_hex: Optional[str]
    dt_str: Optional[str]
    size_bytes: Optional[int]
    context: str
    source_pdf: str

def _extract_text_pypdf2(pdf_path: Path) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(str(pdf_path))
        parts = []
        for p in reader.pages:
            try:
                parts.append(p.extract_text() or "")
            except Exception:
                parts.append("")
        return "\n".join(parts)
    except Exception:
        return ""

def _extract_text_ocr(pdf_path: Path, dpi: int = 200) -> str:
    if fitz is None or pytesseract is None or Image is None:
        return ""
    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return ""
    text_parts: List[str] = []
    for page in doc:
        try:
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            txt = pytesseract.image_to_string(img, lang="rus+eng")
            if txt:
                text_parts.append(txt)
        except Exception:
            continue
    return "\n".join(text_parts)

def _normalize_text(txt: str) -> str:
    txt = txt.replace("\r", "\n")
    return "\n".join(ln.strip() for ln in txt.splitlines())

def _parse_entries(text: str, pdf_name: str, progress: Optional[Callable[[IulEntry], None]] = None) -> List[IulEntry]:
    lines = [ln for ln in text.splitlines() if ln]
    entries: List[IulEntry] = []
    last_crc: Optional[str] = None
    for ln in lines:
        m_crc = CRC_RE.search(ln)
        if m_crc:
            last_crc = m_crc.group(1).upper()

        if ".ifc" in ln or ".IFC" in ln:
            m_ifc = IFC_RE.search(ln)
            if not m_ifc:
                continue
            fname = Path(m_ifc.group(1)).name
            m_dt = DT_RE.search(ln)
            dt = m_dt.group(1) if m_dt else None
            size = None

            m_size = SIZE_RE.search(ln)
            if m_size:
                size = int(m_size.group(1))
            else:
                tail = ln[m_ifc.end():]
                ints = [int(x) for x in re.findall(r"\d+", tail)]
                if ints:
                    size = ints[-1]

            entry = IulEntry(
                basename=fname,
                crc_hex=(last_crc or None),
                dt_str=dt,
                size_bytes=size,
                context=ln,
                source_pdf=pdf_name,
            )
            entries.append(entry)
            if progress:
                try:
                    progress(entry)
                except Exception:
                    pass
    return entries


def extract_iul_entries_from_pdf(pdf_path: Path, progress: Optional[Callable[[IulEntry], None]] = None) -> List[IulEntry]:
    text = _extract_text_pypdf2(pdf_path)
    text = _normalize_text(text)
    entries = _parse_entries(text, pdf_path.name, progress)
    if not entries:
        text_ocr = _extract_text_ocr(pdf_path)
        text_ocr = _normalize_text(text_ocr)
        entries = _parse_entries(text_ocr, pdf_path.name, progress)
    return entries

def extract_iul_entries(
    paths: List[Path],
    progress: Optional[Callable[[IulEntry], None]] = None,
) -> Dict[str, IulEntry]:
    res: Dict[str, IulEntry] = {}
    for p in paths:
        for e in extract_iul_entries_from_pdf(p, progress=progress):
            key = e.basename
            if key not in res:
                res[key] = e
    return res

def pdf_name_ok_lenient(ifc_name: str, pdf_name: str) -> bool:
    ifc_stem = Path(ifc_name).stem.upper()
    stem = Path(pdf_name).stem.upper()
    if ifc_stem not in stem:
        return False
    return bool(IUL_KEYWORD_RE.search(stem))

def pdf_name_ok_strict(ifc_name: str, pdf_name: str) -> bool:
    ifc_stem = Path(ifc_name).stem.upper()
    stem = Path(pdf_name).stem.upper()
    return (ifc_stem in stem) and stem.endswith("_УЛ")
