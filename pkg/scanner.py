# -*- coding: utf-8 -*-
from pathlib import Path

IFC_EXTS = {".ifc"}
PDF_EXTS = {".pdf"}

def collect_ifc_files(folder: Path, recursive: bool = True) -> list[Path]:
    if not folder or not folder.exists() or not folder.is_dir():
        return []
    pattern = "**/*" if recursive else "*"
    files = [p for p in folder.glob(pattern) if p.is_file() and p.suffix.lower() in IFC_EXTS]
    return sorted({p.resolve() for p in files})

def collect_pdf_files(folder: Path, recursive: bool = True) -> list[Path]:
    if not folder or not folder.exists() or not folder.is_dir():
        return []
    pattern = "**/*" if recursive else "*"
    files = [p for p in folder.glob(pattern) if p.is_file() and p.suffix.lower() in PDF_EXTS]
    return sorted({p.resolve() for p in files})
