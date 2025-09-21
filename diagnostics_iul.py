#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Утилита для диагностики OCR и разбора отчётов ИУЛ."""
from __future__ import annotations

import importlib
import os
import subprocess
import sys
import textwrap
import traceback
from pathlib import Path
from types import ModuleType
from typing import Iterable, Optional, Tuple

REQUIRED_MODULES: Tuple[Tuple[str, str], ...] = (
    ("PyPDF2", "PyPDF2"),
    ("pymupdf", "fitz"),
    ("pytesseract", "pytesseract"),
    ("Pillow", "PIL.Image"),
)


def _print_heading(title: str) -> None:
    print("\n" + title)
    print("-" * len(title))


def _import_module(dotted_name: str) -> Optional[ModuleType]:
    try:
        module = importlib.import_module(dotted_name)
    except Exception:
        print(f"[ОШИБКА] Не удалось импортировать {dotted_name}:")
        traceback.print_exc()
        return None
    else:
        print(f"[OK] {dotted_name}")
        return module


def _check_required_modules() -> None:
    _print_heading("Проверка доступности модулей")
    for _, import_name in REQUIRED_MODULES:
        _import_module(import_name)


def _diagnose_pkg_import() -> ModuleType | None:
    _print_heading("Проверка инициализации pkg.iul_reader")
    try:
        from pkg import iul_reader  # комментарий на русском языке
    except Exception:
        print("[ОШИБКА] Не удалось импортировать pkg.iul_reader:")
        traceback.print_exc()
        return None

    pdf_reader_ok = getattr(iul_reader, "PdfReader", None) is not None
    fitz_ok = getattr(iul_reader, "fitz", None) is not None
    pytesseract_ok = getattr(iul_reader, "pytesseract", None) is not None
    image_ok = getattr(iul_reader, "Image", None) is not None

    print(f"PdfReader доступен: {pdf_reader_ok}")
    print(f"PyMuPDF (fitz) доступен: {fitz_ok}")
    print(f"pytesseract доступен: {pytesseract_ok}")
    print(f"Pillow.Image доступен: {image_ok}")
    tesseract_cmd = getattr(iul_reader.pytesseract, "tesseract_cmd", "") if pytesseract_ok else ""
    print(f"Путь, который использует модуль, {tesseract_cmd or '(не задан)'}")

    return iul_reader


def _check_tesseract_configuration(iul_reader: ModuleType | None) -> Path:
    _print_heading("Проверка настроек Tesseract")
    pytesseract = _import_module("pytesseract")
    if pytesseract is None:
        return Path()

    configured_cmd = ""
    if iul_reader is not None:
        module_pytesseract = getattr(iul_reader, "pytesseract", None)
        if module_pytesseract is not None:
            candidates = [
                getattr(module_pytesseract, "tesseract_cmd", ""),
                getattr(getattr(module_pytesseract, "pytesseract", None), "tesseract_cmd", ""),
            ]
            for candidate in candidates:
                if not candidate:
                    continue
                configured_cmd = str(candidate).strip()
                if configured_cmd:
                    break
            if configured_cmd:
                print(f"pkg.iul_reader сообщает путь к tesseract: {configured_cmd}")
                current_cmd = str(getattr(pytesseract, "tesseract_cmd", "")).strip()
                if not current_cmd:
                    setattr(pytesseract, "tesseract_cmd", configured_cmd)

    tesseract_cmd = configured_cmd or str(getattr(pytesseract, "tesseract_cmd", "")).strip()
    print(f"Путь в pytesseract.tesseract_cmd: {tesseract_cmd or '(не задан)'}")
    tessdata_prefix = os.environ.get("TESSDATA_PREFIX")
    print(f"Значение переменной TESSDATA_PREFIX: {tessdata_prefix or '(не задана)'}")

    if tesseract_cmd:
        exe_path = Path(tesseract_cmd)
    else:
        exe_path = Path()
    if exe_path.is_file():
        print(f"Бинарник tesseract найден: {exe_path}")
    else:
        print("[ВНИМАНИЕ] Бинарник tesseract по указанному пути не найден")
    return exe_path


def _list_tesseract_languages(exe_path: Path) -> None:
    _print_heading("Проверка доступных языков Tesseract")
    if not exe_path.is_file():
        print("[ВНИМАНИЕ] Путь к tesseract не задан, пропускаем проверку языков")
        return
    print(f"Используем бинарник tesseract: {exe_path}")
    try:
        result = subprocess.run(
            [str(exe_path), "--list-langs"],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        print("[ОШИБКА] Не удалось запустить tesseract для списка языков:")
        traceback.print_exc()
        return
    if result.stdout:
        print("Доступные языки:\n" + result.stdout.strip())
    if result.stderr.strip():
        print("Сообщения об ошибках:\n" + result.stderr.strip())


def _read_text_preview(text: str, max_lines: int = 10) -> str:
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        return "(пусто)"
    preview = lines[:max_lines]
    return "\n".join(preview)


def _analyze_pdf(pdf_path: Path) -> None:
    _print_heading(f"Диагностика PDF: {pdf_path}")
    if not pdf_path.exists():
        print("[ОШИБКА] Файл не найден")
        return
    if not pdf_path.is_file():
        print("[ОШИБКА] Указанный путь не является файлом")
        return

    try:
        from pkg import iul_reader  # комментарий на русском языке
    except Exception:
        print("[ОШИБКА] Модуль pkg.iul_reader недоступен, анализ невозможен")
        return

    text_pypdf2 = iul_reader._extract_text_pypdf2(pdf_path)
    text_pypdf2_norm = iul_reader._normalize_text(text_pypdf2)
    print("Длина текста, полученного из PyPDF2:", len(text_pypdf2_norm))
    print("Предпросмотр (первые строки):")
    print(textwrap.indent(_read_text_preview(text_pypdf2_norm), prefix="  "))

    if len(text_pypdf2_norm) == 0:
        print("PyPDF2 не вернул текст, пробуем OCR...")

    text_ocr = iul_reader._extract_text_ocr(pdf_path)
    text_ocr_norm = iul_reader._normalize_text(text_ocr)
    print("Длина текста после OCR:", len(text_ocr_norm))
    print("Предпросмотр (первые строки):")
    print(textwrap.indent(_read_text_preview(text_ocr_norm), prefix="  "))

    if text_ocr_norm:
        print("Строки, содержащие .ifc:")
        for line in text_ocr_norm.splitlines():
            if ".ifc" in line.lower():
                print(textwrap.indent(line, prefix="  "))
    else:
        print("OCR не вернул ни одной непустой строки")


def _parse_arguments(argv: Iterable[str]) -> Path:
    args = list(argv)
    if not args:
        return Path()
    return Path(args[0]).expanduser()


def main(argv: Iterable[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    pdf_path = _parse_arguments(argv)

    _print_heading("Сведения об интерпретаторе Python")
    print(sys.version)
    print(f"Исполняемый файл Python: {sys.executable}")

    iul_reader = _diagnose_pkg_import()
    _check_required_modules()
    exe_path = _check_tesseract_configuration(iul_reader)
    _list_tesseract_languages(exe_path)

    if pdf_path:
        _analyze_pdf(pdf_path)
    else:
        _print_heading("Анализ PDF не выполнялся")
        print("Передайте путь к PDF в качестве аргумента, чтобы проверить распознавание")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
