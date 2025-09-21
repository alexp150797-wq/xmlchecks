#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Утилита для проверки распознавания текста в PDF.

Скрипт извлекает текст из PDF двумя способами: через PyPDF2 и через OCR
(PyMuPDF + pytesseract). Полученный текст выводится в консоль, что позволяет
оперативно понять, сработал ли OCR и какой результат он дал.

Запуск из корня проекта (можно указать несколько файлов):

```
python pdf_ocr_debug.py path/to/file1.pdf path/to/file2.pdf
```
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict

from pkg.iul_reader import extract_pdf_text_debug


def _print_preview(text: str, *, max_lines: int) -> None:
    """Выводит первые ``max_lines`` строк текста."""

    if not text:
        return

    lines = text.splitlines()
    preview = lines[:max_lines]
    for ln in preview:
        print(ln)
    if len(lines) > max_lines:
        print("...")


def _print_source_result(name: str, payload: Dict[str, str], *, max_lines: int, show_raw: bool) -> None:
    raw_text = payload.get("raw", "")
    normalized = payload.get("normalized", "")
    chars_raw = len(raw_text)
    chars_normalized = len(normalized)

    print(f"  [{name}] символов (raw/normalized): {chars_raw}/{chars_normalized}")

    text_to_show = raw_text if show_raw else normalized
    _print_preview(text_to_show, max_lines=max_lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Проверка распознавания текста в PDF. Можно сравнить результат PyPDF2"
            " и OCR, чтобы понять, нужен ли дополнительный тюнинг."
        )
    )
    parser.add_argument("pdf", nargs="+", type=Path, help="Пути к PDF-файлам для проверки")
    parser.add_argument("--dpi", type=int, default=300, help="DPI для OCR (по умолчанию 300)")
    parser.add_argument(
        "--no-pypdf2",
        action="store_true",
        help="Не использовать извлечение текста через PyPDF2",
    )
    parser.add_argument(
        "--no-ocr",
        action="store_true",
        help="Не выполнять OCR (PyMuPDF + pytesseract)",
    )
    parser.add_argument(
        "--preview-lines",
        type=int,
        default=40,
        help="Сколько строк текста выводить в качестве превью (по умолчанию 40)",
    )
    parser.add_argument(
        "--show-raw",
        action="store_true",
        help="Показывать «сырой» текст вместо нормализованного",
    )

    args = parser.parse_args(argv)

    if args.no_pypdf2 and args.no_ocr:
        parser.error("некуда извлекать текст: отключены и PyPDF2, и OCR")

    status = 0
    for pdf_path in args.pdf:
        print(f"=== {pdf_path} ===")
        if not pdf_path.exists():
            print("  [ошибка] файл не найден")
            status = 1
            continue
        if not pdf_path.is_file():
            print("  [ошибка] указан не файл")
            status = 1
            continue

        try:
            payload = extract_pdf_text_debug(
                pdf_path,
                dpi=args.dpi,
                include_pypdf2=not args.no_pypdf2,
                include_ocr=not args.no_ocr,
            )
        except Exception as exc:  # pragma: no cover - диагностическое сообщение
            print(f"  [ошибка] не удалось обработать PDF: {exc}")
            status = 1
            continue

        if not payload:
            print("  [предупреждение] нет данных для вывода")
            continue

        for name, data in payload.items():
            _print_source_result(
                name,
                data,
                max_lines=max(1, args.preview_lines),
                show_raw=args.show_raw,
            )

    return status


if __name__ == "__main__":  # pragma: no cover - точка входа для CLI
    raise SystemExit(main())
