#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, sys
CURRENT_DIR = os.path.abspath(os.path.dirname(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from pathlib import Path
import argparse
import logging

from pkg.xml_reader import read_rules, extract_from_xml, extract_iul_pdf_pairs
from pkg.scanner import collect_ifc_files, collect_pdf_files
from pkg.report_builder import build_report
from pkg.xlsx_writer import write_xlsx

from pkg.report_builder_pdf_pairs import build_report_pdf_pairs
from pkg.xlsx_writer_pdf_pairs import write_xlsx_pdf_pairs

from pkg.iul_reader import extract_iul_entries, PdfReader  # type: ignore
from pkg.report_builder_iul import build_report_iul
from pkg.xlsx_writer_iul import write_xlsx_iul

def main():
    ap = argparse.ArgumentParser(description="IFC CRC Checker (CLI) — сверка XML↔IFC и/или ИУЛ(PDF)↔IFC с отчётами XLSX")
    ap.add_argument("--ifc-dir", type=Path, required=True, help="Папка с IFC-файлами")
    ap.add_argument("--recursive-ifc", action="store_true", help="Рекурсивно сканировать подпапки (IFC)")
    ap.add_argument("--out", type=Path, help="Куда сохранить .xlsx (XML), по умолчанию рядом с XML или в CWD")

    # XML
    ap.add_argument("--check-xml", action="store_true", help="Выполнить проверку XML↔IFC")
    ap.add_argument("--xml", type=Path, help="Путь к XML с перечнем IFC")

    # IUL ↔ IFC
    ap.add_argument("--check-iul", action="store_true", help="Выполнить проверку ИУЛ(PDF)↔IFC")

    # PDF pairs
    ap.add_argument("--check-pdf-pairs", action="store_true", help="Выполнить проверку ИУЛ PDF↔PDF из XML")
    ap.add_argument("--iul", type=Path, nargs="*", help="Пути к PDF. Можно несколько")
    ap.add_argument("--iul-dir", type=Path, help="Папка с PDF")
    ap.add_argument("--recursive-pdf", action="store_true", help="Рекурсивно сканировать подпапки (PDF)")
    ap.add_argument("--pdf-name-strict", action="store_true", help="Строгое правило имени PDF (…_УЛ.pdf)")

    ap.add_argument("--force", action="store_true", help="Перезаписать отчёты, если файлы уже существуют")
    ap.add_argument("-v", "--verbose", action="store_true", help="Подробные логи")

    args = ap.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s: %(message)s")

    if not (args.check_xml or args.check_iul or args.check_pdf_pairs):
        args.check_xml = True
        args.check_iul = True
        args.check_pdf_pairs = True

    if not args.ifc_dir.exists() or not args.ifc_dir.is_dir():
        logging.error("Папка с IFC не найдена/не является папкой: %s", args.ifc_dir); return 2
    ifc_files = collect_ifc_files(args.ifc_dir, recursive=args.recursive_ifc)
    if not ifc_files:
        logging.error("В папке не найдено файлов *.ifc"); return 2

    pdfs: list[Path] = []
    if args.iul:
        pdfs.extend(args.iul)
    if args.iul_dir and args.iul_dir.exists():
        pdfs.extend(collect_pdf_files(args.iul_dir, recursive=args.recursive_pdf))

    # XML
    if args.check_xml:
        if not args.xml or not args.xml.exists():
            logging.error("Указана проверка XML, но путь к XML не задан или файл не найден."); return 2
        out_xml = args.out or args.xml.with_name("ifc_crc_report.xlsx")
        if out_xml.exists() and not args.force:
            logging.error("Файл отчёта (XML) уже существует: %s. Запустите с --force для перезаписи.", out_xml); return 2
        rules = read_rules(Path(__file__).with_name("rules.yaml"))
        xml_map = extract_from_xml(args.xml, rules, case_sensitive=True)
        rows_xml = build_report(xml_map, ifc_files, case_sensitive=True)
        exit_xml, stats_xml = write_xlsx(rows_xml, out_xml)
        logging.info("Готово (XML). Отчёт: %s | Итоги: %s", out_xml, stats_xml)

    # IUL
    if args.check_iul:
        if not pdfs:
            logging.error("Указана проверка ИУЛ, но PDF не заданы/не найдены."); return 2
        if PdfReader is None:
            logging.error("Для чтения ИУЛ (PDF) требуется PyPDF2. Установите зависимости."); return 2
        if args.xml:
            out_iul = (args.out or args.xml.with_name("ifc_crc_report.xlsx"))
            out_iul = out_iul.with_name(out_iul.stem.replace('.xlsx','') + "_iul.xlsx")
        else:
            out_iul = Path.cwd() / "ifc_crc_report_iul.xlsx"
        if out_iul.exists() and not args.force:
            logging.error("Файл отчёта (IUL) уже существует: %s. Запустите с --force для перезаписи.", out_iul); return 2
        iul_map = extract_iul_entries(pdfs)
        rows_iul = build_report_iul(iul_map, ifc_files, strict_pdf_name=bool(args.pdf_name_strict))
        exit_iul, stats_iul = write_xlsx_iul(rows_iul, out_iul)
        logging.info("Готово (IUL). Отчёт: %s | Итоги: %s", out_iul, stats_iul)

    if args.check_pdf_pairs:
        if not args.xml or not args.xml.exists():
            logging.error("Указана проверка PDF-пар, но XML не задан или не найден."); return 2
        if not pdfs:
            logging.error("Указана проверка PDF-пар, но PDF не заданы/не найдены."); return 2
        out_pairs = (args.out or args.xml.with_name("pdf_pair_report.xlsx"))
        if out_pairs.exists() and not args.force:
            logging.error("Файл отчёта (PDF пары) уже существует: %s. Запустите с --force для перезаписи.", out_pairs); return 2
        pairs = extract_iul_pdf_pairs(args.xml)
        rows_pairs = build_report_pdf_pairs(pairs, pdfs)
        exit_pairs, stats_pairs = write_xlsx_pdf_pairs(rows_pairs, out_pairs)
        logging.info("Готово (PDF пары). Отчёт: %s | Итоги: %s", out_pairs, stats_pairs)

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
