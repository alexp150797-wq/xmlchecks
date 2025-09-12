#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, sys
CURRENT_DIR = os.path.abspath(os.path.dirname(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from pathlib import Path
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pkg.xml_reader import read_rules, extract_from_xml
from pkg.scanner import collect_ifc_files, collect_pdf_files
from pkg.report_builder import build_report
from pkg.xlsx_writer import write_xlsx

from pkg.iul_reader import (
    extract_iul_entries,
    PdfReader,
    set_tesseract_cmd,
    set_ocr_langs,
)  # type: ignore
from pkg.report_builder_iul import build_report_iul
from pkg.xlsx_writer_iul import write_xlsx_iul

APP_TITLE = "IFC CRC Checker — GUI"
APP_MIN_W, APP_MIN_H = 980, 760

EMOJI = {
    "xml": "🧾",
    "ifc": "🏗️",
    "iul": "📄",
    "search": "🔎",
    "xlsx": "📊",
    "done": "✅",
    "report": "📝",
    "stats": "🧮",
    "path": "📁",
    "ok": "✅",
    "err": "❌",
    "warn": "⚠️",
}

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(APP_MIN_W, APP_MIN_H)
        self.geometry(f"{APP_MIN_W}x{APP_MIN_H}")
        self._apply_theme()

        # Paths & options
        self.var_xml = tk.StringVar()
        self.var_ifc_dir = tk.StringVar()
        self.var_out = tk.StringVar(value=str(Path.cwd() / "ifc_crc_report.xlsx"))
        self.var_recursive_ifc = tk.BooleanVar(value=True)
        self.var_open_after = tk.BooleanVar(value=True)

        # Checks: independently
        self.var_check_xml = tk.BooleanVar(value=True)
        self.var_check_iul = tk.BooleanVar(value=True)

        # IUL selection
        self.iul_files: list[Path] = []
        self.var_iul_label = tk.StringVar(value="(не выбрано)")
        self.var_iul_dir = tk.StringVar()
        self.var_recursive_pdf = tk.BooleanVar(value=True)
        self.var_pdf_name_strict = tk.BooleanVar(value=False)  # строгое имя PDF (_УЛ)
        self.var_tesseract_path = tk.StringVar()
        self.var_ocr_lang = tk.StringVar(value="rus+eng")

        self._build_ui()

    def _apply_theme(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure("TButton", padding=8)
        self.style.configure("TLabel", padding=2)
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        self.style.configure("Small.TLabel", foreground="#666")
        self.style.configure("Accent.TButton", padding=10)
        self.style.map("Accent.TButton", background=[("active", "#007ACC")])

    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        header = ttk.Frame(self)
        header.pack(fill="x")
        ttk.Label(header, text="Сверка: XML↔IFC и ИУЛ(PDF)↔IFC → XLSX", style="Header.TLabel").pack(side="left", **pad)
        ttk.Label(header, text="Зелёный — успех, красный — ошибка. Сравнение имён всегда строгое.", style="Small.TLabel").pack(side="right", **pad)

        body = ttk.Frame(self); body.pack(fill="both", expand=True, padx=10, pady=6)

        # What to check
        box = ttk.LabelFrame(body, text="Что проверять")
        box.grid(row=0, column=0, columnspan=3, sticky="we", **pad)
        ttk.Checkbutton(box, text="XML ↔ IFC", variable=self.var_check_xml).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(box, text="ИУЛ (PDF) ↔ IFC", variable=self.var_check_iul).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(box, text="Открыть отчёты по завершению", variable=self.var_open_after).grid(row=0, column=2, sticky="w", padx=8, pady=4)

        # XML
        ttk.Label(body, text=f"{EMOJI['xml']} XML:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(body, textvariable=self.var_xml).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(body, text="Файл...", command=self._choose_xml).grid(row=1, column=2, **pad)

        # IFC dir
        ttk.Label(body, text=f"{EMOJI['ifc']} Папка с IFC:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(body, textvariable=self.var_ifc_dir).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(body, text="Выбрать...", command=self._choose_ifc_dir).grid(row=2, column=2, **pad)
        ttk.Checkbutton(body, text="Рекурсивно по IFC", variable=self.var_recursive_ifc).grid(row=2, column=3, sticky="w", padx=6)

        # IUL PDFs
        iul_frame = ttk.LabelFrame(body, text="ИУЛ (PDF)")
        iul_frame.grid(row=3, column=0, columnspan=4, sticky="we", padx=10, pady=6)
        ttk.Label(iul_frame, text=f"{EMOJI['iul']} Выбор:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Button(iul_frame, text="PDF-файлы...", command=self._choose_iul_files).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(iul_frame, text="Папка с PDF...", command=self._choose_iul_dir).grid(row=0, column=2, padx=6, pady=4)
        ttk.Label(iul_frame, textvariable=self.var_iul_label).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(iul_frame, text="Рекурсивно по PDF", variable=self.var_recursive_pdf).grid(row=0, column=4, sticky="w", padx=6)

        ttk.Checkbutton(iul_frame, text="Строгое имя PDF (…_УЛ.pdf)", variable=self.var_pdf_name_strict).grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Button(iul_frame, text="Очистить выбор PDF", command=self._clear_iul).grid(row=1, column=2, padx=6, pady=4)

        ttk.Label(iul_frame, text=f"{EMOJI['path']} Tesseract:").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(iul_frame, textvariable=self.var_tesseract_path).grid(row=2, column=1, columnspan=2, sticky="ew", padx=6, pady=4)
        ttk.Button(iul_frame, text="Файл...", command=self._choose_tesseract).grid(row=2, column=3, padx=6, pady=4)

        ttk.Label(iul_frame, text="Языки OCR:").grid(row=3, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(iul_frame, textvariable=self.var_ocr_lang).grid(row=3, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        iul_frame.columnconfigure(1, weight=1)

        # Out
        ttk.Label(body, text=f"{EMOJI['xlsx']} Отчёт (XML):").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(body, textvariable=self.var_out).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Button(body, text="Сохранить как...", command=self._choose_out).grid(row=4, column=2, **pad)

        # Run
        btns = ttk.Frame(body); btns.grid(row=5, column=0, columnspan=4, sticky="w", **pad)
        ttk.Button(btns, text=f"{EMOJI['search']} Сформировать отчёт(ы)", style="Accent.TButton", command=self._run).pack(side="left", padx=6)
        ttk.Button(btns, text="Выход", command=self.destroy).pack(side="left", padx=6)

        # Progress + log
        ttk.Separator(body).grid(row=6, column=0, columnspan=4, sticky="ew", padx=10, pady=4)
        ttk.Label(body, text="Журнал:").grid(row=7, column=0, sticky="w", **pad)
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.grid(row=7, column=1, columnspan=3, sticky="ew", padx=10)
        self.log = tk.Text(body, height=18, wrap="word")
        self.log.grid(row=8, column=0, columnspan=4, sticky="nsew", padx=10, pady=(0,10))

        # Log color tags
        self.log.tag_configure("ok", foreground="#0a8a0a")
        self.log.tag_configure("err", foreground="#c0392b")
        self.log.tag_configure("info", foreground="#1f4e79")
        self.log.tag_configure("warn", foreground="#9a6a00")

        body.columnconfigure(1, weight=1)
        body.rowconfigure(8, weight=1)

    def _choose_xml(self):
        p = filedialog.askopenfilename(title="Выберите XML", filetypes=[("XML","*.xml"),("Все файлы","*.*")])
        if p:
            self.var_xml.set(p)
            if not self.var_out.get():
                self.var_out.set(str(Path(p).with_name("ifc_crc_report.xlsx")))

    def _choose_ifc_dir(self):
        p = filedialog.askdirectory(title="Выберите папку с IFC")
        if p:
            self.var_ifc_dir.set(p)

    def _choose_iul_files(self):
        ps = filedialog.askopenfilenames(title="Выберите ИУЛ (PDF)", filetypes=[("PDF","*.pdf")])
        if ps:
            self.iul_files = [Path(p) for p in ps]
            self._update_iul_label()
        else:
            self._update_iul_label()

    def _choose_iul_dir(self):
        p = filedialog.askdirectory(title="Выберите папку с ИУЛ (PDF)")
        if p:
            self.var_iul_dir.set(p)
            self._update_iul_label()

    def _choose_tesseract(self):
        p = filedialog.askopenfilename(title="Выберите tesseract", filetypes=[("Все файлы","*.*")])
        if p:
            self.var_tesseract_path.set(p)

    def _clear_iul(self):
        self.iul_files = []
        self.var_iul_dir.set("")
        self._update_iul_label()

    def _update_iul_label(self):
        parts = []
        if self.iul_files:
            parts.append(f"файлов: {len(self.iul_files)}")
        if self.var_iul_dir.get():
            parts.append(f"папка: {self.var_iul_dir.get()}")
        self.var_iul_label.set(", ".join(parts) if parts else "(не выбрано)")

    def _choose_out(self):
        p = filedialog.asksaveasfilename(
            title="Сохранить отчёт (XML) как...",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook","*.xlsx")],
            initialfile="ifc_crc_report.xlsx"
        )
        if p:
            self.var_out.set(p)

    def _log(self, msg, kind="info"):
        tag = "info" if kind not in ("ok","err","warn") else kind
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end"); self.update_idletasks()

    def _open_path(self, path: Path):
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            self._log(f"{EMOJI['err']} Не удалось открыть файл: {e}", "err")

    def _ask_overwrite(self, path: Path) -> bool:
        if not path.exists():
            return True
        return messagebox.askyesno("Файл существует", f"Файл:\n{path}\nуже существует.\nЗаменить?")

    def _run(self):
        try:
            self.log.delete("1.0", "end")

            check_xml = bool(self.var_check_xml.get())
            check_iul = bool(self.var_check_iul.get())
            if not (check_xml or check_iul):
                self._log(f"{EMOJI['warn']} Ничего не выбрано для проверки.", "warn")
                return

            xml = Path(self.var_xml.get()) if self.var_xml.get() else None
            ifc_dir = Path(self.var_ifc_dir.get()) if self.var_ifc_dir.get() else None
            out_xml = Path(self.var_out.get()) if self.var_out.get() else None

            self._log(f"{EMOJI['ifc']} Папка IFC: {ifc_dir} (рекурсивно={bool(self.var_recursive_ifc.get())})")
            if not ifc_dir or not ifc_dir.exists() or not ifc_dir.is_dir():
                self._log(f"{EMOJI['err']} [ОШИБКА] Папка с IFC не найдена/не папка.", "err")
                messagebox.showerror("Ошибка", "Папка с IFC не найдена/не папка.")
                return
            files_ifc = collect_ifc_files(ifc_dir, recursive=bool(self.var_recursive_ifc.get()))
            self._log(f"    Найдено IFC: {len(files_ifc)}")
            if not files_ifc:
                self._log(f"{EMOJI['err']} [ОШИБКА] В папке отсутствуют файлы *.ifc", "err")
                messagebox.showerror("Ошибка", "В папке не найдено файлов *.ifc")
                return

            self.progress.start(12)

            # --- XML report ---
            if check_xml:
                if not xml or not xml.exists():
                    self._log(f"{EMOJI['err']} [ОШИБКА] XML не указан или не найден.", "err")
                elif not out_xml:
                    self._log(f"{EMOJI['err']} [ОШИБКА] Путь для XML-отчёта не указан.", "err")
                else:
                    if out_xml.exists() and not self._ask_overwrite(out_xml):
                        self._log(f"{EMOJI['report']} [ОТМЕНЕНО] Перезапись XML-отчёта отменена.")
                    else:
                        self._log(f"{EMOJI['xml']} Чтение XML...")
                        rules = read_rules(Path(__file__).with_name("rules.yaml"))
                        xml_map = extract_from_xml(xml, rules, case_sensitive=True)
                        self._log(f"    Записей IFC в XML: {len(xml_map)}")

                        self._log(f"{EMOJI['search']} Сверка по XML...")
                        rows_xml = build_report(xml_map, files_ifc, case_sensitive=True)
                        for r in rows_xml:
                            status = r.get("Статус","")
                            name = r.get("Имя файла")
                            if status == "OK":
                                self._log(f"{EMOJI['ok']} OK(XML) — {name}", "ok")
                            else:
                                self._log(f"{EMOJI['err']} {status}(XML) — {name} | {r.get('Подробности','')}", "err")

                        self._log(f"{EMOJI['xlsx']} Формирование XLSX (XML)...")
                        exit_xml, stats_xml = write_xlsx(rows_xml, out_xml)
                        self._log(f"{EMOJI['stats']} [ИТОГИ XML] Всего: {stats_xml.get('total')} | OK: {stats_xml.get('ok')} | Ошибок: {stats_xml.get('errors')}")
                        if self.var_open_after.get():
                            self._open_path(out_xml)

            # --- IUL report ---
            if check_iul:
                pdfs = list(self.iul_files)
                if self.var_iul_dir.get():
                    pdfs_dir = collect_pdf_files(Path(self.var_iul_dir.get()), recursive=bool(self.var_recursive_pdf.get()))
                    pdfs.extend(pdfs_dir)
                seen = set(); uniq = []
                for p in pdfs:
                    s = str(Path(p).resolve())
                    if s not in seen:
                        seen.add(s); uniq.append(Path(p))
                pdfs = uniq

                if not pdfs:
                    self._log(f"{EMOJI['warn']} ИУЛ-проверка включена, но PDF не выбраны/не найдены.", "warn")
                else:
                    out_iul = (out_xml.with_name(out_xml.stem.replace('.xlsx','') + "_iul.xlsx")
                               if (out_xml and out_xml.suffix.lower()==".xlsx")
                               else (Path.cwd() / "ifc_crc_report_iul.xlsx"))
                    if out_iul.exists() and not self._ask_overwrite(out_iul):
                        self._log(f"{EMOJI['report']} [ОТМЕНЕНО] Перезапись IUL-отчёта отменена.")
                    else:
                        if PdfReader is None:
                            self._log(f"{EMOJI['err']} [ОШИБКА] Для чтения ИУЛ (PDF) требуется PyPDF2. Установите зависимости.", "err")
                        else:
                            if self.var_tesseract_path.get():
                                set_tesseract_cmd(self.var_tesseract_path.get())
                            if self.var_ocr_lang.get():
                                set_ocr_langs([l for l in self.var_ocr_lang.get().split("+") if l])
                            self._log(f"{EMOJI['iul']} Чтение ИУЛ (PDF)...")
                            iul_map = extract_iul_entries(pdfs)
                            self._log(f"    Извлечено записей из ИУЛ: {len(iul_map)}")

                            self._log(f"{EMOJI['search']} Сверка по ИУЛ... (правило имени PDF: {'строгое' if self.var_pdf_name_strict.get() else 'мягкое'})")
                            rows_iul = build_report_iul(iul_map, files_ifc, strict_pdf_name=bool(self.var_pdf_name_strict.get()))
                            for r in rows_iul:
                                status = r.get("Статус","")
                                name = r.get("Имя файла")
                                if status == "OK":
                                    self._log(f"{EMOJI['ok']} OK(IUL) — {name}", "ok")
                                else:
                                    self._log(f"{EMOJI['err']} {status}(IUL) — {name} | {r.get('Подробности','')}", "err")

                            self._log(f"{EMOJI['xlsx']} Формирование XLSX (IUL)...")
                            exit_iul, stats_iul = write_xlsx_iul(rows_iul, out_iul)
                            self._log(f"{EMOJI['stats']} [ИТОГИ IUL] Всего: {stats_iul.get('total')} | OK: {stats_iul.get('ok')} | Ошибок: {stats_iul.get('errors')}")
                            if self.var_open_after.get():
                                self._open_path(out_iul)

            self._log(f"{EMOJI['done']} Готово.", "ok")

        except Exception as e:
            self._log(f"{EMOJI['err']} [КРИТИЧЕСКАЯ ОШИБКА] {e}", "err")
            messagebox.showerror("Критическая ошибка", str(e))
        finally:
            self.progress.stop()

if __name__ == "__main__":
    App().mainloop()
