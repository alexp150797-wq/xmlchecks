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
from pkg.report_builder_pdf_xml import build_report_pdf_xml
from pkg.iul_reader import extract_iul_entries, PdfReader  # type: ignore
from pkg.report_builder_iul import build_report_iul
from pkg.xlsx_writer_combined import write_combined_xlsx
import socket
import threading

APP_TITLE = "IFC CRC Checker — GUI"
# slightly wider for better layout
APP_MIN_W, APP_MIN_H = 1160, 800

EMOJI = {
    "xml": "🧾",
    "ifc": "🏗️",
    "iul": "📄",
    "pdf": "📄",
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
        self.resizable(False, False)
        self._apply_theme()

        # store critical error messages for optional log saving
        self.error_messages: list[str] = []

        # Paths & options
        self.var_xml = tk.StringVar()
        self.ifc_files: list[Path] = []
        self.var_ifc_label = tk.StringVar(value="(не выбрано)")
        self.var_ifc_dir = tk.StringVar()
        self.var_out = tk.StringVar(value=str(Path.cwd() / "ifc_crc_report.xlsx"))
        self.var_recursive_ifc = tk.BooleanVar(value=True)
        self.var_open_after = tk.BooleanVar(value=True)

        # Checks: independently
        self.var_check_xml = tk.BooleanVar(value=True)
        self.var_check_iul = tk.BooleanVar(value=True)
        self.var_check_pdf_xml = tk.BooleanVar(value=False)

        # IUL selection
        self.iul_files: list[Path] = []
        self.var_iul_label = tk.StringVar(value="(не выбрано)")
        self.var_iul_dir = tk.StringVar()
        self.var_recursive_pdf = tk.BooleanVar(value=True)
        self.var_pdf_name_strict = tk.BooleanVar(value=False)  # строгое имя PDF (_УЛ)

        # Generic PDF selection for PDF↔XML
        self.pdf_files: list[Path] = []
        self.var_pdf_label = tk.StringVar(value="(не выбрано)")
        self.var_pdf_dir = tk.StringVar()
        self.var_recursive_pdf_other = tk.BooleanVar(value=True)

        self._instr_panel = None

        self._build_ui()

    def _apply_theme(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        bg = "#f5f5f5"
        self.configure(background=bg)
        self.style.configure("TFrame", background=bg)
        self.style.configure("TLabel", background=bg, padding=2)
        self.style.configure("TCheckbutton", background=bg)
        self.style.configure("TButton", padding=8)
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), background="#4a90e2", foreground="white", padding=4)
        self.style.configure("Small.TLabel", foreground="#333", background="#4a90e2")
        self.style.configure("Accent.TButton", padding=10, background="#4a90e2", foreground="white")
        self.style.map("Accent.TButton", background=[("active", "#3572a5")])

    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        header = tk.Frame(self, bg="#4a90e2")
        header.pack(fill="x")
        ttk.Label(header, text="Сверка: XML↔IFC и ИУЛ(PDF)↔IFC → XLSX", style="Header.TLabel").pack(side="left", **pad)
        ttk.Label(header, text="Зелёный — успех, красный — ошибка. Имена сопоставляются с учётом регистра.", style="Small.TLabel").pack(side="right", **pad)
        ttk.Button(header, text="Инструкция", command=self._show_instruction).pack(side="right", **pad)

        body = ttk.Frame(self); body.pack(fill="both", expand=True, padx=10, pady=6)

        # What to check (split)
        box_ifc = ttk.LabelFrame(body, text="Проверки к IFC")
        box_ifc.grid(row=0, column=0, columnspan=2, sticky="we", **pad)
        ttk.Checkbutton(box_ifc, text="XML ↔ IFC", variable=self.var_check_xml).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(box_ifc, text="ИУЛ (PDF) ↔ IFC", variable=self.var_check_iul).grid(row=0, column=1, sticky="w", padx=8, pady=4)

        box_pdf = ttk.LabelFrame(body, text="Проверки к PDF")
        box_pdf.grid(row=0, column=2, columnspan=2, sticky="we", **pad)
        ttk.Checkbutton(box_pdf, text="PDF ↔ XML", variable=self.var_check_pdf_xml).grid(row=0, column=0, sticky="w", padx=8, pady=4)

        # XML
        ttk.Label(body, text=f"{EMOJI['xml']} XML:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(body, textvariable=self.var_xml).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)
        ttk.Button(body, text="Файл...", command=self._choose_xml).grid(row=1, column=3, **pad)

        # IFC selection
        ifc_frame = ttk.LabelFrame(body, text="IFC")
        ifc_frame.grid(row=2, column=0, columnspan=2, sticky="we", padx=10, pady=6)
        ttk.Label(ifc_frame, text=f"{EMOJI['ifc']} Выбор:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Button(ifc_frame, text="IFC-файлы...", command=self._choose_ifc_files).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(ifc_frame, text="Папка с IFC...", command=self._choose_ifc_dir).grid(row=0, column=2, padx=6, pady=4)
        ttk.Label(ifc_frame, textvariable=self.var_ifc_label).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(ifc_frame, text="Рекурсивно по IFC", variable=self.var_recursive_ifc).grid(row=0, column=4, sticky="w", padx=6)
        ttk.Button(ifc_frame, text="Очистить выбор IFC", command=self._clear_ifc).grid(row=1, column=2, padx=6, pady=4)

        # IUL PDFs
        iul_frame = ttk.LabelFrame(body, text="ИУЛ (PDF)")
        iul_frame.grid(row=3, column=0, columnspan=2, sticky="we", padx=10, pady=6)
        ttk.Label(iul_frame, text=f"{EMOJI['iul']} Выбор:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Button(iul_frame, text="PDF-файлы...", command=self._choose_iul_files).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(iul_frame, text="Папка с PDF...", command=self._choose_iul_dir).grid(row=0, column=2, padx=6, pady=4)
        ttk.Label(iul_frame, textvariable=self.var_iul_label).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(iul_frame, text="Рекурсивно", variable=self.var_recursive_pdf).grid(row=1, column=0, sticky="w", padx=8)
        ttk.Checkbutton(iul_frame, text="Строгое имя PDF (имяIFC_УЛ.pdf)", variable=self.var_pdf_name_strict).grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Button(iul_frame, text="Очистить выбор PDF", command=self._clear_iul).grid(row=1, column=2, padx=6, pady=4)

        # Generic PDFs for PDF↔XML
        pdf_frame = ttk.LabelFrame(body, text="PDF")
        pdf_frame.grid(row=2, column=2, columnspan=2, rowspan=2, sticky="we", padx=10, pady=6)
        ttk.Label(pdf_frame, text=f"{EMOJI['pdf']} Выбор:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Button(pdf_frame, text="PDF-файлы...", command=self._choose_pdf_files).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(pdf_frame, text="Папка с PDF...", command=self._choose_pdf_dir).grid(row=0, column=2, padx=6, pady=4)
        ttk.Label(pdf_frame, textvariable=self.var_pdf_label).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Button(pdf_frame, text="Очистить выбор PDF", command=self._clear_pdf).grid(row=1, column=2, padx=6, pady=4)
        ttk.Checkbutton(pdf_frame, text="Рекурсивно", variable=self.var_recursive_pdf_other).grid(row=1, column=0, sticky="w", padx=8)

        # Out
        ttk.Label(body, text=f"{EMOJI['xlsx']} Отчёт (XLSX):").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(body, textvariable=self.var_out).grid(row=4, column=1, columnspan=2, sticky="ew", **pad)
        ttk.Button(body, text="Сохранить как...", command=self._choose_out).grid(row=4, column=3, **pad)

        # Run
        btns = ttk.Frame(body); btns.grid(row=5, column=0, columnspan=4, sticky="w", **pad)
        ttk.Button(btns, text=f"{EMOJI['search']} Сформировать отчёт(ы)", style="Accent.TButton", command=self._run).pack(side="left", padx=6)
        ttk.Checkbutton(btns, text="Открыть отчёты по завершению", variable=self.var_open_after).pack(side="left", padx=6)
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
        body.columnconfigure(2, weight=1)
        body.rowconfigure(8, weight=1)

    def _show_instruction(self):
        """Показывает инструкцию в выезжающей панели."""
        if self._instr_panel is not None:
            self._instr_panel.destroy()
            self._instr_panel = None
            return
        instr_path = Path(__file__).with_name("INSTRUCTION.md")
        try:
            text = instr_path.read_text(encoding="utf-8")
        except Exception:
            text = "Файл инструкции не найден."
        panel_width = 400
        panel = tk.Frame(self, width=panel_width, height=self.winfo_height())
        panel.place(x=self.winfo_width(), y=0, relheight=1)
        txt = tk.Text(panel, wrap="word")
        scroll = ttk.Scrollbar(panel, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        txt.insert("1.0", text)
        txt.config(state="disabled")
        txt.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._instr_panel = panel

        def slide():
            x = panel.winfo_x()
            target = self.winfo_width() - panel_width
            if x > target:
                panel.place(x=max(x-20, target), y=0, relheight=1)
                panel.after(10, slide)
        slide()

    def _choose_xml(self):
        p = filedialog.askopenfilename(title="Выберите XML", filetypes=[("XML","*.xml"),("Все файлы","*.*")])
        if p:
            self.var_xml.set(p)
            if not self.var_out.get():
                self.var_out.set(str(Path(p).with_name("ifc_crc_report.xlsx")))

    def _choose_ifc_files(self):
        ps = filedialog.askopenfilenames(title="Выберите IFC", filetypes=[("IFC","*.ifc")])
        if ps:
            self.ifc_files = [Path(p) for p in ps]
            self._update_ifc_label()
        else:
            self._update_ifc_label()

    def _choose_ifc_dir(self):
        p = filedialog.askdirectory(title="Выберите папку с IFC")
        if p:
            self.var_ifc_dir.set(p)
            self._update_ifc_label()

    def _clear_ifc(self):
        self.ifc_files = []
        self.var_ifc_dir.set("")
        self._update_ifc_label()

    def _update_ifc_label(self):
        parts = []
        if self.ifc_files:
            parts.append(f"файлов: {len(self.ifc_files)}")
        if self.var_ifc_dir.get():
            parts.append(f"папка: {self.var_ifc_dir.get()}")
        self.var_ifc_label.set(", ".join(parts) if parts else "(не выбрано)")

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

    def _choose_pdf_files(self):
        ps = filedialog.askopenfilenames(title="Выберите PDF", filetypes=[("PDF","*.pdf")])
        if ps:
            self.pdf_files = [Path(p) for p in ps]
            self._update_pdf_label()
        else:
            self._update_pdf_label()

    def _choose_pdf_dir(self):
        p = filedialog.askdirectory(title="Выберите папку с PDF")
        if p:
            self.var_pdf_dir.set(p)
            self._update_pdf_label()

    def _clear_pdf(self):
        self.pdf_files = []
        self.var_pdf_dir.set("")
        self._update_pdf_label()

    def _update_pdf_label(self):
        parts = []
        if self.pdf_files:
            parts.append(f"файлов: {len(self.pdf_files)}")
        if self.var_pdf_dir.get():
            parts.append(f"папка: {self.var_pdf_dir.get()}")
        self.var_pdf_label.set(", ".join(parts) if parts else "(не выбрано)")

    def _choose_out(self):
        p = filedialog.asksaveasfilename(
            title="Сохранить отчёт как...",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook","*.xlsx")],
            initialfile="ifc_crc_report.xlsx"
        )
        if p:
            self.var_out.set(p)

    def _log(self, msg, kind="info", critical: bool = False):
        tag = "info" if kind not in ("ok", "err", "warn") else kind
        if critical:
            self.error_messages.append(msg)
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end"); self.update()

    def _save_error_log(self):
        p = filedialog.asksaveasfilename(
            title="Сохранить журнал ошибок",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All files", "*.*")],
        )
        if p:
            try:
                Path(p).write_text("\n".join(self.error_messages), encoding="utf-8")
                self._log(f"{EMOJI['path']} Журнал сохранён: {p}")
            except Exception as e:
                self._log(f"{EMOJI['err']} Не удалось сохранить журнал: {e}", "err")

    def _error_dialog(self, message: str):
        if messagebox.askyesno("Ошибка", f"{message}\n\nСохранить журнал ошибок?"):
            self._save_error_log()

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
            self.error_messages = []

            check_xml = bool(self.var_check_xml.get())
            check_iul = bool(self.var_check_iul.get())
            check_pdf_xml = bool(self.var_check_pdf_xml.get())
            if not (check_xml or check_iul or check_pdf_xml):
                self._log(f"{EMOJI['warn']} Ничего не выбрано для проверки.", "warn")
                return
            xml = Path(self.var_xml.get()) if self.var_xml.get() else None
            out_path = Path(self.var_out.get()) if self.var_out.get() else None
            if out_path is None:
                msg = "Путь для отчёта не указан."
                self._log(f"{EMOJI['err']} [ОШИБКА] {msg}", "err", critical=True)
                self._error_dialog(msg)
                return
            if out_path.exists() and not self._ask_overwrite(out_path):
                self._log(f"{EMOJI['report']} [ОТМЕНЕНО] Перезапись отчёта отменена.")
                return

            files_ifc: list[Path] = []
            if check_xml or check_iul:
                files_ifc = list(self.ifc_files)
                if self.var_ifc_dir.get():
                    files_ifc.extend(collect_ifc_files(Path(self.var_ifc_dir.get()), recursive=bool(self.var_recursive_ifc.get())))
                files_ifc = sorted({p.resolve() for p in files_ifc})
                self._log(f"{EMOJI['ifc']} Выбрано IFC: {len(files_ifc)}")
                if not files_ifc:
                    msg = "IFC не выбраны/не найдены."
                    self._log(f"{EMOJI['err']} [ОШИБКА] {msg}", "err", critical=True)
                    self._error_dialog("Не выбраны файлы IFC")
                    return

            iul_pdfs: list[Path] = []
            if check_iul:
                iul_pdfs = list(self.iul_files)
                if self.var_iul_dir.get():
                    iul_pdfs.extend(collect_pdf_files(Path(self.var_iul_dir.get()), recursive=bool(self.var_recursive_pdf.get())))
                iul_pdfs = sorted({p.resolve() for p in iul_pdfs})
                self._log(f"{EMOJI['iul']} Выбрано ИУЛ PDF: {len(iul_pdfs)}")

            pdfs: list[Path] = []
            if check_pdf_xml:
                pdfs = list(self.pdf_files)
                if self.var_pdf_dir.get():
                    pdfs.extend(collect_pdf_files(Path(self.var_pdf_dir.get()), recursive=bool(self.var_recursive_pdf_other.get())))
                pdfs = sorted({p.resolve() for p in pdfs})
                self._log(f"{EMOJI['pdf']} Выбрано PDF: {len(pdfs)}")

            rows_xml: list[dict] = []
            rows_iul: list[dict] = []
            rows_pdf: list[dict] = []

            self.progress.start(12)
            self.update()

            if check_xml:
                if not xml or not xml.exists():
                    self._log(f"{EMOJI['err']} [ОШИБКА] XML не указан или не найден.", "err")
                else:
                    self._log(f"{EMOJI['xml']} Чтение XML...")
                    rules = read_rules(Path(__file__).with_name("rules.yaml"))
                    xml_map, xml_pdf = extract_from_xml(
                        xml, rules, case_sensitive=True, include_sign_files=True
                    )
                    self._log(f"    Записей IFC в XML: {len(xml_map)} | Подписей PDF: {len(xml_pdf)}")
                    self._log(f"{EMOJI['search']} Сверка по XML...")
                    rows_xml = build_report(xml_map, files_ifc, case_sensitive=True)
                    for r in rows_xml:
                        status = r.get("Статус","")
                        name = r.get("Имя файла IFC")
                        if status == "OK":
                            self._log(f"{EMOJI['ok']} OK(XML) — {name}", "ok")
                        else:
                            self._log(f"{EMOJI['err']} {status}(XML) — {name} | {r.get('Подробности','')}", "err")

            if check_pdf_xml:
                if not xml or not xml.exists():
                    self._log(f"{EMOJI['err']} [ОШИБКА] XML не указан или не найден для PDF↔XML.", "err")
                elif not pdfs:
                    self._log(f"{EMOJI['warn']} PDF↔XML включена, но PDF не выбраны/не найдены.", "warn")
                else:
                    self._log(f"{EMOJI['xml']} Чтение XML (PDF)...")
                    rules_pdf = read_rules(Path(__file__).with_name("rules.yaml"))
                    rules_pdf["filter_format"] = "PDF"
                    xml_pdf_map = extract_from_xml(xml, rules_pdf, case_sensitive=True)
                    self._log(f"    Записей PDF в XML: {len(xml_pdf_map)}")
                    self._log(f"{EMOJI['search']} Сверка PDF↔XML...")
                    rows_pdf = build_report_pdf_xml(xml_pdf_map, pdfs, case_sensitive=True)
                    for r in rows_pdf:
                        status = r.get("Статус","")
                        name = r.get("Имя файла IFC")
                        if status == "OK":
                            self._log(f"{EMOJI['ok']} OK(PDF/XML) — {name}", "ok")
                        else:
                            self._log(f"{EMOJI['err']} {status}(PDF/XML) — {name} | {r.get('Подробности','')}", "err")

            if check_iul:
                if not iul_pdfs:
                    self._log(f"{EMOJI['warn']} ИУЛ-проверка включена, но PDF не выбраны/не найдены.", "warn")
                else:
                    if PdfReader is None:
                        self._log(f"{EMOJI['err']} [ОШИБКА] Для чтения ИУЛ (PDF) требуется PyPDF2. Установите зависимости.", "err")
                    else:
                        self._log(f"{EMOJI['iul']} Чтение ИУЛ (PDF)...")
                        iul_map = extract_iul_entries(
                            iul_pdfs,
                            progress=lambda e: self._log(f"    {e.basename} ← {e.source_pdf}"),
                        )
                        self._log(f"    Извлечено записей из ИУЛ: {len(iul_map)}")
                        self._log(
                            f"{EMOJI['search']} Сверка по ИУЛ... (правило имени PDF: {'строгое' if self.var_pdf_name_strict.get() else 'мягкое'})"
                        )
                        rows_iul = build_report_iul(
                            iul_map,
                            files_ifc,
                            iul_pdfs,
                            strict_pdf_name=bool(self.var_pdf_name_strict.get()),
                            include_pdf_name_col=bool(self.var_pdf_name_strict.get()),
                        )
                        for r in rows_iul:
                            status = r.get("Статус","")
                            name = r.get("Имя файла IFC")
                            if status == "OK":
                                self._log(f"{EMOJI['ok']} OK(IUL) — {name}", "ok")
                            else:
                                self._log(f"{EMOJI['err']} {status}(IUL) — {name} | {r.get('Подробности','')}", "err")

            while True:
                try:
                    stats = write_combined_xlsx(
                        rows_xml if check_xml else None,
                        rows_iul if check_iul else None,
                        rows_pdf if check_pdf_xml else None,
                        out_path,
                        include_pdf_name_col=bool(self.var_pdf_name_strict.get()),
                    )
                    break
                except PermissionError:
                    msg = f"Не удалось записать отчёт (возможно открыт): {out_path}"
                    self._log(f"{EMOJI['err']} [ОШИБКА] {msg}", "err", critical=True)
                    if not messagebox.askretrycancel(
                        "Файл занят",
                        f"Не удалось записать файл:\n{out_path}\nЗакройте файл и повторите.",
                    ):
                        self._log(
                            f"{EMOJI['report']} [ОТМЕНЕНО] Запись отчёта отменена.",
                            "warn",
                        )
                        self._error_dialog(msg)
                        return

            if check_xml and stats.get("xml"):
                s = stats["xml"]
                self._log(f"{EMOJI['stats']} [ИТОГИ XML] Всего: {s.get('total')} | OK: {s.get('ok')} | Ошибок: {s.get('errors')}")
            if check_iul and stats.get("iul"):
                s = stats["iul"]
                self._log(f"{EMOJI['stats']} [ИТОГИ IUL] Всего: {s.get('total')} | OK: {s.get('ok')} | Ошибок: {s.get('errors')}")
            if check_pdf_xml and stats.get("pdf_xml"):
                s = stats["pdf_xml"]
                self._log(f"{EMOJI['stats']} [ИТОГИ PDF/XML] Всего: {s.get('total')} | OK: {s.get('ok')} | Ошибок: {s.get('errors')}")

            if self.var_open_after.get():
                self._open_path(out_path)

            self._log(f"{EMOJI['done']} Готово.", "ok")

        except Exception as e:
            msg = f"[КРИТИЧЕСКАЯ ОШИБКА] {e}"
            self._log(f"{EMOJI['err']} {msg}", "err", critical=True)
            self._error_dialog(str(e))
        finally:
            self.progress.stop()

PORT = 65432


def _acquire_instance(app: App) -> bool:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", PORT))
    except OSError:
        try:
            with socket.create_connection(("127.0.0.1", PORT), timeout=1) as s:
                s.send(b"show")
        except OSError:
            pass
        return False
    sock.listen(1)

    def server():
        while True:
            conn, _ = sock.accept()
            conn.close()
            app.after(0, lambda: (app.deiconify(), app.lift(), app.focus_force()))

    threading.Thread(target=server, daemon=True).start()
    return True


if __name__ == "__main__":
    app = App()
    if _acquire_instance(app):
        app.mainloop()
