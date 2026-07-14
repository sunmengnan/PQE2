#!/usr/bin/env python3
"""Polished Tkinter desktop GUI for the FAI/IPQC Excel extractor."""

from __future__ import annotations

import queue
import re
import os
import threading
import traceback
from datetime import date, datetime, time
from pathlib import Path
from tkinter import END, EXTENDED, BooleanVar, Listbox, StringVar, Tk, filedialog, messagebox, ttk
import tkinter.font as tkfont
from typing import Any, Dict, List, Optional, Sequence, Tuple

from fai_excel_extractor import default_output_path, export_workbook, extract_workbook


APP_TITLE = "FAI / IPQC Excel Extractor"
BASE_COLUMNS = ["Date ", "Sampling process", "Sampling time", "Sampling line#/Machine#", "FAI"]
PREVIEW_LIMIT = 300


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_date_filter(text: str) -> Optional[date]:
    text = clean_text(text)
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%y%m%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError("Date format must be 2026-06-01, 2026/06/01, or 260601.")


def parse_time_filter(text: str) -> Optional[time]:
    text = clean_text(text).replace("：", ":")
    if not text:
        return None
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            pass
    raise ValueError("Time format must be 08:00 or 8:00.")


def record_time(record: Dict[str, Any]) -> Optional[time]:
    value = record.get("Sampling time")
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, (int, float)):
        total = int(round((float(value) % 1) * 24 * 3600))
        return time((total // 3600) % 24, (total % 3600) // 60)
    try:
        return parse_time_filter(clean_text(value))
    except ValueError:
        return None


def display_value(value: Any) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, float) and 0 <= value < 1:
        total = int(round(value * 24 * 3600))
        return "%02d:%02d" % ((total // 3600) % 24, (total % 3600) // 60)
    return clean_text(value)


def sample_columns(records: Sequence[Dict[str, Any]]) -> List[str]:
    max_sample = 0
    for record in records:
        for key in record.keys():
            if key.startswith("Sample "):
                try:
                    max_sample = max(max_sample, int(key.split()[-1]))
                except ValueError:
                    pass
    return ["Sample %d" % idx for idx in range(1, max(10, max_sample) + 1)]


def natural_sort_key(value: Any) -> List[Any]:
    text = clean_text(value)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


class ExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x760")
        self.root.minsize(1020, 680)
        self.root.configure(bg="#EEF4FB")

        self.input_var = StringVar()
        self.output_var = StringVar()
        self.prefix_var = BooleanVar(value=False)
        self.source_sheet_var = BooleanVar(value=False)
        self.date_from_var = StringVar()
        self.date_to_var = StringVar()
        self.time_from_var = StringVar()
        self.time_to_var = StringVar()
        self.fai_filter_var = StringVar()
        self.status_var = StringVar(value="Select an Excel file to begin.")
        self.count_var = StringVar(value="Total 0 rows / Filtered 0 rows")

        self.messages = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.all_records: List[Dict[str, Any]] = []
        self.filtered_records: List[Dict[str, Any]] = []
        self.sheet_counts: Dict[str, int] = {}

        self._configure_style()
        self._build_ui()
        self.root.after(150, self._poll_messages)

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        families = set(tkfont.families(self.root))
        preferred_fonts = ["Aptos", "Inter", "Segoe UI", "Ubuntu", "DejaVu Sans", "Arial"]
        ui_font = next((font for font in preferred_fonts if font in families), "TkDefaultFont")
        self.ui_font = ui_font
        style.configure("App.TFrame", background="#EEF4FB")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure("Title.TLabel", background="#EEF4FB", foreground="#153B5C", font=(ui_font, 24, "bold"))
        style.configure("SubTitle.TLabel", background="#EEF4FB", foreground="#5D7488", font=(ui_font, 11))
        style.configure("CardTitle.TLabel", background="#FFFFFF", foreground="#153B5C", font=(ui_font, 13, "bold"))
        style.configure("Body.TLabel", background="#FFFFFF", foreground="#29465B", font=(ui_font, 10))
        style.configure("Status.TLabel", background="#FFFFFF", foreground="#0D7A5F", font=(ui_font, 10, "bold"))
        style.configure("TEntry", padding=6, font=(ui_font, 10))
        style.configure("TCheckbutton", background="#FFFFFF", foreground="#29465B", font=(ui_font, 10))
        style.configure("Accent.TButton", background="#1F6FEB", foreground="#FFFFFF", font=(ui_font, 11, "bold"), padding=(18, 9))
        style.map("Accent.TButton", background=[("active", "#1557B0"), ("disabled", "#9DB7D6")])
        style.configure("Soft.TButton", background="#E8F1FF", foreground="#1F4E79", font=(ui_font, 10), padding=(12, 7))
        style.configure("Treeview", font=(ui_font, 10), rowheight=28, background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=(ui_font, 10, "bold"), background="#DDEBFA", foreground="#153B5C")

    def _card(self, parent: ttk.Frame, row: int, column: int, **grid) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame", padding=16)
        frame.grid(row=row, column=column, sticky="nsew", padx=8, pady=8, **grid)
        return frame

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(outer, text="Extract dimension sampling data from every sheet, then filter and export by date, time, and FAI.", style="SubTitle.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 12))

        top = ttk.Frame(outer, style="App.TFrame")
        top.grid(row=2, column=0, sticky="ew")
        top.columnconfigure(0, weight=3)
        top.columnconfigure(1, weight=2)

        file_card = self._card(top, 0, 0)
        file_card.columnconfigure(1, weight=1)
        ttk.Label(file_card, text="File Settings", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        ttk.Label(file_card, text="Input Excel", style="Body.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
        ttk.Entry(file_card, textvariable=self.input_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(file_card, text="Browse", style="Soft.TButton", command=self._browse_input).grid(row=1, column=2, padx=(8, 0), pady=6)
        ttk.Label(file_card, text="Output Excel", style="Body.TLabel").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=6)
        ttk.Entry(file_card, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(file_card, text="Save As", style="Soft.TButton", command=self._browse_output).grid(row=2, column=2, padx=(8, 0), pady=6)
        ttk.Checkbutton(file_card, text="Add FAI prefix (4 -> FAI4)", variable=self.prefix_var).grid(row=3, column=1, sticky="w", pady=(8, 0))
        ttk.Checkbutton(file_card, text="Include Source sheet column", variable=self.source_sheet_var).grid(row=3, column=2, sticky="w", pady=(8, 0))

        filter_card = self._card(top, 0, 1)
        for idx in range(4):
            filter_card.columnconfigure(idx, weight=1)
        ttk.Label(filter_card, text="Filters", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        ttk.Label(filter_card, text="Date From", style="Body.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(filter_card, textvariable=self.date_from_var, width=13).grid(row=2, column=0, sticky="ew", padx=(0, 8), pady=4)
        ttk.Label(filter_card, text="Date To", style="Body.TLabel").grid(row=1, column=1, sticky="w")
        ttk.Entry(filter_card, textvariable=self.date_to_var, width=13).grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=4)
        ttk.Label(filter_card, text="Time From", style="Body.TLabel").grid(row=1, column=2, sticky="w")
        ttk.Entry(filter_card, textvariable=self.time_from_var, width=10).grid(row=2, column=2, sticky="ew", padx=(0, 8), pady=4)
        ttk.Label(filter_card, text="Time To", style="Body.TLabel").grid(row=1, column=3, sticky="w")
        ttk.Entry(filter_card, textvariable=self.time_to_var, width=10).grid(row=2, column=3, sticky="ew", pady=4)
        ttk.Label(filter_card, text="Dimension / FAI Contains", style="Body.TLabel").grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Entry(filter_card, textvariable=self.fai_filter_var).grid(row=4, column=0, columnspan=2, sticky="ew", padx=(0, 8), pady=4)
        ttk.Button(filter_card, text="Apply", style="Soft.TButton", command=self._apply_filters_from_ui).grid(row=4, column=2, sticky="ew", padx=(0, 8), pady=4)
        ttk.Button(filter_card, text="Clear", style="Soft.TButton", command=self._clear_filters).grid(row=4, column=3, sticky="ew", pady=4)
        ttk.Label(filter_card, text="FAI Multi-select", style="Body.TLabel").grid(row=5, column=0, columnspan=4, sticky="w", pady=(8, 2))
        fai_list_frame = ttk.Frame(filter_card, style="Card.TFrame")
        fai_list_frame.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(0, 4))
        fai_list_frame.columnconfigure(0, weight=1)
        self.fai_listbox = Listbox(fai_list_frame, height=4, selectmode=EXTENDED, exportselection=False, borderwidth=1, relief="solid", font=(self.ui_font, 10))
        fai_scroll = ttk.Scrollbar(fai_list_frame, orient="vertical", command=self.fai_listbox.yview)
        self.fai_listbox.configure(yscrollcommand=fai_scroll.set, bg="#FFFFFF", fg="#153B5C", selectbackground="#1F6FEB", selectforeground="#FFFFFF")
        self.fai_listbox.grid(row=0, column=0, sticky="ew")
        fai_scroll.grid(row=0, column=1, sticky="ns")
        ttk.Label(filter_card, text="Use Ctrl/Shift-click for multiple FAI values. Text contains supports comma-separated terms.", style="Body.TLabel").grid(row=7, column=0, columnspan=4, sticky="w", pady=(4, 0))
        ttk.Label(filter_card, text="Date examples: 2026-06-01 / 260601; time example: 08:00", style="Body.TLabel").grid(row=8, column=0, columnspan=4, sticky="w", pady=(2, 0))

        content = ttk.Frame(outer, style="App.TFrame")
        content.grid(row=3, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)
        preview_card = self._card(content, 0, 0)
        preview_card.columnconfigure(0, weight=1)
        preview_card.rowconfigure(2, weight=1)

        header = ttk.Frame(preview_card, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(1, weight=1)
        ttk.Label(header, text="Data Preview", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.count_var, style="Status.TLabel").grid(row=0, column=1, sticky="w", padx=(14, 0))
        self.extract_button = ttk.Button(header, text="Extract & Export", style="Accent.TButton", command=self._start_extract)
        self.extract_button.grid(row=0, column=2, padx=(8, 0))
        ttk.Button(header, text="Export Filtered", style="Soft.TButton", command=self._export_filtered).grid(row=0, column=3, padx=(8, 0))

        ttk.Label(preview_card, textvariable=self.status_var, style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 8))
        tree_frame = ttk.Frame(preview_card, style="Card.TFrame")
        tree_frame.grid(row=2, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_frame, show="headings")
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(title="Select Input Excel", filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        self.input_var.set(path)
        if not self.output_var.get().strip():
            self.output_var.set(str(default_output_path(Path(path))))

    def _browse_output(self) -> None:
        initial = self.output_var.get().strip()
        path = filedialog.asksaveasfilename(title="Save Extracted Result", initialfile=Path(initial).name if initial else "extracted.xlsx", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.output_var.set(path)

    def _validated_output_path(self) -> Path:
        input_path = Path(self.input_var.get().strip()).expanduser().resolve()
        output_text = self.output_var.get().strip()
        output_path = Path(output_text).expanduser().resolve() if output_text else default_output_path(input_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        self.output_var.set(str(output_path))
        return output_path

    def _start_extract(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        input_text = self.input_var.get().strip()
        if not input_text:
            messagebox.showwarning(APP_TITLE, "Please select an input Excel file first.")
            return
        input_path = Path(input_text).expanduser().resolve()
        if not input_path.exists():
            messagebox.showerror(APP_TITLE, "Input file does not exist:\n%s" % input_path)
            return
        try:
            output_path = self._validated_output_path()
            filters = self._read_filters()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        self.extract_button.configure(state="disabled")
        self.status_var.set("Reading Excel, please wait...")
        self.worker = threading.Thread(target=self._extract_worker, args=(input_path, output_path, filters, self.prefix_var.get(), self.source_sheet_var.get()), daemon=True)
        self.worker.start()

    def _extract_worker(self, input_path: Path, output_path: Path, filters: Tuple[Optional[date], Optional[date], Optional[time], Optional[time], List[str], List[str]], prefix_fai: bool, include_source_sheet: bool) -> None:
        try:
            records, sheet_counts = extract_workbook(input_path, prefix_fai=prefix_fai)
            if not records:
                self.messages.put(("error", "No valid measurement data was found."))
                return
            filtered = self._filter_records(records, filters)
            if not filtered:
                self.messages.put(("error", "No exportable data matches the current filters."))
                return
            export_workbook(filtered, output_path, include_source_sheet=include_source_sheet)
            self.messages.put(("success", (output_path, records, filtered, sheet_counts)))
        except Exception as exc:  # noqa: BLE001
            self.messages.put(("exception", (exc, traceback.format_exc())))

    def _selected_fai_values(self) -> List[str]:
        if not hasattr(self, "fai_listbox"):
            return []
        return [self.fai_listbox.get(index) for index in self.fai_listbox.curselection()]

    def _read_filters(self) -> Tuple[Optional[date], Optional[date], Optional[time], Optional[time], List[str], List[str]]:
        start_date = parse_date_filter(self.date_from_var.get())
        end_date = parse_date_filter(self.date_to_var.get())
        start_time = parse_time_filter(self.time_from_var.get())
        end_time = parse_time_filter(self.time_to_var.get())
        fai_text = clean_text(self.fai_filter_var.get())
        fai_terms = [term.strip().lower() for term in re.split(r"[,;|]", fai_text) if term.strip()]
        selected_fais = [clean_text(value).lower() for value in self._selected_fai_values()]
        if start_date and end_date and start_date > end_date:
            raise ValueError("Start date cannot be later than end date.")
        if start_time and end_time and start_time > end_time:
            raise ValueError("Start time cannot be later than end time.")
        return start_date, end_date, start_time, end_time, fai_terms, selected_fais

    def _filter_records(self, records: Sequence[Dict[str, Any]], filters: Tuple[Optional[date], Optional[date], Optional[time], Optional[time], List[str], List[str]]) -> List[Dict[str, Any]]:
        start_date, end_date, start_time, end_time, fai_terms, selected_fais = filters
        selected_set = set(selected_fais)
        result: List[Dict[str, Any]] = []
        for record in records:
            rec_date = record.get("Date ")
            if isinstance(rec_date, datetime):
                rec_date = rec_date.date()
            if start_date and isinstance(rec_date, date) and rec_date < start_date:
                continue
            if end_date and isinstance(rec_date, date) and rec_date > end_date:
                continue
            rec_time = record_time(record)
            if start_time and rec_time and rec_time < start_time:
                continue
            if end_time and rec_time and rec_time > end_time:
                continue
            fai_value = clean_text(record.get("FAI")).lower()
            if fai_terms or selected_set:
                manual_match = any(term in fai_value for term in fai_terms)
                selected_match = fai_value in selected_set
                if not manual_match and not selected_match:
                    continue
            result.append(record)
        return result

    def _apply_filters_from_ui(self) -> None:
        if not self.all_records:
            messagebox.showinfo(APP_TITLE, "Please extract data first.")
            return
        try:
            filters = self._read_filters()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return
        self.filtered_records = self._filter_records(self.all_records, filters)
        self._refresh_preview()
        self.status_var.set("Filters applied. Click Export Filtered to save the result.")

    def _clear_filters(self) -> None:
        self.date_from_var.set("")
        self.date_to_var.set("")
        self.time_from_var.set("")
        self.time_to_var.set("")
        self.fai_filter_var.set("")
        if hasattr(self, "fai_listbox"):
            self.fai_listbox.selection_clear(0, END)
        if self.all_records:
            self.filtered_records = list(self.all_records)
            self._refresh_preview()

    def _populate_fai_list(self, records: Sequence[Dict[str, Any]]) -> None:
        if not hasattr(self, "fai_listbox"):
            return
        selected_before = set(self._selected_fai_values())
        values = sorted({clean_text(record.get("FAI")) for record in records if clean_text(record.get("FAI"))}, key=natural_sort_key)
        self.fai_listbox.delete(0, END)
        for value in values:
            self.fai_listbox.insert(END, value)
        for index, value in enumerate(values):
            if value in selected_before:
                self.fai_listbox.selection_set(index)

    def _export_filtered(self) -> None:
        if not self.filtered_records:
            messagebox.showinfo(APP_TITLE, "No filtered data is available to export.")
            return
        try:
            output_path = self._validated_output_path()
            export_workbook(self.filtered_records, output_path, include_source_sheet=self.source_sheet_var.get())
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, "Export failed:\n%s" % exc)
            return
        messagebox.showinfo(APP_TITLE, "Export completed.\n\nOutput file:\n%s\n\nRows: %s" % (output_path, len(self.filtered_records)))

    def _poll_messages(self) -> None:
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                self.extract_button.configure(state="normal")
                if kind == "success":
                    output_path, records, filtered, sheet_counts = payload
                    self.all_records = list(records)
                    self.filtered_records = list(filtered)
                    self.sheet_counts = dict(sheet_counts)
                    self._populate_fai_list(self.all_records)
                    self._refresh_preview()
                    self.status_var.set("Exported: %s" % output_path)
                    messagebox.showinfo(APP_TITLE, "Extraction and export completed.\n\nOutput file:\n%s\n\nTotal measurement rows: %s\nFiltered export rows: %s" % (output_path, len(records), len(filtered)))
                elif kind == "error":
                    self.status_var.set("Processing failed.")
                    messagebox.showwarning(APP_TITLE, str(payload))
                elif kind == "exception":
                    exc, trace = payload
                    self.status_var.set("Processing failed.")
                    print(trace)
                    messagebox.showerror(APP_TITLE, "Processing failed:\n%s" % exc)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_messages)

    def _refresh_preview(self) -> None:
        records = self.filtered_records
        columns = BASE_COLUMNS + sample_columns(records)
        if self.source_sheet_var.get():
            columns.append("Source sheet")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150 if col in BASE_COLUMNS else 105, anchor="center", stretch=True)
        for idx, record in enumerate(records[:PREVIEW_LIMIT]):
            values = []
            for col in columns:
                key = "source_sheet" if col == "Source sheet" else col
                values.append(display_value(record.get(key, "")))
            self.tree.insert("", "end", values=values, tags=("odd" if idx % 2 else "even",))
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F7FBFF")
        note = " (showing first %s rows)" % PREVIEW_LIMIT if len(records) > PREVIEW_LIMIT else ""
        self.count_var.set("Total %s rows / Filtered %s rows%s" % (len(self.all_records), len(records), note))


def main() -> None:
    if os.environ.get("FAI_EXTRACTOR_SMOKE_TEST") == "1":
        import openpyxl  # noqa: F401
        import tkinter  # noqa: F401
        import fai_excel_extractor  # noqa: F401
        return
    root = Tk()
    ExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
