#!/usr/bin/env python3
"""Extract IPQC/FAI sampling rows from inspection Excel workbooks.

The reader follows the same idea as the 2177PQE scripts: it reads workbook XML
directly so the result does not depend on Excel formula cache recalculation.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

OUTPUT_BASE_COLUMNS = ["Date ", "Sampling process", "Sampling time", "Sampling line#/Machine#", "FAI"]


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch.upper()) - 64
    return n


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_label(value: Any) -> str:
    return clean_text(value).replace(" ", "").replace("\u3000", "").replace("：", ":")


def parse_cell_range(ref: str) -> Tuple[int, int, int, int]:
    if ":" not in ref:
        match = re.match(r"([A-Z]+)(\d+)$", ref)
        if not match:
            raise ValueError("Bad cell reference: %s" % ref)
        col = col_to_num(match.group(1))
        row = int(match.group(2))
        return row, col, row, col
    start, end = ref.split(":", 1)
    s_match = re.match(r"([A-Z]+)(\d+)$", start)
    e_match = re.match(r"([A-Z]+)(\d+)$", end)
    if not s_match or not e_match:
        raise ValueError("Bad cell range: %s" % ref)
    return int(s_match.group(2)), col_to_num(s_match.group(1)), int(e_match.group(2)), col_to_num(e_match.group(1))


class ExcelXmlReader:
    """Small OOXML reader for values, sheet names and merged-cell ranges."""

    def __init__(self, path: Path):
        self.path = path
        self.zip = ZipFile(path)
        self.shared_strings = self._read_shared_strings()
        self.sheet_targets = self._read_sheet_targets()

    def close(self) -> None:
        self.zip.close()

    def __enter__(self) -> "ExcelXmlReader":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def _read_shared_strings(self) -> List[str]:
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        root = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        strings: List[str] = []
        for si in root.findall(NS_MAIN + "si"):
            strings.append("".join(t.text or "" for t in si.iter(NS_MAIN + "t")))
        return strings

    def _read_sheet_targets(self) -> Dict[str, str]:
        workbook = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        targets: Dict[str, str] = {}
        sheets = workbook.find(NS_MAIN + "sheets")
        if sheets is None:
            return targets
        for sheet in sheets:
            name = sheet.attrib["name"]
            rid = sheet.attrib[NS_REL + "id"]
            target = rid_to_target[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            targets[name] = target
        return targets

    def sheet_names(self) -> List[str]:
        return list(self.sheet_targets.keys())

    def cell_value(self, cell: ET.Element) -> Any:
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(NS_MAIN + "t"))
        if cell_type == "str":
            v = cell.find(NS_MAIN + "v")
            return "" if v is None else (v.text or "")
        v = cell.find(NS_MAIN + "v")
        if v is None:
            return ""
        text = v.text or ""
        if cell_type == "s":
            try:
                return self.shared_strings[int(text)]
            except (ValueError, IndexError):
                return text
        if cell_type == "b":
            return text == "1"
        try:
            return float(text) if any(ch in text for ch in ".Ee") else int(text)
        except ValueError:
            return text

    def read_sheet(self, sheet_name: str) -> Tuple[Dict[int, Dict[int, Any]], List[Tuple[int, int, int, int]]]:
        target = self.sheet_targets[sheet_name]
        root = ET.fromstring(self.zip.read(target))
        rows: Dict[int, Dict[int, Any]] = {}
        sheet_data = root.find(NS_MAIN + "sheetData")
        if sheet_data is not None:
            for row in sheet_data.findall(NS_MAIN + "row"):
                row_num = int(row.attrib.get("r", "0"))
                values: Dict[int, Any] = {}
                for cell in row.findall(NS_MAIN + "c"):
                    ref = cell.attrib.get("r", "")
                    match = re.match(r"([A-Z]+)", ref)
                    if not match:
                        continue
                    values[col_to_num(match.group(1))] = self.cell_value(cell)
                if any(clean_text(v) != "" for v in values.values()):
                    rows[row_num] = values
        merges: List[Tuple[int, int, int, int]] = []
        merge_cells = root.find(NS_MAIN + "mergeCells")
        if merge_cells is not None:
            for merge in merge_cells.findall(NS_MAIN + "mergeCell"):
                ref = merge.attrib.get("ref")
                if ref:
                    try:
                        merges.append(parse_cell_range(ref))
                    except ValueError:
                        pass
        return rows, merges


@dataclass(frozen=True)
class SamplingGroup:
    columns: Tuple[int, ...]
    sample_date: date
    sample_time: time


def merged_lookup(merges: Sequence[Tuple[int, int, int, int]], rows_of_interest: Iterable[int]) -> Dict[Tuple[int, int], Tuple[int, int]]:
    wanted = set(rows_of_interest)
    lookup: Dict[Tuple[int, int], Tuple[int, int]] = {}
    for min_row, min_col, max_row, max_col in merges:
        hit_rows = [r for r in wanted if min_row <= r <= max_row]
        for row in hit_rows:
            for col in range(min_col, max_col + 1):
                lookup[(row, col)] = (min_row, min_col)
    return lookup


def value_at(rows: Dict[int, Dict[int, Any]], row: int, col: int, merge_map: Dict[Tuple[int, int], Tuple[int, int]]) -> Any:
    if row in rows and col in rows[row]:
        direct_value = rows[row][col]
        if clean_text(direct_value) != "":
            return direct_value
    top_left = merge_map.get((row, col))
    if top_left is None:
        return rows.get(row, {}).get(col, "")
    top_row, top_col = top_left
    return rows.get(top_row, {}).get(top_col, "")


def parse_yymmdd(value: Any) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        text = "%06d" % int(value)
    else:
        text = re.sub(r"\D", "", clean_text(value))
        if len(text) == 8 and text.startswith("20"):
            return date(int(text[:4]), int(text[4:6]), int(text[6:8]))
    if len(text) != 6:
        return None
    yy, mm, dd = int(text[:2]), int(text[2:4]), int(text[4:6])
    year = 2000 + yy if yy < 70 else 1900 + yy
    try:
        return date(year, mm, dd)
    except ValueError:
        return None


def parse_hhmm(value: Any) -> Optional[time]:
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, timedelta):
        total_seconds = int(value.total_seconds()) % (24 * 3600)
        return time(total_seconds // 3600, (total_seconds % 3600) // 60)
    if isinstance(value, (int, float)):
        if value < 0:
            return None
        total_seconds = int(round((float(value) % 1) * 24 * 3600))
        return time((total_seconds // 3600) % 24, (total_seconds % 3600) // 60)
    text = clean_text(value).replace("：", ":")
    match = re.search(r"(\d{1,2})\s*:\s*(\d{1,2})", text)
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    if hour == 24:
        hour = 0
    if 0 <= hour <= 23 and 0 <= minute <= 59:
        return time(hour, minute)
    return None


def excel_time_fraction(t: time) -> float:
    return (t.hour * 3600 + t.minute * 60 + t.second) / 86400.0


def display_fai(value: Any, prefix_fai: bool = False) -> Any:
    text = clean_text(value)
    if prefix_fai and text and not text.upper().startswith("FAI"):
        return "FAI" + text
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def contains_chinese(value: Any) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", clean_text(value)))


def is_rejected_sample(value: Any) -> bool:
    text = clean_text(value)
    normalized = text.replace(" ", "").replace("\u3000", "").upper()
    return normalized in ("OK", "NG", "/", "\\", "-", "--", "NA", "N/A", "NULL", "NONE") or contains_chinese(value)


def is_meaningful_sample(value: Any) -> bool:
    """Return True only for real measurement data.

    Template result rows such as appearance OK/NG rows, judgment rows and empty
    placeholder rows should not be exported as Sample data.
    """
    text = clean_text(value)
    if text == "":
        return False
    normalized = text.replace(" ", "").replace("\u3000", "").upper()
    if is_rejected_sample(value):
        return False
    return True


def find_row_containing(rows: Dict[int, Dict[int, Any]], keywords: Sequence[str]) -> Optional[int]:
    for row_num in sorted(rows):
        for value in rows[row_num].values():
            label = normalize_label(value)
            if all(keyword in label for keyword in keywords):
                return row_num
    return None


def find_label_row(rows: Dict[int, Dict[int, Any]], labels: Sequence[str]) -> Optional[int]:
    for row_num in sorted(rows):
        for value in rows[row_num].values():
            label = normalize_label(value)
            if any(expected in label for expected in labels):
                return row_num
    return None


def find_header_value(rows: Dict[int, Dict[int, Any]], merge_map: Dict[Tuple[int, int], Tuple[int, int]], labels: Sequence[str], max_row: int = 8) -> str:
    max_col = max((col for row in rows.values() for col in row.keys()), default=0)
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            if any(label in normalize_label(value_at(rows, row, col, merge_map)) for label in labels):
                for next_col in range(col + 1, min(max_col, col + 4) + 1):
                    value = clean_text(value_at(rows, row, next_col, merge_map))
                    if value:
                        return value
    return ""


def find_dimension_header(rows: Dict[int, Dict[int, Any]]) -> Optional[Tuple[int, int]]:
    for row_num in sorted(rows):
        for col, value in rows[row_num].items():
            if "尺寸序号" in normalize_label(value):
                return row_num, col
    return None


def build_sampling_groups(
    rows: Dict[int, Dict[int, Any]],
    date_row: int,
    time_row: int,
    merge_map: Dict[Tuple[int, int], Tuple[int, int]],
) -> List[SamplingGroup]:
    max_col = max((col for row in rows.values() for col in row.keys()), default=0)
    groups: List[SamplingGroup] = []
    current_cols: List[int] = []
    current_key: Optional[Tuple[date, time]] = None
    for col in range(1, max_col + 1):
        sample_date = parse_yymmdd(value_at(rows, date_row, col, merge_map))
        sample_time = parse_hhmm(value_at(rows, time_row, col, merge_map))
        key = (sample_date, sample_time) if sample_date is not None and sample_time is not None else None
        if key is not None and key == current_key:
            current_cols.append(col)
            continue
        if current_key is not None and current_cols:
            groups.append(SamplingGroup(tuple(current_cols), current_key[0], current_key[1]))
        current_cols = [col] if key is not None else []
        current_key = key
    if current_key is not None and current_cols:
        groups.append(SamplingGroup(tuple(current_cols), current_key[0], current_key[1]))
    return groups


def extract_sheet_records(sheet_name: str, rows: Dict[int, Dict[int, Any]], merges: Sequence[Tuple[int, int, int, int]], prefix_fai: bool = False) -> List[Dict[str, Any]]:
    dim_header = find_dimension_header(rows)
    date_row = find_row_containing(rows, ["检验", "日期"]) or find_label_row(rows, ["出炉日期", "日期(yymmdd)", "日期"])
    time_row = find_row_containing(rows, ["检验", "时间"]) or find_label_row(rows, ["出炉时间", "时间(hh:mm)", "时间"])
    if dim_header is None or date_row is None or time_row is None:
        return []

    header_row, fai_col = dim_header
    rows_of_interest = set([4, date_row, time_row])
    rows_of_interest.update(r for r in rows.keys() if r >= header_row)
    merge_map = merged_lookup(merges, rows_of_interest)
    groups = build_sampling_groups(rows, date_row, time_row, merge_map)
    if not groups:
        return []

    process = ""
    if "工序" in normalize_label(value_at(rows, 4, 7, merge_map)):
        process = clean_text(value_at(rows, 4, 8, merge_map))
    if not process:
        process = find_header_value(rows, merge_map, ["工序"])
    machine = ""
    if "机台" in normalize_label(value_at(rows, 4, 9, merge_map)):
        machine = clean_text(value_at(rows, 4, 10, merge_map))
    if not machine:
        machine = find_header_value(rows, merge_map, ["机台", "机台号"])

    records: List[Dict[str, Any]] = []
    for row_num in sorted(r for r in rows.keys() if r > header_row):
        raw_fai = value_at(rows, row_num, fai_col, merge_map)
        if clean_text(raw_fai) == "":
            continue
        if normalize_label(raw_fai) in ("尺寸序号", "判定", "检验员"):
            continue
        for group in groups:
            raw_samples = [value_at(rows, row_num, col, merge_map) for col in group.columns]
            filled_samples = [sample for sample in raw_samples if clean_text(sample) != ""]
            if not filled_samples or any(is_rejected_sample(sample) for sample in filled_samples):
                continue
            samples = [sample for sample in filled_samples if is_meaningful_sample(sample)]
            if not samples:
                continue
            record: Dict[str, Any] = {
                "source_sheet": sheet_name,
                "Date ": group.sample_date,
                "Sampling process": process,
                "Sampling time": excel_time_fraction(group.sample_time),
                "Sampling line#/Machine#": machine,
                "FAI": display_fai(raw_fai, prefix_fai=prefix_fai),
            }
            for idx, sample in enumerate(samples, start=1):
                record["Sample %d" % idx] = sample
            records.append(record)
    return records


def extract_workbook(path: Path, prefix_fai: bool = False) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
    records: List[Dict[str, Any]] = []
    sheet_counts: Dict[str, int] = {}
    with ExcelXmlReader(path) as reader:
        for sheet_name in reader.sheet_names():
            rows, merges = reader.read_sheet(sheet_name)
            sheet_records = extract_sheet_records(sheet_name, rows, merges, prefix_fai=prefix_fai)
            if sheet_records:
                records.extend(sheet_records)
                sheet_counts[sheet_name] = len(sheet_records)
    return records, sheet_counts


def export_workbook(records: Sequence[Dict[str, Any]], output_path: Path, include_source_sheet: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "IPQC template"

    max_samples = 0
    for record in records:
        for key in record.keys():
            match = re.match(r"Sample (\d+)$", key)
            if match:
                max_samples = max(max_samples, int(match.group(1)))
    columns = list(OUTPUT_BASE_COLUMNS)
    columns.extend("Sample %d" % idx for idx in range(1, max(10, max_samples) + 1))
    if include_source_sheet:
        columns.append("Source sheet")

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(name="Arial", bold=True)
    body_font = Font(name="Arial")
    center = Alignment(horizontal="center", vertical="center")

    for col_idx, header in enumerate(columns, start=1):
        cell = ws.cell(1, col_idx, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    for row_idx, record in enumerate(records, start=2):
        for col_idx, header in enumerate(columns, start=1):
            key = "source_sheet" if header == "Source sheet" else header
            cell = ws.cell(row_idx, col_idx, record.get(key, ""))
            cell.font = body_font
            cell.alignment = center
            if header == "Date ":
                cell.number_format = "yyyy-mm-dd"
            elif header == "Sampling time":
                cell.number_format = "hh:mm"

    for col_idx, header in enumerate(columns, start=1):
        values = [clean_text(ws.cell(row, col_idx).value) for row in range(1, min(ws.max_row, 200) + 1)]
        width = max([len(header)] + [len(v) for v in values if v]) + 2
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 12), 32)
    ws.freeze_panes = "A2"
    wb.save(output_path)


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(input_path.stem + "_extracted.xlsx")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract all inspection sheet FAI sampling data to a new Excel file.")
    parser.add_argument("input", type=Path, help="Input .xlsx/.xlsm inspection workbook")
    parser.add_argument("-o", "--output", type=Path, help="Output .xlsx path")
    parser.add_argument("--prefix-fai", action="store_true", help="Prefix FAI values with 'FAI' (for example 4 -> FAI4)")
    parser.add_argument("--include-source-sheet", action="store_true", help="Append a Source sheet column for traceability")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    input_path = args.input.expanduser().resolve()
    output_path = (args.output or default_output_path(input_path)).expanduser().resolve()
    if not input_path.exists():
        print("Input file does not exist: %s" % input_path, file=sys.stderr)
        return 2
    try:
        records, sheet_counts = extract_workbook(input_path, prefix_fai=args.prefix_fai)
    except BadZipFile:
        print("Input is not a valid .xlsx/.xlsm workbook: %s" % input_path, file=sys.stderr)
        return 2
    if not records:
        print("No matching inspection records found.", file=sys.stderr)
        return 1
    export_workbook(records, output_path, include_source_sheet=args.include_source_sheet)
    print("Output: %s" % output_path)
    print("Records: %d" % len(records))
    for sheet_name, count in sheet_counts.items():
        print("  %s: %d" % (sheet_name, count))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())