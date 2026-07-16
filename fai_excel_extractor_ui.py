#!/usr/bin/env python3
"""Streamlit UI for FAI/IPQC Excel sampling extraction."""

from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import streamlit as st

from fai_excel_extractor import default_multi_output_path, export_workbook, extract_workbooks


DEFAULT_DIR = Path(__file__).resolve().parent


def save_uploaded_file(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp.write(uploaded_file.getbuffer())
    temp.flush()
    temp.close()
    return Path(temp.name)


def records_to_frame(records: List[dict], include_source_sheet: bool) -> pd.DataFrame:
    df = pd.DataFrame(records)
    if df.empty:
        return df
    source_count = df["source_file"].dropna().astype(str).nunique() if "source_file" in df.columns else 0
    include_source_file = include_source_sheet or source_count > 1
    if not include_source_file and "source_file" in df.columns:
        df = df.drop(columns=["source_file"])
    if not include_source_sheet and "source_sheet" in df.columns:
        df = df.drop(columns=["source_sheet"])
    df = df.rename(columns={"source_sheet": "Source sheet", "source_file": "Source file"})
    base = ["Date ", "Sampling process", "Sampling time", "Sampling line#/Machine#", "FAI"]
    samples = sorted([c for c in df.columns if c.startswith("Sample ")], key=lambda name: int(name.split()[-1]))
    extra = [c for c in df.columns if c not in base + samples + ["Source file", "Source sheet"]]
    ordered = base + samples + extra
    if include_source_file and "Source file" in df.columns:
        ordered.append("Source file")
    if include_source_sheet and "Source sheet" in df.columns:
        ordered.append("Source sheet")
    return df[[c for c in ordered if c in df.columns]]


def workbook_bytes(records: List[dict], include_source_sheet: bool) -> bytes:
    output = BytesIO()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp:
        temp_path = Path(temp.name)
    export_workbook(records, temp_path, include_source_sheet=include_source_sheet)
    data = temp_path.read_bytes()
    temp_path.unlink(missing_ok=True)
    output.write(data)
    return output.getvalue()


def load_local(path_text: str, prefix_fai: bool) -> Tuple[List[dict], dict, Path]:
    input_paths = [Path(part.strip()).expanduser().resolve() for part in path_text.split(";") if part.strip()]
    records, sheet_counts = extract_workbooks(input_paths, prefix_fai=prefix_fai)
    return records, sheet_counts, default_multi_output_path(input_paths)


def main() -> None:
    st.set_page_config(page_title="FAI Excel 提取工具", layout="wide")
    st.title("FAI / IPQC Excel 提取工具")
    st.caption("读取输入 Excel 的所有 Sheet，提取检验日期、时间、工序、机台、FAI 和 Sample 数据到新 Excel。")

    with st.sidebar:
        st.header("输入")
        mode = st.radio("文件来源", ["上传文件", "本地路径"], horizontal=True)
        prefix_fai = st.checkbox("FAI 值加前缀（4 -> FAI4）", value=False)
        include_source_sheet = st.checkbox("输出 Source file / Source sheet 追溯列", value=True)
        local_path = st.text_input("本地 Excel 路径（多个文件用 ; 分隔）", str(DEFAULT_DIR / "input.xlsx"), disabled=mode != "本地路径")
        uploaded_files = st.file_uploader("上传一个或多个 .xlsx/.xlsm", type=["xlsx", "xlsm"], accept_multiple_files=True, disabled=mode != "上传文件")
        run = st.button("开始提取", type="primary")

    if not run:
        st.info("请选择或上传 Excel 后点击开始提取。")
        return

    try:
        if mode == "上传文件":
            if not uploaded_files:
                st.warning("请先上传 Excel 文件。")
                return
            input_paths = []
            display_names = []
            for uploaded_file in uploaded_files:
                input_paths.append(save_uploaded_file(uploaded_file))
                display_names.append(uploaded_file.name)
            records, sheet_counts = extract_workbooks(input_paths, prefix_fai=prefix_fai)
            for record in records:
                temp_name = record.get("source_file")
                for temp_path, original_name in zip(input_paths, display_names):
                    if temp_path.name == temp_name:
                        record["source_file"] = original_name
                        break
            download_name = "FAI_IPQC_Excel_Extracted.xlsx" if len(uploaded_files) > 1 else Path(uploaded_files[0].name).stem + "_extracted.xlsx"
        else:
            records, sheet_counts, output_path = load_local(local_path, prefix_fai=prefix_fai)
            download_name = output_path.name
    except Exception as exc:  # noqa: BLE001 - UI should show friendly error
        st.error("提取失败：%s" % exc)
        return

    if not records:
        st.warning("没有找到符合模板的检验数据。")
        return

    st.success("提取完成，共 %d 行。" % len(records))
    summary = pd.DataFrame([{"Sheet": sheet, "Rows": count} for sheet, count in sheet_counts.items()])
    st.subheader("Sheet 统计")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.subheader("预览")
    df = records_to_frame(records, include_source_sheet=include_source_sheet)
    st.dataframe(df.head(500), use_container_width=True, height=520)

    st.download_button(
        "下载提取结果 Excel",
        workbook_bytes(records, include_source_sheet=include_source_sheet),
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()