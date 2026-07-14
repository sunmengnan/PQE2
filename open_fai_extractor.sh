#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
conda run -n yolov5 streamlit run fai_excel_extractor_ui.py