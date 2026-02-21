#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: $0 <workbook.xlsx|workbook.xlsm> [--create]"
}

if [[ $# -lt 1 || $# -gt 2 ]]; then
  usage
  exit 1
fi

workbook="$1"
create_flag="${2:-}"

if [[ ! "$workbook" =~ \.(xlsx|xlsm)$ ]]; then
  echo "Error: workbook must end with .xlsx or .xlsm"
  exit 1
fi

if [[ "$create_flag" == "--create" ]]; then
  python - "$workbook" <<'PY'
import os
import sys
from openpyxl import Workbook

path = sys.argv[1]
if not os.path.exists(path) or os.path.getsize(path) == 0:
    wb = Workbook()
    wb.save(path)
PY
fi

if [[ ! -f "$workbook" ]]; then
  echo "Error: workbook not found: $workbook"
  echo "Available .xlsx/.xlsm files:"
  find . -maxdepth 1 -type f \( -name "*.xlsx" -o -name "*.xlsm" \) -printf "%f\n" | sort
  exit 1
fi

base_name="$(basename "$workbook")"
stem="${base_name%.*}"
target_dir="codexcel/$stem"
target_script="$target_dir/${stem}0.py"

mkdir -p codexcel
mkdir -p "$target_dir"
touch "$target_script"

echo "Prepared: $target_script"
