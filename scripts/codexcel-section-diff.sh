#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: $0 <workbook.xlsx|workbook.xlsm> <edit-script.py> [xleak args...]"
  echo "Example: $0 report.xlsx codexcel/report/report2.py --sheet \"Summary\""
}

if [[ $# -lt 2 ]]; then
  usage
  exit 1
fi

workbook="$1"
edit_script="$2"
shift 2

if [[ ! -f "$workbook" ]]; then
  echo "Error: workbook not found: $workbook"
  exit 1
fi

if [[ ! -f "$edit_script" ]]; then
  echo "Error: edit script not found: $edit_script"
  exit 1
fi

before_file="$(mktemp)"
after_file="$(mktemp)"
trap 'rm -f "$before_file" "$after_file"' EXIT

xleak "$workbook" "$@" > "$before_file"
python "$edit_script"
xleak "$workbook" "$@" > "$after_file"

git --no-pager diff --no-index --word-diff=color -- "$before_file" "$after_file" || true
