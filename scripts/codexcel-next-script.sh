#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: $0 <workbook.xlsx|workbook.xlsm>"
}

if [[ $# -ne 1 ]]; then
  usage
  exit 1
fi

workbook="$1"
if [[ ! "$workbook" =~ \.(xlsx|xlsm)$ ]]; then
  echo "Error: workbook must end with .xlsx or .xlsm"
  exit 1
fi
if [[ ! -f "$workbook" ]]; then
  echo "Error: workbook not found: $workbook"
  exit 1
fi

base_name="$(basename "$workbook")"
stem="${base_name%.*}"
target_dir="codexcel/$stem"

if [[ ! -d "$target_dir" ]]; then
  echo "Error: missing $target_dir. Run scripts/codexcel-setup.sh $workbook first."
  exit 1
fi

max_version=-1
shopt -s nullglob
for file_path in "$target_dir"/"$stem"[0-9]*.py; do
  file_name="$(basename "$file_path")"
  suffix="${file_name#$stem}"
  suffix="${suffix%.py}"
  if [[ "$suffix" =~ ^[0-9]+$ ]] && (( suffix > max_version )); then
    max_version=$suffix
  fi
done
shopt -u nullglob

next_version=$((max_version + 1))
target_script="$target_dir/${stem}${next_version}.py"

cat > "$target_script" <<EOF
from openpyxl import load_workbook

WORKBOOK_PATH = "${workbook}"


def apply() -> None:
    wb = load_workbook(WORKBOOK_PATH, keep_vba=WORKBOOK_PATH.lower().endswith(".xlsm"))
    ws = wb.active
    # Replace with explicit sheet and cell updates.
    # Prefer formulas, e.g. ws["F2"] = "=IF(E2=0,\\"\\",D2/E2)"
    wb.save(WORKBOOK_PATH)


if __name__ == "__main__":
    apply()
EOF

echo "Created: $target_script"
