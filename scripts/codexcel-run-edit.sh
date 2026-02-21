#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Usage: $0 <path/to/edit-script.py>"
  exit 1
fi

edit_script="$1"
if [[ ! -f "$edit_script" ]]; then
  echo "Error: script not found: $edit_script"
  exit 1
fi

python "$edit_script"
echo "Applied: $edit_script"
