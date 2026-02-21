# Codexcel

`codexcel` is a Codex skill for working with Excel files (`.xlsx`, `.xlsm`) using:
- `xleak` for viewing/sheet inspection
- versioned Python scripts (`openpyxl`) for edits
- before/after terminal diffs for change display

## Repository Layout

```text
.
├── SKILL.md
├── agents/openai.yaml
├── references/openpyxl.md
└── scripts/
    ├── codexcel-setup.sh
    ├── codexcel-next-script.sh
    ├── codexcel-run-edit.sh
    └── codexcel-section-diff.sh
```

## Requirements

- `python` 3.x
- `openpyxl` (`python -m pip install openpyxl`)
- `xleak`
- `git` (for `--no-index` colored diffs)

## Install Skill Into Codex

From this repo root:

```bash
mkdir -p "${CODEX_HOME:-$HOME/.codex}/skills/codexcel"
cp SKILL.md "${CODEX_HOME:-$HOME/.codex}/skills/codexcel/"
cp -r agents references scripts "${CODEX_HOME:-$HOME/.codex}/skills/codexcel/"
chmod 755 "${CODEX_HOME:-$HOME/.codex}/skills/codexcel/scripts"/codexcel-*.sh
```

Restart Codex after install/update.

## Quick Usage

1) Prepare workspace for a workbook:

```bash
scripts/codexcel-setup.sh report.xlsx --create
```

2) Inspect workbook/sheets:

```bash
xleak report.xlsx
xleak report.xlsx --sheet "Summary"
```

3) Create the next versioned edit script:

```bash
scripts/codexcel-next-script.sh report.xlsx
```

4) Edit generated `codexcel/report/reportX.py`, then run:

```bash
scripts/codexcel-run-edit.sh codexcel/report/reportX.py
```

5) Show before/after diff for a section:

```bash
scripts/codexcel-section-diff.sh report.xlsx codexcel/report/reportX.py --sheet "Summary"
```

## Long Cell Text Fallback

If `xleak` truncates long text, print a full cell value via Python:

```bash
python -c 'from openpyxl import load_workbook; path="report.xlsx"; wb=load_workbook(path, keep_vba=path.lower().endswith(".xlsm")); ws=wb["Summary"]; print(ws["B12"].value)'
```
