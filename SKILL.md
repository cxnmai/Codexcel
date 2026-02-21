---
name: codexcel
description: View, edit, and display Excel (.xlsx/.xlsm) files by combining xleak inspection with versioned openpyxl scripts in a codexcel workspace. Use when creating new workbooks, preparing existing files for edits, applying formula-based spreadsheet updates, handling sheet-specific changes, or presenting before/after diff output for workbook modifications.
---

# Codexcel Workflow

Use three phases: viewing, editing, and displaying.

## Requirements

- Require `xleak` on PATH.
- Require `python` with `openpyxl`.
- Use `pandas` only when needed for null checks with `pd.notna()`.

## 1) Setup

Initialize workspace:

```bash
mkdir -p codexcel
```

Create a new workbook:

```bash
touch <filename>.xlsx
scripts/codexcel-setup.sh <filename>.xlsx --create
```

Use `touch` to reserve the name, then initialize a valid workbook with `--create`.

Add an existing workbook:

```bash
ls -1 *.xlsx *.xlsm 2>/dev/null
```

If the requested workbook is not found, stop and tell the user it cannot be found.

Prepare `codexcel/<filename>/` and initial version script:

```bash
scripts/codexcel-setup.sh <filename>.xlsx
# or for macro-enabled files:
scripts/codexcel-setup.sh <filename>.xlsm
```

If `codexcel/<filename>/<filename>0.py` does not exist, create it.

## 2) Viewing

List sheets and preview first sheet:

```bash
xleak <filename>.xlsx
```

View specific sheet:

```bash
xleak <filename>.xlsx --sheet "<sheetname>"
```

For scoped before/after output, capture section before editing:

```bash
xleak <filename>.xlsx --sheet "<sheetname>" > /tmp/codexcel-before.txt
```

Add extra xleak flags when a smaller section is needed.

## 3) Editing

Create a new versioned Python edit script for each change:

```bash
scripts/codexcel-next-script.sh <filename>.xlsx
```

This creates `codexcel/<filename>/<filename>X.py` where `X` increments (`0, 1, 2, ...`).

Write edits in that file using `openpyxl` patterns from `references/openpyxl.md`.

Run the edit:

```bash
scripts/codexcel-run-edit.sh codexcel/<filename>/<filename>X.py
```

Prefer fully rewriting the script file when implementing a new satisfied change set. Treat each script as disposable change logic.

## 4) Displaying Changes

Capture section after edit:

```bash
xleak <filename>.xlsx --sheet "<sheetname>" > /tmp/codexcel-after.txt
```

Show red/green diff:

```bash
git --no-pager diff --no-index --word-diff=color -- /tmp/codexcel-before.txt /tmp/codexcel-after.txt || true
```

Or run a combined before/edit/after flow:

```bash
scripts/codexcel-section-diff.sh <filename>.xlsx codexcel/<filename>/<filename>X.py --sheet "<sheetname>"
```

## Editing Rules

- Use Excel formulas instead of Python-calculated hardcoded values.
- Convert all logic to Excel-style formulas before writing to cells.
- Search all requested matches, not only first match.
- Guard division formulas against zero (avoid `#DIV/0!`).
- Verify all cell references to avoid `#REF!`.
- Use proper cross-sheet formula references (`Sheet1!A1`, `='My Sheet'!B2`).
- Handle null checks with `pd.notna()` when using pandas-based preprocessing.
