# openpyxl Editing Reference

Use this guide when writing `codexcel/<filename>/<filename>X.py` edit scripts.

Source: official openpyxl docs (stable).

## Install and imports

```bash
python -m pip install openpyxl
```

```python
from openpyxl import Workbook, load_workbook
```

## Create and save a workbook

```python
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws["A1"] = "Hello"
wb.save("example.xlsx")
```

Important: saving uses the same extension as the filename (`.xlsx`, `.xlsm`, etc.).

## Load existing workbook

```python
wb = load_workbook("report.xlsx")
```

For macro-enabled files, preserve VBA:

```python
wb = load_workbook("report.xlsm", keep_vba=True)
```

Useful load options from docs:

- `data_only=True`: read cached values instead of formulas.
- `read_only=True`: optimize for large read workflows.

## Select worksheets

```python
ws = wb.active
ws = wb["Summary"]
sheet_names = wb.sheetnames
```

## Read and write cells

```python
value = ws["B2"].value
ws["C2"] = value
ws.cell(row=5, column=2, value=42)
```

Iterate rows/columns:

```python
for row in ws.iter_rows(min_row=2, max_row=10, min_col=1, max_col=5):
    for cell in row:
        _ = cell.value
```

## Write formulas (Excel-side logic)

Prefer formulas over Python-computed constants so sheets stay dynamic.

```python
ws["F2"] = "=IF(E2=0,\"\",D2/E2)"
ws["G2"] = "=SUM(B2:E2)"
ws["H2"] = "=IFERROR(VLOOKUP(A2,Lookup!A:B,2,FALSE),\"\")"
ws["I2"] = "='Revenue 2026'!B2"
```

openpyxl stores formulas; Excel calculates them when opened/recalculated.

## Save changes

```python
wb.save("report.xlsx")
```

The save operation overwrites the target file.

## Known caveats from docs

- Keep extensions and content aligned (`.xlsm` with `keep_vba=True` when needed).
- Shapes and some drawing objects may be lost when files are opened/saved.

## Script pattern

```python
from openpyxl import load_workbook

WORKBOOK_PATH = "report.xlsx"

def apply() -> None:
    wb = load_workbook(WORKBOOK_PATH, keep_vba=WORKBOOK_PATH.lower().endswith(".xlsm"))
    ws = wb["Summary"]
    ws["F2"] = "=IF(E2=0,\"\",D2/E2)"
    wb.save(WORKBOOK_PATH)

if __name__ == "__main__":
    apply()
```

## Official docs links

- https://openpyxl.readthedocs.io/en/stable/tutorial.html
- https://openpyxl.readthedocs.io/en/stable/simple_formulae.html
