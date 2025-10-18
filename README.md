# CPUMemorySysgaugeData_converter

CPU/Memory TXT → Excel (CPUMemory) — GUI + CLI

Convert process CPU/Memory text exports (tabs or 2+ spaces) into a formatted Excel workbook with a filterable table, typed number formats, and calculated columns.

## Features
- Parse text with columns like `Process Name`, `Instances`, `CPU`, `Memory`, `Threads`, `Handles`, `Data`, `Status`
- Accepts tab-delimited or multi‑space separated data; tolerates numbers with commas and duplicate header lines
- Exports an Excel `.xlsx` workbook with:
  - A striped, filterable Excel Table and frozen header row
  - Autosized columns and typed number formats (ints, floats, percentages)
  - Calculated columns per row:
    - Cumulative Total CPU
    - Total Usage of  CPU in 100% (B / C *100) CPU
    - TOTAL PERCENTAGE CPU (left blank by design)
    - Cumulative Total Memory
    - Total Usage of  Memory in 100% (B / C *100)
    - TOTAL PERCENTAGE MEMORY (left blank by design)
- Multi‑file mode: one sheet per input file (sanitized/unique sheet names)
- No pandas required; uses `openpyxl` for Excel writing

## Requirements
- Python 3.8+
- `openpyxl` (for Excel output)
- Optional: `tkinter` for the GUI (falls back to CLI if unavailable)

Install dependency:

```bash
pip install -U openpyxl
```

## Quick start

GUI (if `tkinter` is available):

```bash
python con_cpu_memory.py
```

Force CLI or when `tkinter` is not available:

```bash
python con_cpu_memory.py --cli -i cpu_mem.txt -o report.xlsx
```

Multiple inputs (each file becomes its own sheet):

```bash
python con_cpu_memory.py --cli -i cpu_mem1.txt cpu_mem2.txt -o report.xlsx
```

Run self‑tests:

```bash
python con_cpu_memory.py --run-tests
```

## CLI options (subset)
- `--cli`: force command‑line mode (disable GUI)
- `-i, --input`: one or more input `.txt` files
- `-o, --output`: output Excel path (e.g., `report.xlsx`)

Note: In multi‑file mode, each input file is written to its own sheet with a sanitized unique name. In single‑file mode, the sheet name defaults to the input filename (sanitized).

## Input format
Headers similar to:

```
Process Name	Instances	CPU	Memory	Threads	Handles	Data	Status
System	1	0.06	0.13	332	5,250	0.00	Normal
smss	1	0.00	0.62	2	64	0.00	Normal
... (more rows)
```

Rules and tolerances:
- Delimiters: tabs or 2+ spaces
- Numbers may contain commas
- `Status` column is optional
- Duplicate header lines anywhere in the file are ignored

## Output details
- File type: `.xlsx`
- Excel Table style: `TableStyleMedium9`
- Header row is frozen; numeric columns are formatted appropriately
- Calculated columns use workbook formulas so totals and percentages reflect the data

## Troubleshooting
- "openpyxl is required": install with `pip install openpyxl`
- No GUI appears: your Python may not include `tkinter`; use `--cli`
- Empty output: ensure your input has the expected columns and is tab/multi‑space separated

---

Made with Python and `openpyxl`. See `con_cpu_memory.py` for implementation details and additional helpers.
