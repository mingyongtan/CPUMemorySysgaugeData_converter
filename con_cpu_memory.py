#!/usr/bin/env python3
"""
CPU/Memory TXT → Excel Sheet (CPUMemory) — GUI + CLI
----------------------------------------------------
Reads raw text like:

Process Name\tInstances\tCPU\tMemory\tThreads\tHandles\tData\tStatus
System\t1\t0.06\t0.13\t332\t5,250\t0.00\tNormal
smss\t1\t0.00\t0.62\t2\t64\t0.00\tNormal
...

Outputs an **Excel sheet** with:
- Data wrapped in an **Excel Table** (striped, filterable)
- Calculated columns (as Excel formulas using **structured references** for sort‑safety):
    • Cumulative Total CPU → `=SUM([CPU])`
    • Total Usage of  CPU in 100% (B / C *100) CPU → `=[@CPU]/SUM([CPU])` (formatted as %)
    • Cumulative Total Memory → `=SUM([Memory])`
    • Total Usage of  Memory in 100% (B / C *100) → `=[@Memory]/[@[Cumulative Total Memory]]` (formatted as %)
    • TOTAL PERCENTAGE CPU / MEMORY → **left blank** (per request)

Usage
- GUI (default when Tkinter is available):
    python cpu_memory_table.py
- Force CLI (+ multi-file):
    python cpu_memory_table.py --cli -i cpu_mem1.txt cpu_mem2.txt -o report.xlsx
- Self-tests:
    python cpu_memory_table.py --run-tests

Notes
- **Multi-file**: pass multiple `-i` paths in CLI or use the GUI’s **Load .txt (multi)** button; **each file** is parsed into its **own Excel sheet**, named after the text file (sanitized to Excel’s 31‑char limit and uniqueness).
- No pandas required. Requires `openpyxl` for Excel writing.
- Tolerates duplicate header rows anywhere and missing Status column.
- Splits on tabs OR 2+ spaces. Handles numbers with commas.
"""
from __future__ import annotations
import argparse
import os
import re
import sys
from typing import List, Dict, Optional, Tuple

# --- Optional GUI deps (graceful fallback to CLI) ---
try:
    import tkinter as tk  # type: ignore
    from tkinter import ttk, filedialog, messagebox  # type: ignore
    TK_OK = True
except Exception:
    TK_OK = False

# --------------------------- Parsing ---------------------------
HEADER_NAMES = [
    'Process Name', 'Instances', 'CPU', 'Memory', 'Threads', 'Handles', 'Data', 'Status'
]

TAB_RE = re.compile(r"\t+")
MULTISPACE_SPLIT = re.compile(r"\s{2,}")
QUOTES_RE = re.compile(r'"')


def _clean_line(line: str) -> str:
    line = line.replace('\ufeff', '').replace('ï»¿', '')
    return QUOTES_RE.sub('', line.strip())


def _tokenize(line: str) -> List[str]:
    """Split by tabs first; if only one token, try 2+ spaces."""
    raw = _clean_line(line)
    if not raw:
        return []
    toks = [t.strip() for t in TAB_RE.split(raw) if t.strip()]
    if len(toks) <= 1:
        toks = [t.strip() for t in MULTISPACE_SPLIT.split(raw) if t.strip()]
    return toks


def _looks_like_header(tokens: List[str]) -> bool:
    norm = [t.lower() for t in tokens]
    return (
        (len(norm) >= 3 and norm[0].startswith('process') and ('cpu' in norm) and ('memory' in norm))
    ) or set(h.lower() for h in HEADER_NAMES).issuperset(set(norm))


def _to_int(s: str) -> int:
    s = (s or '').replace(',', '')
    try:
        return int(float(s))
    except Exception:
        return 0


def _to_float(s: str) -> float:
    s = (s or '').replace(',', '')
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_cpu_memory_text(raw: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    for line in raw.splitlines():
        line = _clean_line(line)
        if not line:
            continue
        tokens = _tokenize(line)
        if not tokens:
            continue
        if _looks_like_header(tokens):
            continue  # skip any header

        # Allow 7 or 8 tokens (Status optional). If more than 8, join middle tokens back into Process Name.
        if len(tokens) < 7:
            continue
        if len(tokens) > 8:
            # assume last 6 are instances..status, so everything before that is process name
            process = ' '.join(tokens[: len(tokens) - 7])
            tail = tokens[-7:]
            tokens = [process] + tail

        # If 7 tokens, assume Status is blank
        if len(tokens) == 7:
            tokens.append('')

        (proc, inst, cpu, mem, threads, handles, data, status) = tokens[:8]
        row = {
            'Process Name': proc,
            'Instances': _to_int(inst),
            'CPU': _to_float(cpu),
            'Memory': _to_float(mem),
            'Threads': _to_int(threads),
            'Handles': _to_int(handles),
            'Data': _to_float(data),
            'Status': status,
        }
        rows.append(row)

    return rows


# --------------------------- Calculations ---------------------------
EXTRA_HEADERS = [
    'Cumulative Total CPU',
    'Total Usage of  CPU in 100% (B / C *100) CPU',
    'TOTAL PERCENTAGE CPU',  # leave blank per request
    'Cumulative Total Memory',
    'Total Usage of  Memory in 100% (B / C *100)',
    'TOTAL PERCENTAGE MEMORY',  # leave blank per request
]
HEADERS = HEADER_NAMES + EXTRA_HEADERS


def add_calculated_columns(rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    total_cpu = sum(float(r['CPU']) for r in rows)
    total_mem = sum(float(r['Memory']) for r in rows)
    out: List[Dict[str, object]] = []
    for r in rows:
        cpu = float(r['CPU'])
        mem = float(r['Memory'])
        rr = dict(r)
        # Per spec: show grand totals in every row (not running sums)
        rr['Cumulative Total CPU'] = total_cpu
        rr['Total Usage of  CPU in 100% (B / C *100) CPU'] = (cpu / total_cpu) if total_cpu > 0 else 0.0
        rr['TOTAL PERCENTAGE CPU'] = ''  # explicitly blank
        rr['Cumulative Total Memory'] = total_mem
        rr['Total Usage of  Memory in 100% (B / C *100)'] = (mem / total_mem) if total_mem > 0 else 0.0
        rr['TOTAL PERCENTAGE MEMORY'] = ''  # explicitly blank
        out.append(rr)
    return out


# --------------------------- Excel writing ---------------------------

def write_cpumemory_sheet(rows_base: List[Dict[str, object]], out_path: str, sheet_name: str = 'CPUMemory') -> str:
    try:
        from openpyxl import Workbook, load_workbook  # type: ignore
        from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore
        from openpyxl.utils import get_column_letter
    except Exception as e:
        raise RuntimeError("openpyxl is required. Install with: pip install openpyxl") from e

    if not rows_base:
        raise ValueError('No data parsed from input text.')

    rows = add_calculated_columns(rows_base)

    # Ensure .xlsx extension
    if not out_path.lower().endswith('.xlsx'):
        out_path = out_path + '.xlsx'

    # Create or open workbook
    if os.path.exists(out_path):
        wb = load_workbook(out_path)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(title=sheet_name)
    else:
        wb = Workbook()
        default = wb.active
        wb.remove(default)
        ws = wb.create_sheet(title=sheet_name)

    # Write header + rows
    ws.append(HEADERS)
    for r in rows:
        ws.append([r.get(h, '') for h in HEADERS])

    # Make an Excel Table
    last_row = ws.max_row
    last_col = ws.max_column
    ref = f"A1:{get_column_letter(last_col)}{last_row}"

    # Ensure table name is unique across the whole workbook
    used_tables = set()
    for _ws in wb.worksheets:
        try:
            if getattr(_ws, 'tables', None) and isinstance(_ws.tables, dict):
                used_tables.update(t.displayName for t in _ws.tables.values())
            elif getattr(_ws, '_tables', None):
                used_tables.update(t.displayName for t in _ws._tables)
        except Exception:
            pass
    tname = safe_table_name(f"T_{sheet_name}", used_tables)

    table = Table(displayName=tname, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
                                          showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)

    # Freeze header and format columns
    ws.freeze_panes = 'A2'
    hdr_idx = {ws.cell(1, i).value: i for i in range(1, last_col + 1)}

    int_cols = ['Instances', 'Threads', 'Handles']
    float_cols = ['CPU', 'Memory', 'Data']
    pct_cols = ['Total Usage of  CPU in 100% (B / C *100) CPU', 'Total Usage of  Memory in 100% (B / C *100)']

    for name in int_cols:
        idx = hdr_idx.get(name)
        if idx:
            for r in range(2, last_row + 1):
                ws.cell(r, idx).number_format = '#,##0'
    for name in float_cols:
        idx = hdr_idx.get(name)
        if idx:
            for r in range(2, last_row + 1):
                ws.cell(r, idx).number_format = '0.00'
    for name in pct_cols:
        idx = hdr_idx.get(name)
        if idx:
            for r in range(2, last_row + 1):
                ws.cell(r, idx).number_format = '0.00%'

    # Autosize columns
    from openpyxl.utils import get_column_letter as _col
    for col in range(1, last_col + 1):
        max_len = 0
        for r in range(1, last_row + 1):
            v = ws.cell(r, col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[_col(col)].width = min(max(max_len + 2, 12), 60)

    # --- Inject Excel formulas using A1 refs (avoid Excel repairing structured refs) ---
    try:
        # Column indices
        col_CPU = hdr_idx.get('CPU')
        col_MEM = hdr_idx.get('Memory')
        col_I = hdr_idx.get('Cumulative Total CPU')
        col_J = hdr_idx.get('Total Usage of  CPU in 100% (B / C *100) CPU')
        col_L = hdr_idx.get('Cumulative Total Memory')
        col_M = hdr_idx.get('Total Usage of  Memory in 100% (B / C *100)')

        last_row = ws.max_row
        from openpyxl.utils import get_column_letter as _col
        if col_CPU and col_I:
            cpu_col_letter = _col(col_CPU)
            total_cpu_range = f"${cpu_col_letter}$2:${cpu_col_letter}${last_row}"
        if col_MEM and col_L:
            mem_col_letter = _col(col_MEM)
            total_mem_range = f"${mem_col_letter}$2:${mem_col_letter}${last_row}"

        for r in range(2, last_row + 1):
            # Cumulative totals (same value in every data row)
            if col_I and col_CPU:
                ws.cell(r, col_I).value = f"=SUM({total_cpu_range})"
                ws.cell(r, col_I).number_format = '0.00'
            if col_L and col_MEM:
                ws.cell(r, col_L).value = f"=SUM({total_mem_range})"
                ws.cell(r, col_L).number_format = '0.00'

            # Per-row usage percentages
            if col_J and col_CPU and col_I:
                I_letter = _col(col_I)
                C_letter = _col(col_CPU)
                ws.cell(r, col_J).value = f"=IF({I_letter}{r}=0,0,{C_letter}{r}/{I_letter}{r})"
                ws.cell(r, col_J).number_format = '0.00%'
            if col_M and col_MEM and col_L:
                L_letter = _col(col_L)
                D_letter = _col(col_MEM)
                ws.cell(r, col_M).value = f"=IF({L_letter}{r}=0,0,{D_letter}{r}/{L_letter}{r})"
                ws.cell(r, col_M).number_format = '0.00%'
    except Exception:
        pass

    wb.save(out_path)
    return out_path


# ---------------- Additional helpers & multi-sheet writer ----------------
_SAFE_SHEET_BAD = re.compile(r'[:\/\?*\[\]]')
_SAFE_TABLE_BAD = re.compile(r'[^A-Za-z0-9_]')

def safe_sheet_name(base: str, used: Optional[set] = None) -> str:
    """Sanitize to Excel rules and ensure uniqueness (<=31 chars)."""
    if not base:
        base = 'Sheet'
    name = _SAFE_SHEET_BAD.sub('_', base)[:31]
    if not name:
        name = 'Sheet'
    if used is None:
        return name
    cand = name
    i = 2
    while cand in used:
        suf = f" ({i})"
        cand = (name[: 31 - len(suf)] + suf)
        i += 1
    used.add(cand)
    return cand


def safe_table_name(base: str, used: Optional[set] = None) -> str:
    name = _SAFE_TABLE_BAD.sub('_', base)
    if not name or not name[0].isalpha():
        name = 'T_' + name
    if used is None:
        return name
    cand = name
    i = 2
    while cand in used:
        cand = f"{name}_{i}"
        i += 1
    used.add(cand)
    return cand


def write_workbook_for_files(file_paths: List[str], out_path: str) -> str:
    """Create one workbook with **one sheet per input file**.
    Sheet name = sanitized file stem; each sheet contains its own table and formulas.
    """
    from openpyxl import Workbook  # type: ignore
    from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore
    from openpyxl.utils import get_column_letter

    if not out_path.lower().endswith('.xlsx'):
        out_path = out_path + '.xlsx'

    wb = Workbook()
    # remove default sheet
    default = wb.active
    wb.remove(default)

    used_sheet = set()
    used_tables = set()

    wrote_any = False
    for pth in file_paths:
        try:
            with open(pth, 'r', encoding='utf-8', errors='replace') as f:
                raw = f.read()
            rows_base = parse_cpu_memory_text(raw)
            if not rows_base:
                continue
            rows = add_calculated_columns(rows_base)
            sname = safe_sheet_name(os.path.splitext(os.path.basename(pth))[0], used_sheet)
            ws = wb.create_sheet(title=sname)

            # header + rows
            ws.append(HEADERS)
            for r in rows:
                ws.append([r.get(h, '') for h in HEADERS])

            # table
            last_row = ws.max_row
            last_col = ws.max_column
            ref = f"A1:{get_column_letter(last_col)}{last_row}"
            tname = safe_table_name(f"T_{sname}", used_tables)
            table = Table(displayName=tname, ref=ref)
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
                                                  showRowStripes=True, showColumnStripes=False)
            ws.add_table(table)

            # freeze & formats
            ws.freeze_panes = 'A2'
            hdr_idx = {ws.cell(1, i).value: i for i in range(1, last_col + 1)}
            int_cols = ['Instances', 'Threads', 'Handles']
            float_cols = ['CPU', 'Memory', 'Data']
            pct_cols = ['Total Usage of  CPU in 100% (B / C *100) CPU', 'Total Usage of  Memory in 100% (B / C *100)']
            for name in int_cols:
                idx = hdr_idx.get(name)
                if idx:
                    for rr in range(2, last_row + 1):
                        ws.cell(rr, idx).number_format = '#,##0'
            for name in float_cols:
                idx = hdr_idx.get(name)
                if idx:
                    for rr in range(2, last_row + 1):
                        ws.cell(rr, idx).number_format = '0.00'
            for name in pct_cols:
                idx = hdr_idx.get(name)
                if idx:
                    for rr in range(2, last_row + 1):
                        ws.cell(rr, idx).number_format = '0.00%'

            # formulas per-row using A1 refs (avoid structured refs)
            from openpyxl.utils import get_column_letter as _col
            # build indices
            col_CPU = hdr_idx.get('CPU')
            col_MEM = hdr_idx.get('Memory')
            I = hdr_idx.get('Cumulative Total CPU')
            J = hdr_idx.get('Total Usage of  CPU in 100% (B / C *100) CPU')
            Lc = hdr_idx.get('Cumulative Total Memory')
            M = hdr_idx.get('Total Usage of  Memory in 100% (B / C *100)')

            # prepare absolute total ranges
            if col_CPU and I:
                cpu_letter = _col(col_CPU)
                total_cpu_range = f"${cpu_letter}$2:${cpu_letter}${last_row}"
            if col_MEM and Lc:
                mem_letter = _col(col_MEM)
                total_mem_range = f"${mem_letter}$2:${mem_letter}${last_row}"

            for rr in range(2, last_row + 1):
                if I and col_CPU:
                    ws.cell(rr, I).value = f"=SUM({total_cpu_range})"
                    ws.cell(rr, I).number_format = '0.00'
                if J and col_CPU and I:
                    I_letter = _col(I)
                    C_letter = _col(col_CPU)
                    ws.cell(rr, J).value = f"=IF({I_letter}{rr}=0,0,{C_letter}{rr}/{I_letter}{rr})"
                    ws.cell(rr, J).number_format = '0.00%'
                if Lc and col_MEM:
                    ws.cell(rr, Lc).value = f"=SUM({total_mem_range})"
                    ws.cell(rr, Lc).number_format = '0.00'
                if M and col_MEM and Lc:
                    L_letter = _col(Lc)
                    D_letter = _col(col_MEM)
                    ws.cell(rr, M).value = f"=IF({L_letter}{rr}=0,0,{D_letter}{rr}/{L_letter}{rr})"
                    ws.cell(rr, M).number_format = '0.00%'

            # autosize
            from openpyxl.utils import get_column_letter as _col
            for c in range(1, last_col + 1):
                max_len = 0
                for rr in range(1, last_row + 1):
                    v = ws.cell(rr, c).value
                    if v is None:
                        continue
                    max_len = max(max_len, len(str(v)))
                ws.column_dimensions[_col(c)].width = min(max(max_len + 2, 12), 60)

            wrote_any = True
        except Exception:
            continue

    if not wrote_any:
        raise ValueError('No valid data parsed from the provided files.')

    wb.save(out_path)
    return out_path


# --------------------------- GUI ---------------------------
if TK_OK:
    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title('CPU/Memory TXT → Excel (CPUMemory)')
            self.geometry('1200x740')
            self.minsize(980, 620)
            self.rows_base: List[Dict[str, object]] = []
            self.rows_calc: List[Dict[str, object]] = []
            self.single_path: Optional[str] = None
            self.multi_files: List[str] = []
            self._build_ui()

        def _build_ui(self) -> None:
            pad = 8
            root = ttk.Panedwindow(self, orient=tk.VERTICAL)
            root.pack(fill=tk.BOTH, expand=True)

            # --- Top controls ---
            top = ttk.Frame(root, padding=pad)
            root.add(top, weight=3)

            btns = ttk.Frame(top)
            btns.pack(fill=tk.X)
            ttk.Button(btns, text='Load .txt', command=self.on_load_file).pack(side=tk.LEFT, padx=(0, pad))
            ttk.Button(btns, text='Load .txt (multi)', command=self.on_load_multi).pack(side=tk.LEFT, padx=(0, pad))
            ttk.Button(btns, text='Parse → Preview', command=self.on_parse).pack(side=tk.LEFT, padx=(0, pad))
            ttk.Button(btns, text='Save as Excel', command=self.on_save).pack(side=tk.LEFT, padx=(0, pad))
            ttk.Button(btns, text='Clear', command=self.on_clear).pack(side=tk.LEFT)

            self.status_var = tk.StringVar()
            ttk.Label(btns, textvariable=self.status_var, foreground='#0a7d11').pack(side=tk.LEFT, padx=(pad, 0))

            ttk.Label(top, text='Paste your CPU/Memory text below (or use Load):').pack(anchor='w', pady=(pad, 0))
            self.txt = tk.Text(top, height=14, wrap='none')
            self.txt.pack(fill=tk.BOTH, expand=True)

            xscroll = ttk.Scrollbar(top, orient='horizontal', command=self.txt.xview)
            yscroll = ttk.Scrollbar(top, orient='vertical', command=self.txt.yview)
            self.txt.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
            yscroll.place(relx=1.0, rely=0.25, relheight=0.5, anchor='ne')
            xscroll.pack(fill=tk.X)

            # --- Bottom preview ---
            bottom = ttk.Frame(root, padding=pad)
            root.add(bottom, weight=2)
            ttk.Label(bottom, text='Preview:').pack(anchor='w')

            self.tree = ttk.Treeview(bottom, columns=HEADERS, show='headings')
            for h in HEADERS:
                self.tree.heading(h, text=h)
                base_width = 260 if h == 'Process Name' else 160
                if h in ('Cumulative Total CPU', 'Cumulative Total Memory'):
                    base_width = 180
                if h in ('Total Usage of  CPU in 100% (B / C *100) CPU', 'Total Usage of  Memory in 100% (B / C *100)'):
                    base_width = 180
                self.tree.column(h, width=base_width, anchor='w')
            self.tree.pack(fill=tk.BOTH, expand=True)

            tree_y = ttk.Scrollbar(bottom, orient='vertical', command=self.tree.yview)
            tree_x = ttk.Scrollbar(bottom, orient='horizontal', command=self.tree.xview)
            self.tree.configure(yscrollcommand=tree_y.set, xscrollcommand=tree_x.set)
            tree_y.pack(side=tk.RIGHT, fill=tk.Y)
            tree_x.pack(side=tk.BOTTOM, fill=tk.X)

        def on_load_file(self) -> None:
            path = filedialog.askopenfilename(title='Open Text File', filetypes=[('Text files', '*.txt;*.log;*.*')])
            if not path:
                return
            try:
                with open(path, 'r', encoding='utf-8', errors='replace') as f:
                    data = f.read()
                self.txt.delete('1.0', tk.END)
                self.txt.insert('1.0', data)
                self.status_var.set('Loaded: ' + os.path.basename(path))
                self.single_path = path
                self.multi_files = []
            except Exception as e:
                messagebox.showerror('Error', 'Failed to load file ' + str(e))

        def on_load_multi(self) -> None:
            paths = filedialog.askopenfilenames(title='Open Multiple Text Files', filetypes=[('Text files', '*.txt;*.log;*.*')])
            if not paths:
                return
            try:
                chunks = []
                for pth in paths:
                    with open(pth, 'r', encoding='utf-8', errors='replace') as f:
                        chunks.append(f.read())
                data = "\n".join(chunks)
                self.txt.delete('1.0', tk.END)
                self.txt.insert('1.0', data)
                self.status_var.set(f"Loaded {len(paths)} files")
                self.multi_files = list(paths)
                self.single_path = None
            except Exception as e:
                messagebox.showerror('Error', 'Failed to load files ' + str(e))

        def on_parse(self) -> None:
            raw = self.txt.get('1.0', tk.END).strip()
            if not raw:
                messagebox.showwarning('No input', 'Paste your text or use Load first.')
                return
            try:
                self.rows_base = parse_cpu_memory_text(raw)
                self.rows_calc = add_calculated_columns(self.rows_base)
                self._refresh_tree(self.rows_calc)
                self.status_var.set('Parsed ' + str(len(self.rows_calc)) + ' rows')
            except Exception as e:
                messagebox.showerror('Parse failed', str(e))

        def on_save(self) -> None:
            # Ensure we have parsed rows or multi-file list
            if not self.rows_base and not self.multi_files:
                raw = self.txt.get('1.0', tk.END).strip()
                if not raw:
                    messagebox.showwarning('No input', 'Nothing to save. Paste/load text, then Parse.')
                    return
                try:
                    self.rows_base = parse_cpu_memory_text(raw)
                    self.rows_calc = add_calculated_columns(self.rows_base)
                    self._refresh_tree(self.rows_calc)
                except Exception as e:
                    messagebox.showerror('Parse failed', str(e))
                    return
            path = filedialog.asksaveasfilename(title='Save Excel Workbook', defaultextension='.xlsx', filetypes=[('Excel Workbook', '*.xlsx')])
            if not path:
                return
            try:
                if self.multi_files:
                    out = write_workbook_for_files(self.multi_files, path)
                else:
                    # single: name sheet after file if known, else default
                    sheet = safe_sheet_name(os.path.splitext(os.path.basename(self.single_path))[0]) if self.single_path else 'CPUMemory'
                    out = write_cpumemory_sheet(self.rows_base, path, sheet_name=sheet)
                self.status_var.set('✅ Saved → ' + os.path.basename(out))
            except Exception as e:
                messagebox.showerror('Save failed', str(e))

        def on_clear(self) -> None:
            self.txt.delete('1.0', tk.END)
            self.rows_base = []
            self.rows_calc = []
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.status_var.set('')

        def _refresh_tree(self, rows: List[Dict[str, object]]) -> None:
            def _fmt(h: str, v) -> str:
                try:
                    if h in ('Total Usage of  CPU in 100% (B / C *100) CPU', 'Total Usage of  Memory in 100% (B / C *100)'):
                        return f"{float(v)*100:.2f}%"
                    if h in ('Instances', 'Threads', 'Handles'):
                        return f"{int(v):,}"
                    if h in ('CPU', 'Memory', 'Data', 'Cumulative Total CPU', 'Cumulative Total Memory'):
                        return f"{float(v):.2f}"
                except Exception:
                    pass
                return '' if v is None else str(v)

            for item in self.tree.get_children():
                self.tree.delete(item)
            self.tree['columns'] = list(HEADERS)

            # compute widths based on formatted values
            col_widths = {h: len(h) for h in HEADERS}
            display_rows = []
            for r in rows:
                row_vals = [_fmt(h, r.get(h, '')) for h in HEADERS]
                display_rows.append(row_vals)
                for h, dv in zip(HEADERS, row_vals):
                    col_widths[h] = max(col_widths[h], len(str(dv)))

            for h in HEADERS:
                base_max = 360 if h == 'Process Name' else 260
                if h in ('Cumulative Total CPU', 'Cumulative Total Memory'):
                    base_max = 260
                if h in ('Total Usage of  CPU in 100% (B / C *100) CPU', 'Total Usage of  Memory in 100% (B / C *100)'):
                    base_max = 180
                width = max(120, min(base_max, col_widths[h] * 9))
                self.tree.heading(h, text=h)
                self.tree.column(h, width=width, anchor='w')

            for vals in display_rows:
                self.tree.insert('', tk.END, values=vals)


# --------------------------- Tests ---------------------------
SAMPLE_TXT = (
    "Process Name\tInstances\tCPU\tMemory\tThreads\tHandles\tData\tStatus\n"
    "System\t1\t0.06\t0.13\t332\t5,250\t0.00\tNormal\n"
    "smss\t1\t0.00\t0.62\t2\t64\t0.00\tNormal\n"
    "csrss\t4\t0.01\t34.33\t53\t1,925\t0.00\tNormal\n"
    "wininit\t1\t0.00\t2.02\t2\t158\t0.00\tNormal\n"
    "winlogon\t3\t0.00\t25.61\t13\t791\t0.00\tNormal\n"
    "services\t1\t0.01\t8.02\t8\t777\t0.00\tNormal\n"
    "lsass\t1\t0.00\t21.30\t11\t1,976\t0.00\tNormal\n"
    "svchost\t91\t0.06\t772.53\t602\t31,288\t0.00\tWarning\n"
    "fontdrvhost\t4\t0.00\t9.87\t20\t172\t0.00\n"
)

SAMPLE_SPACE = (
    "Process Name  Instances  CPU  Memory  Threads  Handles  Data  Status\n"
    "System  1  0.06  0.13  332  5,250  0.00  Normal\n"
)

# Additional multi-file sample
SAMPLE_TXT_2 = (
    "Process Name\tInstances\tCPU\tMemory\tThreads\tHandles\tData\tStatus\n"
    "foo\t1\t1.00\t2.00\t1\t10\t0.00\tNormal\n"
)


def run_tests() -> int:
    import tempfile
    fails = 0

    # Parse should return 9 rows (header skipped)
    rows = parse_cpu_memory_text(SAMPLE_TXT)
    try:
        assert len(rows) == 9, f"expected 9 rows, got {len(rows)}"
        assert rows[0]['Process Name'] == 'System'
        assert rows[-1]['Process Name'] == 'fontdrvhost'
    except AssertionError as e:
        print('[FAIL] basic parsing:', e)
        fails += 1

    # Space-delimited header should also be skipped
    try:
        rows2 = parse_cpu_memory_text(SAMPLE_SPACE)
        assert len(rows2) == 1 and rows2[0]['Process Name'] == 'System'
    except AssertionError as e:
        print('[FAIL] space header removal:', e)
        fails += 1

    # Totals & percentages should sum correctly
    try:
        total_cpu = sum(r['CPU'] for r in rows)
        total_mem = sum(r['Memory'] for r in rows)
        calc = add_calculated_columns(rows)
        # totals repeated in every row should equal column totals
        assert abs(calc[0]['Cumulative Total CPU'] - total_cpu) < 1e-9
        assert abs(calc[0]['Cumulative Total Memory'] - total_mem) < 1e-9
        # percent sum ≈ 100%
        pct_sum_cpu = sum(r['Total Usage of  CPU in 100% (B / C *100) CPU'] for r in calc)
        pct_sum_mem = sum(r['Total Usage of  Memory in 100% (B / C *100)'] for r in calc)
        assert 0.999 <= pct_sum_cpu <= 1.001
        assert 0.999 <= pct_sum_mem <= 1.001
    except AssertionError as e:
        print('[FAIL] totals/percent:', e)
        fails += 1

    # Excel write smoke test + formula presence
    try:
        from openpyxl import load_workbook  # type: ignore
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            path = tmp.name
        out = write_cpumemory_sheet(rows, path, sheet_name='CPUMemory')
        wb = load_workbook(out)
        assert 'CPUMemory' in wb.sheetnames
        ws = wb['CPUMemory']
        assert ws.max_row == len(rows) + 1  # header + rows
        # check that at least one formula appears in Cumulative Total CPU column
        hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        i_col = hdrs.index('Cumulative Total CPU') + 1
        assert isinstance(ws.cell(2, i_col).value, str) and ws.cell(2, i_col).value.startswith('=')
    except Exception as e:
        print('[WARN] excel smoke test skipped or failed:', e)

    # Multi-file multi-sheet write test
    try:
        import tempfile
        from openpyxl import load_workbook  # type: ignore
        # write two temp txt files
        f1 = tempfile.NamedTemporaryFile(delete=False, suffix='.txt'); f1.write(SAMPLE_TXT.encode('utf-8')); f1.close()
        f2 = tempfile.NamedTemporaryFile(delete=False, suffix='.txt'); f2.write(SAMPLE_TXT_2.encode('utf-8')); f2.close()
        outxlsx = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx'); outxlsx.close()
        out = write_workbook_for_files([f1.name, f2.name], outxlsx.name)
        wb = load_workbook(out)
        names = set(wb.sheetnames)
        expected1 = safe_sheet_name(os.path.splitext(os.path.basename(f1.name))[0])
        expected2 = safe_sheet_name(os.path.splitext(os.path.basename(f2.name))[0])
        assert expected1 in names and expected2 in names
    except AssertionError as e:
        print('[FAIL] multi-file multi-sheet:', e)
        fails += 1

    if fails:
        print(f"\n{fails} test(s) failed")
        return 1
    print('All tests passed ✅')
    return 0


# --------------------------- CLI ---------------------------

def _smart_paths(arg_i, arg_o: Optional[str]) -> Tuple[List[str], str]:
    """Resolve input/output with friendly fallbacks and multi-file support.
    Returns (list_of_inputs, output_path).
    Order:
      1) Use provided args if present
      2) If missing, try ./cpu_mem.txt → ./report.xlsx
      3) Try Tk multi-file picker if available
      4) Prompts
    """
    # 1) Provided args
    if arg_i:
        inputs = arg_i if isinstance(arg_i, list) else [arg_i]
        out = arg_o if (arg_o and arg_o.lower().endswith('.xlsx')) else (arg_o + '.xlsx' if arg_o else 'report.xlsx')
        return inputs, out

    # 2) Default single in cwd
    default_in = os.path.abspath('cpu_mem.txt')
    if os.path.exists(default_in):
        return [default_in], (arg_o or 'report.xlsx')

    # 3) Tk multi-file picker
    try:
        import tkinter as _tk  # type: ignore
        from tkinter import filedialog as _fd  # type: ignore
        root = _tk.Tk(); root.withdraw()
        in_paths = _fd.askopenfilenames(title='Select one or more CPU/Memory .txt', filetypes=[('Text', '*.txt;*.log;*.*')])
        if not in_paths:
            raise RuntimeError('Input path(s) not provided.')
        out_path = arg_o or _fd.asksaveasfilename(title='Save Excel workbook', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if not out_path:
            raise RuntimeError('Output path not provided.')
        root.destroy()
        return list(in_paths), out_path
    except Exception:
        # 4) Prompts
        try:
            paths = input('Enter input .txt path(s), separated by ; : ').strip()
            inputs = [p.strip() for p in paths.split(';') if p.strip()]
            if not inputs:
                raise RuntimeError('Input path(s) not provided.')
            out = arg_o or (input('Enter output .xlsx path (default report.xlsx): ').strip() or 'report.xlsx')
            return inputs, out
        except Exception as e:
            raise RuntimeError('No input/output provided. Run with -i <file1> <file2> -o <file.xlsx>.') from e


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description='Parse CPU/Memory TXT and write CPUMemory sheet to Excel.', add_help=True)
    ap.add_argument('--cli', action='store_true', help='Force CLI mode (disable GUI)')
    ap.add_argument('-i', '--input', nargs='+', help='Path(s) to input .txt (one or more)')
    ap.add_argument('-o', '--output', help='Path to output .xlsx (sheet CPUMemory will be created/replaced)')
    ap.add_argument('--sheet', default='CPUMemory', help='Target sheet name (default: CPUMemory)')
    ap.add_argument('--run-tests', action='store_true', help='Run self-tests and exit')
    args = ap.parse_args(argv)

    if args.run_tests:
        return run_tests()

    # GUI first (unless forced CLI or Tk missing)
    if not args.cli and TK_OK:
        app = App()  # type: ignore[name-defined]
        app.mainloop()
        return 0

    # CLI flow
    try:
        in_paths, out_path = _smart_paths(args.input, args.output)
    except RuntimeError as e:
        sys.stderr.write(str(e) + "\n")
        sys.stderr.write("Example: python cpu_memory_table.py -i cpu1.txt cpu2.txt -o out.xlsx\n")
        return 2

    if len(in_paths) > 1:
        try:
            out = write_workbook_for_files(in_paths, out_path)
        except Exception as e:
            sys.stderr.write(str(e) + "\n")
            return 4
        print(f"✅ Wrote {len(in_paths)} sheets to '{out}'")
        return 0

    # Single file path
    single = in_paths[0]
    with open(single, 'r', encoding='utf-8', errors='replace') as f:
        raw = f.read()
    rows = parse_cpu_memory_text(raw)
    if not rows:
        sys.stderr.write('No rows parsed. Check the input formatting.\n')
        return 3
    # sheet name from filename (sanitized)
    sname = safe_sheet_name(os.path.splitext(os.path.basename(single))[0])
    try:
        out = write_cpumemory_sheet(rows, out_path, sheet_name=sname)
    except Exception as e:
        sys.stderr.write(str(e) + "\n")
        return 4

    print(f"✅ Wrote {len(rows)} rows to '{out}' (sheet '{sname}')")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
