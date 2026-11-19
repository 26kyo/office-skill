---
name: office
description: >-
  Read, create, and edit Excel (.xlsx), Word (.docx), and PowerPoint (.pptx)
  files via local open-source Python tools with token-compact inspection,
  Excel formula recalculation guardrails, and safe OOXML editing. Use when the
  user asks to inspect, summarize, extract, generate, modify, fill templates,
  or verify Office documents.
---

# Office Skill

Cross-platform (Linux / WSL / macOS / Windows) read & write toolkit for `.xlsx`, `.docx`, and `.pptx`.

## Running the Helper Scripts

Resolve `SKILL_DIR` to the directory containing this `SKILL.md`.

- **Preferred (zero-setup via `uv`)**:
  ```bash
  uv run "$SKILL_DIR/scripts/office_read.py" ...
  uv run "$SKILL_DIR/scripts/xlsx_recalc.py" ...
  uv run "$SKILL_DIR/scripts/docx_xml.py" ...
  ```
- **Fallback (local `.venv` installed from `requirements.txt`)**:
  - Linux / WSL / macOS: `"$SKILL_DIR/.venv/bin/python" "$SKILL_DIR/scripts/..."`
  - Windows: `"$SKILL_DIR/.venv/Scripts/python.exe" "$SKILL_DIR/scripts/..."`
- **Environment check**:
  ```bash
  python3 "$SKILL_DIR/scripts/doctor.py"
  ```

---

## 1. Read-Only Inspection (`scripts/office_read.py`)

Default output is token-compact text/TSV; pass `--raw` for single-line JSON.

### Excel (`.xlsx`)
1. **Workbook overview** (sheet names and dimensions):
   `uv run "$SKILL_DIR/scripts/office_read.py" xlsx path/to/book.xlsx`
2. **Sheet / A1 range** (TSV cached values + `# formulas` showing `formula => value`):
   `uv run "$SKILL_DIR/scripts/office_read.py" xlsx path/to/book.xlsx --sheet "Sheet1" --range "A1:F20"`
   Omitting `--range` caps at `--max-rows 100`.

### Word (`.docx`)
1. **Outline** (headings & table sizes):
   `uv run "$SKILL_DIR/scripts/office_read.py" docx path/to/doc.docx --outline`
2. **Content + tables**:
   `uv run "$SKILL_DIR/scripts/office_read.py" docx path/to/doc.docx --max-paras 120`

### PowerPoint (`.pptx`)
1. **Deck overview**:
   `uv run "$SKILL_DIR/scripts/office_read.py" pptx path/to/deck.pptx`
2. **Slide content** (1-based `1-3`, `2,5`, or `all`; recurses into grouped shapes):
   `uv run "$SKILL_DIR/scripts/office_read.py" pptx path/to/deck.pptx --slides 1-3`
3. **Render slides to PNG** (requires `libreoffice` + `pdftoppm`):
   `uv run "$SKILL_DIR/scripts/office_read.py" pptx path/to/deck.pptx --render 2,5 --out-dir /tmp/pptx-preview`

---

## 2. Writing & Editing Guardrails

For ad-hoc Python creation/editing scripts, run with `uv run --with-requirements "$SKILL_DIR/requirements.txt" python your_script.py` (or the local `.venv` Python).

### Excel (`.xlsx`)
1. **Formulas over hardcoded outputs**: write Excel formulas (`=SUM(...)`, `=AVERAGE(...)`) rather than precomputing totals in Python.
2. **Preserve existing workbooks**: edit via `openpyxl.load_workbook(path)` rather than rebuilding from scratch.
3. **Mandatory recalculation gate**: after creating or modifying any formula workbook, run:
   ```bash
   uv run "$SKILL_DIR/scripts/xlsx_recalc.py" output.xlsx
   ```
   - Requires `libreoffice`; never ship while `error_summary` reports `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, or `#NULL!`.
   - For complex existing workbooks (macros, pivot tables, custom themes), pass `--backup` (saves `.xlsx.bak` before overwriting) or `--no-write` to inspect recalculation results without replacing the original file.

### Word (`.docx`)
1. **Edit existing `.docx` via OOXML when layout, comments, or tracked changes matter**:
   `python-docx` rewrites can drop unsupported OOXML features. Unpack, edit `word/document.xml` (or header/footer/comment XML), validate, and repack:
   ```bash
   uv run "$SKILL_DIR/scripts/docx_xml.py" unpack input.docx /tmp/docx_work
   uv run "$SKILL_DIR/scripts/docx_xml.py" pack /tmp/docx_work output.docx
   ```
   `unpack` preserves original XML bytes by default (pass `--pretty` only if you want formatted XML); `pack` validates all `.xml`/`.rels` files before writing.
2. **New `.docx` / template filling**: use `python-docx` or `docxtpl`.
3. **Dual-width tables (prevent collapsed columns across Word, LibreOffice, and Google Docs)**:
   Disable autofit (`fixed` layout) and set widths in **both** `w:tblGrid/w:gridCol` AND every cell's `w:tcPr/w:tcW` in DXA (`1 inch = 1440 twips`):
   ```python
   from docx.oxml import OxmlElement
   from docx.oxml.ns import qn

   def set_table_col_widths_dxa(table, col_widths_dxa):
       table.autofit = False
       tblPr = table._tbl.tblPr
       layout = tblPr.find(qn("w:tblLayout"))
       if layout is None:
           layout = OxmlElement("w:tblLayout")
           tblPr.append(layout)
       layout.set(qn("w:type"), "fixed")
       tblGrid = table._tbl.tblGrid
       for g in list(tblGrid):
           tblGrid.remove(g)
       for w in col_widths_dxa:
           gc = OxmlElement("w:gridCol")
           gc.set(qn("w:w"), str(w))
           tblGrid.append(gc)
       for row in table.rows:
           for cell, w in zip(row.cells, col_widths_dxa):
               tcPr = cell._tc.get_or_add_tcPr()
               tcW = tcPr.find(qn("w:tcW"))
               if tcW is None:
                   tcW = OxmlElement("w:tcW")
                   tcPr.append(tcW)
               tcW.set(qn("w:w"), str(w))
               tcW.set(qn("w:type"), "dxa")
   ```

### Visual Layout Check (when formatting/layout matters)
```bash
libreoffice --headless --convert-to pdf --outdir /tmp/office-preview output.docx
pdftoppm -png -r 150 /tmp/office-preview/output.pdf /tmp/office-preview/page
```
Then inspect the rendered PNGs with your file-reading/image tool.
