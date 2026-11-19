#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "openpyxl>=3.1",
#   "python-docx>=1.1",
#   "python-pptx>=1.0",
# ]
# ///
"""Read-only, token-compact inspector for .xlsx, .pptx, and .docx files."""

import argparse
import datetime
import json
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
import openpyxl
from openpyxl.utils import get_column_letter
from pptx import Presentation


def fmt_cell(val) -> str:
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.isoformat()
    s = str(val)
    return (
        s.replace("\\", "\\\\")
        .replace("\t", "\\t")
        .replace("\r\n", "\\n")
        .replace("\n", "\\n")
    )


def install_hint_render() -> str:
    sys_name = platform.system()
    if sys_name == "Darwin":
        return "brew install --cask libreoffice && brew install poppler"
    if sys_name == "Windows":
        return "winget install TheDocumentFoundation.LibreOffice && winget install oschwartz10612.Poppler"
    return "sudo apt update && sudo apt install -y libreoffice poppler-utils"


def find_soffice() -> str | None:
    found = shutil.which("libreoffice") or shutil.which("soffice")
    if found:
        return found
    candidates = [
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
    ]
    for p in candidates:
        if p.is_file():
            return str(p)
    return None


def parse_pages(spec: str, total: int) -> list[int]:
    if spec == "all":
        return list(range(1, total + 1))
    pages: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a_str, b_str = part.split("-", 1)
            a, b = int(a_str), int(b_str)
            if a < 1 or b < a or b > total:
                sys.exit(f"Invalid page range '{part}' for {total} slides")
            pages.update(range(a, b + 1))
        elif part.isdigit():
            p = int(part)
            if p < 1 or p > total:
                sys.exit(f"Page {p} out of range (1..{total})")
            pages.add(p)
        else:
            sys.exit(f"Invalid page spec: '{part}'")
    return sorted(pages)


def parse_a1_range(range_str: str) -> tuple[int, int, int, int]:
    from openpyxl.utils.cell import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    return min_row, min_col, max_row, max_col


def cmd_xlsx(args: argparse.Namespace) -> None:
    wb_val = openpyxl.load_workbook(args.file, read_only=True, data_only=True)
    try:
        if not args.sheet:
            sheets = []
            for idx, name in enumerate(wb_val.sheetnames, 1):
                ws = wb_val[name]
                dim = ws.calculate_dimension()
                sheets.append(
                    {
                        "index": idx,
                        "name": name,
                        "dimensions": dim,
                        "max_row": ws.max_row,
                        "max_column": ws.max_column,
                    }
                )
            if args.raw:
                print(json.dumps({"file": args.file, "sheets": sheets}, ensure_ascii=False))
                return
            lines = [f"file: {args.file} | sheets: {len(sheets)}"]
            for s in sheets:
                lines.append(
                    f"{s['index']}. {s['name']} | {s['dimensions']} ({s['max_row']}x{s['max_column']})"
                )
            print("\n".join(lines))
            return

        if args.sheet not in wb_val.sheetnames:
            sys.exit(f"Sheet '{args.sheet}' not found. Available: {', '.join(wb_val.sheetnames)}")

        ws_val = wb_val[args.sheet]
        wb_frm = openpyxl.load_workbook(args.file, read_only=True, data_only=False)
        try:
            ws_frm = wb_frm[args.sheet]
            truncated = False
            if args.range:
                r1, c1, r2, c2 = parse_a1_range(args.range)
                range_label = args.range.upper()
            else:
                r1, c1 = 1, 1
                total_rows = ws_val.max_row or 1
                c2 = ws_val.max_column or 1
                r2 = min(total_rows, args.max_rows)
                if total_rows > args.max_rows:
                    truncated = True
                range_label = f"A1:{get_column_letter(c2)}{r2}"

            val_rows = [
                list(row)
                for row in ws_val.iter_rows(
                    min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True
                )
            ]
            frm_rows = [
                list(row)
                for row in ws_frm.iter_rows(
                    min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True
                )
            ]
        finally:
            wb_frm.close()
    finally:
        wb_val.close()

    formulas = []
    for dr, row in enumerate(frm_rows):
        for dc, cell_val in enumerate(row):
            if isinstance(cell_val, str) and cell_val.startswith("="):
                addr = f"{get_column_letter(c1 + dc)}{r1 + dr}"
                cached = val_rows[dr][dc] if dr < len(val_rows) and dc < len(val_rows[dr]) else None
                formulas.append({"cell": addr, "formula": cell_val, "value": cached})

    if args.raw:
        payload = {
            "file": args.file,
            "sheet": args.sheet,
            "range": range_label,
            "truncated": truncated,
            "values": [
                [v.isoformat() if isinstance(v, (datetime.datetime, datetime.date)) else v for v in r]
                for r in val_rows
            ],
            "formulas": [
                {
                    **f,
                    "value": (
                        f["value"].isoformat()
                        if isinstance(f["value"], (datetime.datetime, datetime.date))
                        else f["value"]
                    ),
                }
                for f in formulas
            ],
        }
        print(json.dumps(payload, ensure_ascii=False))
        return

    trunc_note = f" [truncated from {ws_val.max_row} rows]" if truncated else ""
    lines = [
        f"sheet: {args.sheet} | range: {range_label} ({len(val_rows)}x{c2 - c1 + 1}){trunc_note}",
        "# values (TSV)",
    ]
    for r in val_rows:
        lines.append("\t".join(fmt_cell(v) for v in r))
    lines.append(f"# formulas ({len(formulas)})")
    for f in formulas:
        lines.append(f"{f['cell']}: {f['formula']} => {fmt_cell(f['value'])}")
    print("\n".join(lines))


def _walk_shapes(shapes, items: list[dict]) -> None:
    for shape in shapes:
        if shape.shape_type == 6:  # GROUP
            _walk_shapes(shape.shapes, items)
        elif shape.has_table:
            rows = []
            for row in shape.table.rows:
                rows.append("\t".join(fmt_cell(c.text.strip()) for c in row.cells))
            items.append({"kind": "table", "name": shape.name, "text": "\n".join(rows)})
        elif shape.has_chart:
            title = shape.chart.chart_title.text_frame.text.strip() if shape.chart.has_title else ""
            items.append({"kind": "chart", "name": shape.name, "text": title})
        elif shape.shape_type == 13:  # PICTURE
            items.append({"kind": "image", "name": shape.name, "text": ""})
        elif shape.has_text_frame:
            text = "\n".join(p.text for p in shape.text_frame.paragraphs).strip()
            if text:
                items.append({"kind": "text", "name": shape.name, "text": text})


def extract_slide_elements(slide) -> list[dict]:
    items: list[dict] = []
    _walk_shapes(slide.shapes, items)
    return items


def render_pptx_pages(pptx_path: str, pages: list[int], out_dir: str) -> list[dict]:
    soffice = find_soffice()
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        sys.exit(
            "Rendering requires LibreOffice and pdftoppm. Install with:\n"
            f"  {install_hint_render()}"
        )
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    rendered = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        prof_dir = tmp_path / "lo_profile"
        pdf_dir = tmp_path / "pdf"
        prof_dir.mkdir(parents=True, exist_ok=True)
        pdf_dir.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                f"-env:UserInstallation={prof_dir.as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_dir),
                pptx_path,
            ],
            check=True,
            timeout=120,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
        )
        pdf_files = list(pdf_dir.glob("*.pdf"))
        if not pdf_files:
            sys.exit("LibreOffice PDF conversion produced no output")
        pdf = pdf_files[0]
        for page in pages:
            prefix = out / f"slide-{page:02d}"
            subprocess.run(
                [
                    pdftoppm,
                    "-png",
                    "-r",
                    "150",
                    "-singlefile",
                    "-f",
                    str(page),
                    "-l",
                    str(page),
                    str(pdf),
                    str(prefix),
                ],
                check=True,
                timeout=60,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
            )
            png_path = Path(f"{prefix}.png")
            rendered.append({"page": page, "path": str(png_path), "bytes": png_path.stat().st_size})
    return rendered


def cmd_pptx(args: argparse.Namespace) -> None:
    prs = Presentation(args.file)
    total = len(prs.slides)

    if not args.slides and not args.render:
        overview = []
        for idx, slide in enumerate(prs.slides, 1):
            title = slide.shapes.title.text.strip() if slide.shapes.title else ""
            counts: dict[str, int] = {}
            for el in extract_slide_elements(slide):
                counts[el["kind"]] = counts.get(el["kind"], 0) + 1
            has_notes = bool(slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip())
            overview.append({"page": idx, "title": title, "kinds": counts, "has_notes": has_notes})
        if args.raw:
            print(json.dumps({"file": args.file, "slideCount": total, "slides": overview}, ensure_ascii=False))
            return
        lines = [f"file: {args.file} | slides: {total}"]
        for s in overview:
            kinds = ",".join(f"{k}:{v}" for k, v in s["kinds"].items()) or "empty"
            notes = " [notes]" if s["has_notes"] else ""
            title = f" {s['title']}" if s["title"] else ""
            lines.append(f"{s['page']}.{title} [{kinds}]{notes}")
        print("\n".join(lines))
        return

    slides_out = []
    if args.slides:
        for page in parse_pages(args.slides, total):
            slide = prs.slides[page - 1]
            title = slide.shapes.title.text.strip() if slide.shapes.title else ""
            notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
            slides_out.append(
                {
                    "page": page,
                    "title": title,
                    "layout": slide.slide_layout.name,
                    "elements": extract_slide_elements(slide),
                    "notes": notes or None,
                }
            )

    rendered = []
    if args.render:
        pages = parse_pages(args.render, total)
        rendered = render_pptx_pages(args.file, pages, args.out_dir)

    if args.raw:
        print(
            json.dumps(
                {"file": args.file, "slideCount": total, "slides": slides_out, "rendered": rendered},
                ensure_ascii=False,
            )
        )
        return

    lines = [f"file: {args.file} | slides: {total}"]
    for s in slides_out:
        lines.append(f"=== Slide {s['page']}: {s['title']} ({s['layout']}) ===")
        for el in s["elements"]:
            header = f"[{el['kind']}:{el['name']}]"
            if "\n" in el["text"]:
                lines.extend([header, el["text"]])
            elif el["text"]:
                lines.append(f"{header} {el['text']}")
            else:
                lines.append(header)
        if s["notes"]:
            lines.append(f"[notes] {s['notes']}")
    for r in rendered:
        lines.append(f"rendered: slide-{r['page']:02d} -> {r['path']} ({r['bytes']}B)")
    print("\n".join(lines))


def cmd_docx(args: argparse.Namespace) -> None:
    doc = Document(args.file)
    paras = [(i + 1, p.style.name if p.style else "", p.text.strip()) for i, p in enumerate(doc.paragraphs)]
    non_empty = [(i, st, txt) for i, st, txt in paras if txt]

    if args.outline:
        headings = [(i, st, txt) for i, st, txt in non_empty if st.startswith("Heading") or st == "Title"]
        if args.raw:
            print(
                json.dumps(
                    {"file": args.file, "paragraphs": len(paras), "tables": len(doc.tables), "headings": headings},
                    ensure_ascii=False,
                )
            )
            return
        lines = [
            f"file: {args.file} | paragraphs: {len(paras)} (non-empty: {len(non_empty)}) | tables: {len(doc.tables)}"
        ]
        for i, st, txt in headings:
            lines.append(f"[P{i} {st}] {txt}")
        for ti, tbl in enumerate(doc.tables, 1):
            lines.append(f"[Table {ti}: {len(tbl.rows)}x{len(tbl.columns)}]")
        print("\n".join(lines))
        return

    shown_paras = non_empty[: args.max_paras]
    tables_out = []
    for ti, tbl in enumerate(doc.tables, 1):
        rows = [[fmt_cell(c.text.strip()) for c in row.cells] for row in tbl.rows]
        tables_out.append({"index": ti, "rows": rows})

    if args.raw:
        print(
            json.dumps(
                {
                    "file": args.file,
                    "paragraphs": shown_paras,
                    "tables": tables_out,
                    "truncated": len(non_empty) > args.max_paras,
                },
                ensure_ascii=False,
            )
        )
        return

    lines = [f"file: {args.file} | paragraphs: {len(non_empty)} | tables: {len(doc.tables)}"]
    for i, st, txt in shown_paras:
        tag = f"[{st}] " if st.startswith("Heading") or st == "Title" else ""
        lines.append(f"P{i}: {tag}{txt}")
    if len(non_empty) > args.max_paras:
        lines.append(f"... [truncated {len(non_empty) - args.max_paras} paragraphs; pass --max-paras]")
    for tbl in tables_out:
        lines.append(f"# Table {tbl['index']} ({len(tbl['rows'])}x{len(tbl['rows'][0]) if tbl['rows'] else 0})")
        for r in tbl["rows"]:
            lines.append("\t".join(r))
    print("\n".join(lines))


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only Office document inspector")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_xlsx = sub.add_parser("xlsx")
    p_xlsx.add_argument("file")
    p_xlsx.add_argument("--sheet")
    p_xlsx.add_argument("--range")
    p_xlsx.add_argument("--max-rows", type=int, default=100)
    p_xlsx.add_argument("--raw", action="store_true")

    p_pptx = sub.add_parser("pptx")
    p_pptx.add_argument("file")
    p_pptx.add_argument("--slides")
    p_pptx.add_argument("--render")
    p_pptx.add_argument("--out-dir", default=str(Path(tempfile.gettempdir()) / "office-read-render"))
    p_pptx.add_argument("--raw", action="store_true")

    p_docx = sub.add_parser("docx")
    p_docx.add_argument("file")
    p_docx.add_argument("--outline", action="store_true")
    p_docx.add_argument("--max-paras", type=int, default=120)
    p_docx.add_argument("--raw", action="store_true")

    args = parser.parse_args()
    if args.cmd == "xlsx":
        cmd_xlsx(args)
    elif args.cmd == "pptx":
        cmd_pptx(args)
    elif args.cmd == "docx":
        cmd_docx(args)


if __name__ == "__main__":
    main()
