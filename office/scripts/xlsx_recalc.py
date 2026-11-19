#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "openpyxl>=3.1",
# ]
# ///
"""Recalculate .xlsx formulas via headless LibreOffice and scan for Excel errors."""

import argparse
import json
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

EXCEL_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def install_hint_libreoffice() -> str:
    sys_name = platform.system()
    if sys_name == "Darwin":
        return "brew install --cask libreoffice"
    if sys_name == "Windows":
        return "winget install TheDocumentFoundation.LibreOffice"
    return "sudo apt update && sudo apt install -y libreoffice"


def find_soffice() -> str:
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
    sys.exit(f"LibreOffice not installed. Install with:\n  {install_hint_libreoffice()}")


def recalc_copy(soffice: str, xlsx_path: Path, out_dir: Path, timeout: int = 60) -> Path:
    prof_dir = out_dir / "lo_profile"
    conv_dir = out_dir / "lo_out"
    prof_dir.mkdir(parents=True, exist_ok=True)
    conv_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        f"-env:UserInstallation={prof_dir.as_uri()}",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(conv_dir),
        str(xlsx_path),
    ]
    res = subprocess.run(cmd, timeout=timeout, capture_output=True, text=True)
    if res.returncode != 0:
        sys.exit(f"LibreOffice recalculation failed (rc={res.returncode}): {res.stderr.strip()}")
    out_file = conv_dir / xlsx_path.name
    if not out_file.is_file():
        sys.exit(f"LibreOffice did not produce recalculated file: {out_file}")
    return out_file


def scan_errors(xlsx_path: Path) -> dict:
    wb_val = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    wb_frm = openpyxl.load_workbook(xlsx_path, data_only=False, read_only=True)
    error_summary: dict[str, list[str]] = {}
    formula_count = 0
    total_errors = 0
    try:
        for sheet_name in wb_val.sheetnames:
            ws_v = wb_val[sheet_name]
            ws_f = wb_frm[sheet_name]
            for r_idx, (row_v, row_f) in enumerate(
                zip(ws_v.iter_rows(values_only=True), ws_f.iter_rows(values_only=True)), 1
            ):
                for c_idx, (val, frm) in enumerate(zip(row_v, row_f), 1):
                    if isinstance(frm, str) and frm.startswith("="):
                        formula_count += 1
                    if isinstance(val, str):
                        for err in EXCEL_ERRORS:
                            if err in val:
                                addr = f"{sheet_name}!{openpyxl.utils.get_column_letter(c_idx)}{r_idx}"
                                error_summary.setdefault(err, []).append(addr)
                                total_errors += 1
                                break
    finally:
        wb_val.close()
        wb_frm.close()

    return {
        "file": str(xlsx_path),
        "status": "errors_found" if total_errors else "success",
        "total_formulas": formula_count,
        "total_errors": total_errors,
        "error_summary": error_summary,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Recalculate .xlsx formulas via headless LibreOffice and scan for Excel errors."
    )
    parser.add_argument("file", help="Path to the .xlsx workbook")
    parser.add_argument("--scan-only", action="store_true", help="Scan cached values without running LibreOffice")
    parser.add_argument("--no-write", action="store_true", help="Recalculate in temp dir without overwriting the original file")
    parser.add_argument("--backup", action="store_true", help="Save a .bak copy before overwriting the original file")
    parser.add_argument("--timeout", type=int, default=60, help="LibreOffice timeout in seconds (default: 60)")
    args = parser.parse_args()

    xlsx_path = Path(args.file).resolve()
    if not xlsx_path.is_file():
        sys.exit(f"File not found: {xlsx_path}")

    if args.scan_only:
        report = scan_errors(xlsx_path)
    else:
        soffice = find_soffice()
        with tempfile.TemporaryDirectory() as tmp:
            recalc_file = recalc_copy(soffice, xlsx_path, Path(tmp), timeout=args.timeout)
            report = scan_errors(recalc_file)
            report["file"] = str(xlsx_path)
            if not args.no_write:
                if args.backup:
                    backup_path = xlsx_path.with_suffix(xlsx_path.suffix + ".bak")
                    shutil.copyfile(xlsx_path, backup_path)
                    report["backup"] = str(backup_path)
                shutil.copyfile(recalc_file, xlsx_path)

    print(json.dumps(report, ensure_ascii=False))
    if report["total_errors"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
