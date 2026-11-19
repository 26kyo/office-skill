#!/usr/bin/env python3
"""Environment self-check for the office agent skill (zero external dependencies)."""

import argparse
import importlib.util
import json
import platform
import shutil
import sys
from pathlib import Path

REQUIRED_MODULES = ("openpyxl", "docx", "pptx", "lxml")
OPTIONAL_MODULES = ("pandas", "docxtpl")


def find_soffice() -> str | None:
    found = shutil.which("libreoffice") or shutil.which("soffice")
    if found:
        return found
    for p in (
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
    ):
        if p.is_file():
            return str(p)
    return None


def install_commands(sys_name: str) -> dict[str, str]:
    if sys_name == "Darwin":
        return {
            "uv": "brew install uv",
            "libreoffice": "brew install --cask libreoffice",
            "pdftoppm": "brew install poppler",
        }
    if sys_name == "Windows":
        return {
            "uv": "winget install astral-sh.uv",
            "libreoffice": "winget install TheDocumentFoundation.LibreOffice",
            "pdftoppm": "winget install oschwartz10612.Poppler",
        }
    return {
        "uv": "curl -LsSf https://astral.sh/uv/install.sh | sh",
        "libreoffice": "sudo apt update && sudo apt install -y libreoffice",
        "pdftoppm": "sudo apt update && sudo apt install -y poppler-utils",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Check Office skill runtime dependencies")
    parser.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    args = parser.parse_args()

    sys_name = platform.system()
    cmds = install_commands(sys_name)

    uv_path = shutil.which("uv")
    soffice_path = find_soffice()
    pdftoppm_path = shutil.which("pdftoppm")

    mods = {m: importlib.util.find_spec(m) is not None for m in REQUIRED_MODULES + OPTIONAL_MODULES}
    missing_req = [m for m in REQUIRED_MODULES if not mods[m]]
    python_ok = sys.version_info >= (3, 10)
    ready_core = python_ok and (bool(uv_path) or not missing_req)

    hints: list[str] = []
    if not uv_path and missing_req:
        hints.append(f"Install uv ({cmds['uv']}) OR install requirements.txt into a venv.")
    if not soffice_path:
        hints.append(f"Optional (needed for formula recalc & PDF/PNG rendering): {cmds['libreoffice']}")
    if not pdftoppm_path:
        hints.append(f"Optional (needed for PNG visual layout inspection): {cmds['pdftoppm']}")

    report = {
        "os": sys_name,
        "python": platform.python_version(),
        "python_ok": python_ok,
        "uv": uv_path,
        "modules": mods,
        "libreoffice": soffice_path,
        "pdftoppm": pdftoppm_path,
        "ready_core": ready_core,
        "ready_full": ready_core and bool(soffice_path) and bool(pdftoppm_path),
        "hints": hints,
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"OS: {sys_name} | Python: {report['python']} ({'OK' if python_ok else 'need >=3.10'})")
        print(f"uv: {uv_path or 'not found'}")
        print(
            "Python modules (current interpreter): "
            + ", ".join(f"{k}={'ok' if v else 'missing'}" for k, v in mods.items())
        )
        print(f"LibreOffice: {soffice_path or 'not found'}")
        print(f"pdftoppm: {pdftoppm_path or 'not found'}")
        print(f"Core read/write ready: {ready_core} | Full (recalc + render) ready: {report['ready_full']}")
        for h in hints:
            print(f"Hint: {h}")

    if not ready_core:
        sys.exit(1)


if __name__ == "__main__":
    main()
