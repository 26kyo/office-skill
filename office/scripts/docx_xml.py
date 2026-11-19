#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "lxml>=5.0",
# ]
# ///
"""Unpack, validate, and pack .docx OOXML archives for safe XML-level edits."""

import argparse
import sys
import zipfile
from pathlib import Path

from lxml import etree


def cmd_unpack(docx_path: Path, out_dir: Path, pretty: bool = False) -> None:
    if out_dir.exists() and any(out_dir.iterdir()):
        sys.exit(f"Output directory is not empty: {out_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)
    out_root = out_dir.resolve()
    with zipfile.ZipFile(docx_path, "r") as zf:
        for info in zf.infolist():
            target = (out_root / info.filename).resolve()
            if not target.is_relative_to(out_root):
                sys.exit(f"Unsafe zip entry: {info.filename}")
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            data = zf.read(info.filename)
            if pretty and info.filename.endswith((".xml", ".rels")):
                root = etree.fromstring(data)
                data = etree.tostring(root, encoding="utf-8", xml_declaration=True, pretty_print=True)
            target.write_bytes(data)
    print(f"unpacked: {docx_path} -> {out_dir}")


def validate_dir(src_dir: Path) -> list[Path]:
    if not (src_dir / "[Content_Types].xml").is_file():
        sys.exit(f"Missing [Content_Types].xml in {src_dir}")
    if not (src_dir / "word" / "document.xml").is_file():
        sys.exit(f"Missing word/document.xml in {src_dir}")
    files = sorted(p for p in src_dir.rglob("*") if p.is_file())
    for p in files:
        if p.suffix in (".xml", ".rels"):
            try:
                etree.parse(str(p))
            except etree.XMLSyntaxError as exc:
                sys.exit(f"XML syntax error in {p.relative_to(src_dir)}: {exc}")
    return files


def cmd_pack(src_dir: Path, out_docx: Path) -> None:
    files = validate_dir(src_dir)
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_docx, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in files:
            arcname = p.relative_to(src_dir).as_posix()
            zf.write(p, arcname)
    print(f"packed: {src_dir} -> {out_docx} ({len(files)} entries)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Safe OOXML unpack/validate/pack helper for .docx")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_unpack = sub.add_parser("unpack")
    p_unpack.add_argument("docx", type=Path)
    p_unpack.add_argument("out_dir", type=Path)
    p_unpack.add_argument("--pretty", action="store_true", help="Pretty-print .xml/.rels files on unpack")

    p_val = sub.add_parser("validate")
    p_val.add_argument("src_dir", type=Path)

    p_pack = sub.add_parser("pack")
    p_pack.add_argument("src_dir", type=Path)
    p_pack.add_argument("out_docx", type=Path)

    args = parser.parse_args()
    if args.cmd == "unpack":
        cmd_unpack(args.docx, args.out_dir, pretty=args.pretty)
    elif args.cmd == "validate":
        files = validate_dir(args.src_dir)
        print(f"valid: {args.src_dir} ({len(files)} entries)")
    elif args.cmd == "pack":
        cmd_pack(args.src_dir, args.out_docx)


if __name__ == "__main__":
    main()
