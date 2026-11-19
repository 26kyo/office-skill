import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
import openpyxl
from pptx import Presentation
from pptx.util import Inches

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = REPO_ROOT / "office" / "scripts"


def run_script(script_name: str, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPTS_DIR / script_name), *args],
        capture_output=True,
        text=True,
    )


class OfficeSkillSmokeTests(unittest.TestCase):
    def test_doctor_runs(self) -> None:
        res = run_script("doctor.py", "--json")
        self.assertEqual(res.returncode, 0, res.stderr)
        data = json.loads(res.stdout)
        self.assertTrue(data["ready_core"])

    def test_xlsx_read_and_scan(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx = Path(tmp) / "sample.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Budget"
            ws.append(["Item", "Cost"])
            ws.append(["A", 10])
            ws.append(["B", 20])
            ws.append(["Total", "=SUM(B2:B3)"])
            wb.save(xlsx)

            res = run_script("office_read.py", "xlsx", str(xlsx))
            self.assertEqual(res.returncode, 0, res.stderr)
            self.assertIn("Budget", res.stdout)

            res = run_script("office_read.py", "xlsx", str(xlsx), "--sheet", "Budget")
            self.assertEqual(res.returncode, 0, res.stderr)
            self.assertIn("B4: =SUM(B2:B3)", res.stdout)

            if shutil.which("libreoffice") or shutil.which("soffice"):
                res = run_script("xlsx_recalc.py", str(xlsx), "--backup")
                self.assertEqual(res.returncode, 0, res.stderr)
                report = json.loads(res.stdout)
                self.assertEqual(report["status"], "success")
                self.assertEqual(report["total_formulas"], 1)
                self.assertTrue((Path(tmp) / "sample.xlsx.bak").is_file())

                res = run_script("office_read.py", "xlsx", str(xlsx), "--sheet", "Budget")
                self.assertIn("B4: =SUM(B2:B3) => 30", res.stdout)

    def test_xlsx_recalc_catches_error(self) -> None:
        if not (shutil.which("libreoffice") or shutil.which("soffice")):
            self.skipTest("LibreOffice not installed")
        with tempfile.TemporaryDirectory() as tmp:
            xlsx = Path(tmp) / "bad.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = "=1/0"
            wb.save(xlsx)

            res = run_script("xlsx_recalc.py", str(xlsx), "--no-write")
            self.assertNotEqual(res.returncode, 0)
            report = json.loads(res.stdout)
            self.assertEqual(report["status"], "errors_found")
            self.assertIn("#DIV/0!", report["error_summary"])

    def test_docx_read_and_xml_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            docx_in = tmp_path / "in.docx"
            doc = Document()
            doc.add_heading("Report Title", level=1)
            doc.add_paragraph("Original body text.")
            tbl = doc.add_table(rows=2, cols=2)
            tbl.cell(0, 0).text = "Col1"
            tbl.cell(0, 1).text = "Col2"
            tbl.cell(1, 0).text = "v1"
            tbl.cell(1, 1).text = "v2"
            doc.save(docx_in)

            res = run_script("office_read.py", "docx", str(docx_in), "--outline")
            self.assertEqual(res.returncode, 0, res.stderr)
            self.assertIn("Report Title", res.stdout)
            self.assertIn("[Table 1: 2x2]", res.stdout)

            work = tmp_path / "unpacked"
            res = run_script("docx_xml.py", "unpack", str(docx_in), str(work))
            self.assertEqual(res.returncode, 0, res.stderr)

            doc_xml = work / "word" / "document.xml"
            doc_xml.write_text(doc_xml.read_text(encoding="utf-8").replace("Original body text.", "Updated text."), encoding="utf-8")

            docx_out = tmp_path / "out.docx"
            res = run_script("docx_xml.py", "pack", str(work), str(docx_out))
            self.assertEqual(res.returncode, 0, res.stderr)

            res = run_script("office_read.py", "docx", str(docx_out))
            self.assertEqual(res.returncode, 0, res.stderr)
            self.assertIn("Updated text.", res.stdout)

    def test_docx_xml_rejects_zip_slip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            evil = tmp_path / "evil.docx"
            with zipfile.ZipFile(evil, "w") as zf:
                zf.writestr("../escape.txt", "boom")
            out_dir = tmp_path / "out"
            res = run_script("docx_xml.py", "unpack", str(evil), str(out_dir))
            self.assertNotEqual(res.returncode, 0)
            self.assertIn("Unsafe zip entry", res.stderr)

    def test_pptx_reads_grouped_shapes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            pptx_path = Path(tmp) / "deck.pptx"
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            group = slide.shapes.add_group_shape()
            tb = group.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
            tb.text_frame.text = "Inside group shape"
            prs.save(pptx_path)

            res = run_script("office_read.py", "pptx", str(pptx_path), "--slides", "1")
            self.assertEqual(res.returncode, 0, res.stderr)
            self.assertIn("Inside group shape", res.stdout)


if __name__ == "__main__":
    unittest.main()
