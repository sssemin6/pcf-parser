from contextlib import closing, redirect_stdout
from io import BytesIO, StringIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document as make_word
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image

from pcf_parser import database
from pcf_parser.cli import main
from pcf_parser.ingest import ingest
from pcf_parser.parser import parse_document


_HEADERS = ("No.", "Buyer’s Comments", "Builder's Reply", "Status")


def _picture() -> BytesIO:
    data = BytesIO()
    Image.new("RGB", (2, 2), "red").save(data, format="PNG")
    data.seek(0)
    return data


def _write_excel(path: Path) -> None:
    workbook = Workbook()
    cover = workbook.active
    cover.title = "Cover"
    cover.append(["Plan No.", "PL-0042"])
    sheet = workbook.create_sheet("PCF")
    sheet.append(["Hull No.", None, "H-101"])
    sheet.merge_cells("A1:B1")
    sheet.append(["Plan Title: Pump arrangement", None, "Approval Status", "Approved with comments"])
    sheet.append(_HEADERS)
    sheet.append(["1", "1st: Verify pump capacity", "1st: Measurement scheduled"])
    sheet.append([None, "2nd: Check report\nAttachment #2: sheet.pdf", "2nd: Report provided"])
    sheet.merge_cells("A4:A5")
    sheet.add_image(ExcelImage(_picture()), "B5")
    sheet.append(_HEADERS)
    sheet.append([None, None, "3rd: Builder looks forward to confirmation"])
    sheet.append([None, "4th: Acknowledged", "4th: Confirmed"])
    sheet.append([None, "5th: Accepted", "5th: Closed", "Closed"])
    sheet.append(["2", "Verify fan speed", "Confirmed"])
    second = workbook.create_sheet("Other title")
    second.append(_HEADERS)
    second.append(["1", "Confirm spare parts", "Confirmed"])
    workbook.save(path)
    workbook.close()


def _write_word(path: Path) -> None:
    document = make_word()
    metadata = document.add_table(rows=2, cols=3)
    metadata.cell(0, 0).merge(metadata.cell(0, 1)).text = "Hull No."
    metadata.cell(0, 2).text = "H-205"
    metadata.cell(1, 0).text = "Plan Title"
    metadata.cell(1, 1).merge(metadata.cell(1, 2)).text = "Fan arrangement"
    body = document.add_table(rows=0, cols=4)
    rows = (
        ("Comment No.", "Class Comments", "Yard Reply", "Remarks"),
        ("3", "a) Confirm airflow\nb) Inspect fan", "a) Verified", "Open"),
        ("", "", "b) Inspection complete", "Closed"),
        ("Comment No.", "Class Comments", "Yard Reply", "Remarks"),
        ("", "c) Accepted", "c) Confirmed", "Closed"),
    )
    for row in rows:
        for cell, value in zip(body.add_row().cells, row):
            cell.text = value
    body.cell(1, 0).merge(body.cell(2, 0))
    body.cell(1, 1).merge(body.cell(2, 1))
    body.cell(2, 2).paragraphs[0].add_run().add_picture(_picture())
    document.save(path)


class ParserTest(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.xlsx = self.root / "example.xlsx"
        self.docx = self.root / "example.docx"
        self.db_path = self.root / "index.sqlite3"

    def test_excel_merges_repeated_headers_and_multiple_sheets(self):
        _write_excel(self.xlsx)
        document = parse_document(self.xlsx)
        self.assertEqual(document.metadata, {
            "hull_no": "H-101", "plan_title": "Pump arrangement", "plan_no": "PL-0042",
            "comment_party": "Buyer", "approval_status": "Approved with comments",
        })
        self.assertEqual([(t.source, t.comment_no) for t in document.threads], [
            ("PCF", "1"), ("PCF", "2"), ("Other title", "1"),
        ])
        first = document.threads[0]
        self.assertEqual(first.raw_status, "Closed")
        self.assertEqual([t.round_label for t in first.turns], ["1st", "2nd", "3rd", "4th", "5th"])
        self.assertIn("[IMAGE_ATTACHED]", first.turns[1].buyer_comment)
        self.assertIn("[FILE_ATTACHED: sheet.pdf]", first.turns[1].buyer_comment)
        self.assertEqual(first.turns[1].attachments, ("sheet.pdf",))

    def test_word_merges_and_pairs_replies_by_round_across_rows(self):
        _write_word(self.docx)
        document = parse_document(self.docx)
        self.assertEqual(document.metadata["hull_no"], "H-205")
        self.assertEqual(document.metadata["plan_title"], "Fan arrangement")
        self.assertEqual(document.metadata["comment_party"], "Class")
        thread, = document.threads
        self.assertEqual([t.round_label for t in thread.turns], ["a)", "b)", "c)"])
        self.assertEqual(thread.turns[1].buyer_comment, "Inspect fan")
        self.assertIn("Inspection complete", thread.turns[1].builder_reply)
        self.assertTrue(thread.turns[1].has_image)

    def test_excel_embedded_attachment_is_tagged_at_its_anchor(self):
        source = self.root / "source.xlsx"
        workbook = Workbook()
        workbook.active.append(_HEADERS)
        workbook.active.append([1, "Review", "See attached"])
        workbook.save(source)
        workbook.close()
        office_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        drawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        ole_object = (
            f'<oleObjects><oleObject xmlns:r="{office_rel}" r:id="rId1">'
            f'<objectPr><anchor><from xmlns:xdr="{drawing}"><xdr:row>1</xdr:row><xdr:col>2</xdr:col>'
            '</from></anchor></objectPr></oleObject></oleObjects>'
        )
        relationships = (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{office_rel}/oleObject" '
            'Target="../embeddings/oleObject1.bin"/></Relationships>'
        )
        with ZipFile(source) as original, ZipFile(self.xlsx, "w") as modified:
            for name in original.namelist():
                data = original.read(name)
                if name == "xl/worksheets/sheet1.xml":
                    data = data.replace(b"</worksheet>", (ole_object + "</worksheet>").encode())
                modified.writestr(name, data)
            modified.writestr("xl/worksheets/_rels/sheet1.xml.rels", relationships)
            modified.writestr("xl/embeddings/oleObject1.bin", b"synthetic")

        turn = parse_document(self.xlsx).threads[0].turns[0]
        self.assertEqual(turn.attachments, ("oleObject1.bin",))
        self.assertEqual(turn.builder_reply, "See attached\n[FILE_ATTACHED: oleObject1.bin]")

    def test_number_and_metadata_fallback_without_guessing(self):
        path = self.root / "Hull_No_H104_Plan_No_D-002.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Round", "Owner Comment", "Remark", "Item", "Yard Response"])
        sheet.append(["2nd", "Second request", None, "A-2", "Second reply"])
        sheet.append(["1st", "First request", None, None, "First reply"])
        workbook.save(path)
        workbook.close()
        result = parse_document(path)
        self.assertEqual(result.metadata["hull_no"], "H104")
        self.assertEqual(result.metadata["plan_no"], "D-002")
        self.assertNotIn("plan_title", result.metadata)
        self.assertIn("Missing metadata: plan_title", result.warnings)
        thread, = result.threads
        self.assertEqual(thread.comment_no, "A-2")
        self.assertEqual([turn.buyer_comment for turn in thread.turns], ["First request", "Second request"])

    def test_unmatched_round_preserves_original_speakers(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(_HEADERS)
        sheet.append([1, "Lead text\na) First request", "a) First reply\nb) Second reply"])
        workbook.save(self.xlsx)
        workbook.close()
        result = parse_document(self.xlsx)
        thread, = result.threads
        self.assertEqual(thread.turns[1].buyer_comment, "First request")
        self.assertEqual(thread.turns[1].builder_reply, "First reply")
        self.assertEqual(thread.turns[0].buyer_comment, "Lead text")
        self.assertEqual(thread.turns[0].builder_reply, "")
        self.assertEqual(thread.turns[2].buyer_comment, "")
        self.assertEqual(thread.turns[2].builder_reply, "Second reply")
        self.assertIn("ambiguous rounds", result.warnings[0])

    def test_incremental_index_retains_threads_and_cleans_stale_matches(self):
        _write_excel(self.xlsx)
        _write_word(self.docx)
        with closing(database.connect(self.db_path)) as db:
            report = ingest(db, self.root)
            self.assertEqual((report.parsed, report.threads, report.turns), (2, 4, 10))
            self.assertEqual(ingest(db, self.root).unchanged, 2)
            found, = database.search(db, '"pump capacity"', 20)
            self.assertEqual(len(found["turns"]), 5)
            self.assertIn("Builder looks forward", found["turns"][2]["builder_reply"])
            self.assertEqual(database.search(db, '"looks forward"', 20), [])
            self.assertEqual(len(database.search(db, '"pump arrangement"', 20)), 3)
            original_bytes = self.xlsx.read_bytes()
            self.xlsx.write_text("bad workbook")
            self.assertEqual(len(ingest(db, self.root).failed), 1)
            self.assertEqual(len(database.search(db, '"pump capacity"', 20)), 1)
            self.xlsx.write_bytes(original_bytes)
            workbook = load_workbook(self.xlsx)
            workbook["PCF"]["B4"] = "1st: Check pressure"
            workbook.save(self.xlsx)
            workbook.close()
            self.assertEqual(ingest(db, self.root).parsed, 1)
            self.assertEqual(database.search(db, '"pump capacity"', 20), [])
            self.assertEqual(len(database.search(db, '"check pressure"', 20)[0]["turns"]), 5)

    def test_scan_logs_supported_skip_reasons(self):
        (self.root / "~$locked.xlsx").touch()
        (self.root / "scan.pdf").touch()
        (self.root / "old.xls").touch()
        encrypted = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + "EncryptedPackage".encode("utf-16-le")
        self.xlsx.write_bytes(encrypted)
        with patch("pcf_parser.ingest.shutil.which", return_value=None):
            with closing(database.connect(self.db_path)) as db:
                report = ingest(db, self.root)
                reasons = {row["reason"] for row in db.execute("SELECT reason FROM skipped_files")}
        self.assertEqual((report.total_files, report.skipped, report.failed), (4, 4, {}))
        self.assertEqual(reasons, {"office_temporary_file", "unsupported_pdf", "encrypted_office",
                                   "legacy_converter_unavailable"})

    def test_cli_ingest_then_search(self):
        _write_excel(self.xlsx)
        output = StringIO()
        with redirect_stdout(output):
            self.assertEqual(main(["ingest", str(self.root), "--db", str(self.db_path)]), 0)
        self.assertEqual(json.loads(output.getvalue())["parsed"], 1)
        output = StringIO()
        with redirect_stdout(output):
            self.assertEqual(main(["search", '"pump capacity"', "--db", str(self.db_path)]), 0)
        self.assertEqual(len(json.loads(output.getvalue())[0]["turns"]), 5)
