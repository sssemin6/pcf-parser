from contextlib import closing, redirect_stdout
from io import BytesIO, StringIO
import json
from pathlib import Path
import sqlite3
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
from pcf_parser.parser import parse_document, search_text


_HEADERS = ("No.", "Buyer’s Comments", "Builder's Reply", "Status")
_SPLIT_HEADERS = ("No.", "BUYER'S COMMENTS", "", "BUILDER'S REPLY", "", "Initial", "Status")


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


def _write_split_excel(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    for row in (
        ["Hull No. 3395-3401"],
        ["174,000 CBM LNG CARRIER"],
        ["Plan No.: 6U-2853-303", "Plan Title: Pump arrangement"],
        ["Approved", "", "Approved with comments", "✓"],
        _SPLIT_HEADERS,
        ["1-a", "1st", "Check pump capacity", "1st", "Capacity is 500 m3/h", "Alex", "Open"],
        [None, "2nd", "Review diagram\nAttachment: sketch.pdf", "2nd", "Noted, thanks", None, "Closed"],
    ):
        sheet.append(row)
    sheet.merge_cells("B5:C5")
    sheet.merge_cells("D5:E5")
    sheet.add_image(ExcelImage(_picture()), "C7")
    workbook.save(path)
    workbook.close()


def _write_split_word(path: Path) -> None:
    document = make_word()
    table = document.add_table(rows=0, cols=7)
    for row in (
        ("Hull No. 3228/3229\n174,000 CBM LNG CARRIER", "Plan Title: Fan arrangement", "", "", "", "", ""),
        ("Not approved", "X", "", "", "", "", ""),
        _SPLIT_HEADERS,
        ("1", "a)", "Verify fan speed", "a)", "Speed confirmed", "QELNG/DK", "Open"),
        ("", "b)", "Check drawing", "b)", "", "", "Closed"),
    ):
        for cell, value in zip(table.add_row().cells, row):
            cell.text = value
    table.cell(2, 1).merge(table.cell(2, 2))
    table.cell(2, 3).merge(table.cell(2, 4))
    table.cell(3, 5).merge(table.cell(4, 5))
    table.cell(4, 4).paragraphs[0].add_run().add_picture(_picture())
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
        workbook.active.append([1, "Review", "See attached\nAttachment: plan.pdf"])
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
        self.assertEqual(turn.attachments, ("plan.pdf", "oleObject1.bin"))
        self.assertIn("[FILE_ATTACHED: plan.pdf]", turn.builder_reply)
        self.assertIn("[FILE_ATTACHED: oleObject1.bin]", turn.builder_reply)

    def test_merge_ranges_with_separate_round_and_content_columns(self):
        _write_split_excel(self.xlsx)
        _write_split_word(self.docx)
        for path, labels, initial in (
            (self.xlsx, ["1st", "2nd"], "Alex"),
            (self.docx, ["a)", "b)"], "QELNG/DK"),
        ):
            with self.subTest(path=path):
                result = parse_document(path, include_preview=True)
                thread, = result.threads
                self.assertEqual(thread.reviewer_initial, initial)
                self.assertEqual([turn.round_label for turn in thread.turns], labels)
                self.assertEqual(thread.raw_status, "Closed")
                self.assertNotEqual(thread.turns[0].buyer_comment, labels[0])
                self.assertTrue(thread.turns[1].has_image)
                cells = [cell for table in result.raw_preview for row in table["rows"] for cell in row["cells"]]
                self.assertIn({"columns": [2, 3], "text": "BUYER'S COMMENTS", "has_image": False,
                               "attachments": []}, cells)
        excel_thread = parse_document(self.xlsx).threads[0]
        self.assertEqual(excel_thread.turns[0].buyer_comment, "Check pump capacity")
        self.assertEqual(excel_thread.turns[0].builder_reply, "Capacity is 500 m3/h")
        self.assertEqual(excel_thread.turns[1].attachments, ("sketch.pdf",))
        self.assertEqual(parse_document(self.docx).threads[0].turns[1].builder_reply, "[IMAGE_ATTACHED]")

    def test_inline_hull_capacity_description_and_checked_approval(self):
        _write_split_excel(self.xlsx)
        _write_split_word(self.docx)
        for path, hull, approval in (
            (self.xlsx, "3395-3401", "Approved with comments"),
            (self.docx, "3228/3229", "Not approved"),
        ):
            with self.subTest(path=path):
                values = parse_document(path).metadata
                self.assertEqual(values["hull_no"], hull)
                self.assertEqual(values["ship_type"], "174,000 CBM LNG CARRIER")
                self.assertEqual(values["approval_status"], approval)

    def test_search_text_filters_only_standalone_admin_lines(self):
        _write_split_excel(self.xlsx)
        thread = parse_document(self.xlsx).threads[0]
        text = ("Waiting for buyer's confirmation.\nCan Builder understand that the comment is closed?\n"
                "Noted, thanks\nThank you. Confirmed / Agreed\nWaiting for confirmation of pump pressure.")
        thread.turns[1].builder_reply = text
        indexed = search_text(thread)
        self.assertIn("Waiting for confirmation of pump pressure.", indexed)
        for phrase in text.splitlines()[:-1]:
            self.assertNotIn(phrase, indexed)
        self.assertEqual(thread.turns[1].builder_reply, text)

    def test_hyphen_search_and_separate_initial_columns(self):
        _write_split_excel(self.xlsx)
        with closing(database.connect(self.db_path)) as db:
            ingest(db, self.root)
            for query in ("6U-2853-303", "3395-3401", "1-a", '"6U-2853-303"'):
                with self.subTest(query=query):
                    found, = database.search(db, query, 5)
                    self.assertEqual(found["reviewer_initial"], "Alex")
                    self.assertEqual(found["initial_comment"], "Check pump capacity")
                    self.assertEqual(found["plan_no"], "6U-2853-303")
            self.assertEqual(database.search(db, 'pump:capacity', 5)[0]["comment_no"], "1-a")

    def test_legacy_database_initial_is_preserved_until_force_ingest(self):
        _write_split_excel(self.xlsx)
        with closing(sqlite3.connect(self.db_path)) as old:
            old.executescript(database._SCHEMA.replace("initial_comment TEXT,", "initial TEXT,")
                              .replace("    reviewer_initial TEXT,\n", ""))
            with old:
                old.execute("INSERT INTO pcf_documents VALUES (1, ?, 'old', 'xlsx', NULL, NULL, 'Test',"
                            " NULL, NULL, NULL, 'rule', CURRENT_TIMESTAMP)", (str(self.xlsx),))
                old.execute("INSERT INTO pcf_threads (doc_id, comment_no, source, initial, search_text)"
                            " VALUES (1, '1', 'PCF', 'Old initial', 'old initial')")
                old.execute("INSERT INTO pcf_turns (thread_id, turn_order, has_image, attachments)"
                            " VALUES (1, 1, 0, '[]')")
        with closing(database.connect(self.db_path, create=False)) as old:
            self.assertEqual(database.search(old, "old", 2)[0]["initial_comment"], "Old initial")
        with closing(database.connect(self.db_path)) as db:
            self.assertEqual(tuple(db.execute("SELECT initial_comment, reviewer_initial FROM pcf_threads").fetchone()),
                             ("Old initial", None))
            self.assertEqual(ingest(db, self.root).parsed, 1)
            self.assertEqual(ingest(db, self.root).unchanged, 1)
            self.assertEqual(ingest(db, self.root, force=True).parsed, 1)

    def test_cli_parse_review_and_pipeline_with_mock_gateway(self):
        _write_split_excel(self.xlsx)
        nested = self.root / "sub"
        nested.mkdir()
        _write_split_word(nested / self.docx.name)
        out = self.root / "work"
        parsed = out / "parsed"
        feedback = out / "feedback"
        env = {
            "MAILSEARCH_GATEWAY_URL": "https://review.example.test:8443/v1",
            "MAILSEARCH_GATEWAY_API_KEY": "test-key",
            "MAILSEARCH_GATEWAY_ANSWER_MODELS": " review-model,other-model",
        }
        result = {
            "verdict": "WARN",
            "metadata_issues": [],
            "thread_issues": [{"pattern": "병합 매칭", "detail": "라운드 위치 근거를 다시 확인"}],
            "parser_fix_suggestions": ["parser.py의 병합 범위 회귀 테스트 추가"],
        }
        fenced = "\x60\x60\x60json\n" + json.dumps(result, ensure_ascii=False) + "\n\x60\x60\x60"
        response = json.dumps({"choices": [{"message": {"content": fenced}}]}).encode()
        with (patch.dict("os.environ", env, clear=True),
              patch("pcf_parser.llm.http.client.HTTPSConnection") as https,
              redirect_stdout(StringIO())):
            connection = https.return_value
            connection.getresponse.return_value.status = 200
            connection.getresponse.return_value.read.return_value = response
            self.assertEqual(main(["parse", str(self.root), "--out", str(parsed)]), 0)
            self.assertEqual(main(["parse", str(self.root), "--out", str(parsed)]), 0)
            self.assertEqual(json.loads((parsed / "_parse_report.json").read_text(encoding="utf-8"))["unchanged"], 2)
            self.assertEqual(main(["review", str(parsed), "--out", str(feedback)]), 0)
            self.assertEqual(main(["pipeline", str(self.root), "--out", str(out), "--force"]), 0)
            connection.request.assert_called()
            request = connection.request.call_args
            self.assertEqual(request.args, ("POST", "/v1/chat/completions"))
            body = json.loads(request.kwargs["body"])
            self.assertEqual((body["model"], body["temperature"], body["stream"]), ("review-model", 0, False))
            prompt = json.loads(body["messages"][1]["content"])
            self.assertEqual(set(prompt), {"metadata", "threads", "warnings", "raw_preview"})
            self.assertTrue(prompt["raw_preview"])
            self.assertEqual(connection.request.call_count, 4)
        data = json.loads((parsed / "example.xlsx.json").read_text(encoding="utf-8"))
        self.assertEqual((len(data["file_hash"]), data["file_type"]), (64, "xlsx"))
        self.assertTrue(data["parsed_at"])
        self.assertEqual(data["threads"][0]["reviewer_initial"], "Alex")
        self.assertEqual(data["threads"][0]["turns"][0]["turn_order"], 1)
        report = json.loads((feedback / "feedback_report.json").read_text(encoding="utf-8"))
        self.assertEqual(report["counts"]["WARN"], 2)
        summary = (feedback / "feedback_summary.md").read_text(encoding="utf-8")
        self.assertIn("병합 매칭 (2개 파일)", summary)
        self.assertIn("sub/example.docx.json", summary)
        self.assertIn("parser.py의 병합 범위 회귀 테스트 추가", summary)

    def test_pipeline_records_gateway_failure_and_keeps_unmapped_preview(self):
        workbook = Workbook()
        workbook.active.append(["Unrecognized", "Comments"])
        workbook.save(self.xlsx)
        workbook.close()
        output = self.root / "work"
        with patch.dict("os.environ", {}, clear=True), redirect_stdout(StringIO()):
            self.assertEqual(main(["pipeline", str(self.root), "--out", str(output)]), 1)
        parsed = json.loads((output / "parsed" / "example.xlsx.json").read_text(encoding="utf-8"))
        self.assertEqual(parsed["threads"], [])
        self.assertTrue(parsed["raw_preview"])
        self.assertTrue(parsed["warnings"])
        summary = (output / "feedback" / "feedback_summary.md").read_text(encoding="utf-8")
        self.assertIn("검토 실행 오류", summary)
        self.assertIn("PCF_LLM_URL", summary)

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
