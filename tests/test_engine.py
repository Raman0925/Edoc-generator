"""Exercise real PDF generation/merging with an injected fake Excel adapter.

Fixture .xlsx paths contain JSON solely for the fake adapter, not real workbooks.
Microsoft Excel integration is tested separately on an Office-equipped Windows PC.
"""

import json
import tempfile
import threading
import unittest
from contextlib import contextmanager
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from edoc.config import Mapping, Profile, ValidationError
from edoc.engine import discover_reports, match_certificate, merge_pdfs, preview, run_batch, safe_name, sha256


def pdf(path, width=600, pages=1, encrypted=False):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=width, height=800)
    if encrypted:
        writer.encrypt("secret")
    writer.write(str(path))
    writer.close()


class FakeExcel:
    def __enter__(self):
        return self

    def __exit__(self, *_):
        pass

    @contextmanager
    def open(self, path, writable=False):
        data = json.loads(path.read_text(encoding="utf-8"))
        data["_path"] = path
        yield data

    def sheets(self, book):
        return ["Approval"] if book.get("template") else ["Report"]

    def read(self, book, mapping):
        if mapping.source_sheet not in self.sheets(book):
            raise ValidationError("Missing worksheet")
        return book.get(mapping.source_cell)

    def check_target(self, book, mapping):
        if mapping.target_sheet != "Approval":
            raise ValidationError("Missing target worksheet")
        return "Approval!" + ("F6" if book.get("merged") else mapping.target_cell)

    def write(self, book, mapping, value):
        book[mapping.target_cell] = value

    def calculate(self, book):
        pass

    def save(self, book):
        book["_path"].write_text(json.dumps({k: v for k, v in book.items() if k != "_path"}), encoding="utf-8")

    def export_pdf(self, book, destination, sheet=""):
        pdf(destination, width=500 if book.get("template") else 600)


class EngineTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.input = self.root / "input"
        self.input.mkdir()
        self.output = self.root / "output"
        self.template = self.root / "template.xlsx"
        self.template.write_text(json.dumps({"template": True}), encoding="utf-8")
        self.profile = Profile(template=str(self.template), input_dir=str(self.input), output_dir=str(self.output),
                               mappings=[Mapping("Serial", "Report", "B4", "Approval", "F6", "text")])

    def tearDown(self):
        self.temp.cleanup()

    def report(self, name, value="ACC-01"):
        path = self.input / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"B4": value}), encoding="utf-8")
        return path

    def test_five_reports_produce_five_complete_documents(self):
        sources = [self.report(f"report-{i}.xlsx", f"ACC-{i}") for i in range(5)]
        original_hashes = [sha256(path) for path in [self.template, *sources]]
        events = []
        result = run_batch(self.profile, events.append, session_factory=FakeExcel)
        self.assertEqual(result.succeeded, 5)
        self.assertEqual(len(list(Path(result.run_dir).glob("*/final.pdf"))), 5)
        self.assertEqual([sha256(path) for path in [self.template, *sources]], original_hashes)
        for i, item in enumerate(result.results):
            final = Path(item.output)
            reader = PdfReader(str(final))
            self.assertEqual([int(page.mediabox.width) for page in reader.pages], [500, 600])
            audit = json.loads((final.parent / "audit.json").read_text())
            self.assertEqual(audit["values"][0]["value"], f"ACC-{i}")
            self.assertEqual(audit["final_sha256"], sha256(final))
            filled = json.loads((final.parent / "approval.xlsx").read_text())
            self.assertEqual(filled["F6"], f"ACC-{i}")
            self.assertFalse((final.parent / "source.xlsx").exists())
        self.assertEqual(len([event for event in events if event["kind"] == "result"]), 5)

    def test_blank_failure_does_not_stop_other_reports(self):
        self.report("a.xlsx", None)
        self.report("b.xlsx", "ok")
        result = run_batch(self.profile, session_factory=FakeExcel)
        self.assertEqual([item.status for item in result.results], ["failed", "success"])
        self.assertIn("empty", result.results[0].error)
        self.assertFalse(list(Path(result.run_dir).glob(".pending-*")))

    def test_nan_is_rejected(self):
        self.report("a.xlsx", float("nan"))
        result = run_batch(self.profile, session_factory=FakeExcel)
        self.assertEqual(result.results[0].status, "failed")
        self.assertIn("finite", result.results[0].error)

    def test_zero_is_a_valid_required_value(self):
        self.report("zero.xlsx", 0)
        self.assertEqual(run_batch(self.profile, session_factory=FakeExcel).succeeded, 1)

    def test_optional_blank_is_allowed(self):
        self.report("a.xlsx", None)
        self.profile.mappings[0].required = False
        self.assertEqual(run_batch(self.profile, session_factory=FakeExcel).succeeded, 1)

    def test_same_stem_different_extensions_no_collision(self):
        self.report("a.xlsx")
        self.report("a.xlsm")
        result = run_batch(self.profile, session_factory=FakeExcel)
        self.assertEqual(result.succeeded, 2)
        self.assertNotEqual(result.results[0].output, result.results[1].output)

    def test_new_run_never_overwrites_previous_run(self):
        self.report("a.xlsx")
        first = run_batch(self.profile, session_factory=FakeExcel)
        second = run_batch(self.profile, session_factory=FakeExcel)
        self.assertNotEqual(first.run_dir, second.run_dir)
        self.assertTrue(Path(first.results[0].output).exists())

    def test_existing_certificate_attached_in_correct_order(self):
        self.report("acc.xlsx")
        certs = self.root / "certs"
        certs.mkdir()
        pdf(certs / "acc.PDF", width=750, pages=2)
        self.profile.certificate_mode = "pdf"
        self.profile.certificate_dir = str(certs)
        result = run_batch(self.profile, session_factory=FakeExcel)
        self.assertEqual(result.succeeded, 1)
        self.assertEqual([int(page.mediabox.width) for page in PdfReader(result.results[0].output).pages], [500, 750, 750])

    def test_missing_certificate_fails_without_publishing_partial_output(self):
        self.report("acc.xlsx")
        self.profile.certificate_mode = "pdf"
        self.profile.certificate_dir = str(self.root)
        result = run_batch(self.profile, session_factory=FakeExcel)
        self.assertEqual(result.succeeded, 0)
        self.assertIn("found 0", result.results[0].error)
        self.assertFalse(list(Path(result.run_dir).glob("*/final.pdf")))

    def test_approval_only_and_optional_components(self):
        self.report("a.xlsx")
        self.profile.certificate_mode = "approval_only"
        self.profile.keep_components = False
        result = run_batch(self.profile, session_factory=FakeExcel)
        final = Path(result.results[0].output)
        self.assertEqual(len(PdfReader(str(final)).pages), 1)
        self.assertFalse((final.parent / "approval.pdf").exists())
        self.assertTrue((final.parent / "approval.xlsx").exists())

    def test_preview_changes_no_file_and_creates_no_output(self):
        sample = self.report("a.xlsx", "00012")
        original = sha256(sample)
        result = preview(self.profile, sample, session_factory=FakeExcel)
        self.assertEqual(result["values"][0]["value"], "00012")
        self.assertEqual(sha256(sample), original)
        self.assertFalse(self.output.exists())

    def test_merged_destination_collision_detected(self):
        self.report("a.xlsx")
        self.template.write_text(json.dumps({"template": True, "merged": True}))
        self.profile.mappings.append(Mapping("Other", "Report", "B4", "Approval", "G6"))
        result = run_batch(self.profile, session_factory=FakeExcel)
        self.assertEqual(result.succeeded, 0)
        self.assertIn("merged cell", result.results[0].error)

    def test_cancel_keeps_finished_report(self):
        self.report("a.xlsx")
        self.report("b.xlsx")
        event = threading.Event()
        def progress(item):
            if item["kind"] == "result":
                event.set()
        result = run_batch(self.profile, progress, event, FakeExcel)
        self.assertEqual(result.succeeded, 1)
        self.assertTrue(result.cancelled)
        manifest = json.loads((Path(result.run_dir) / "batch.json").read_text())
        self.assertEqual(manifest["not_processed"], 1)

    def test_empty_input_creates_no_run(self):
        with self.assertRaisesRegex(ValidationError, "No Excel reports"):
            run_batch(self.profile, session_factory=FakeExcel)
        self.assertFalse(self.output.exists())

    def test_discovery_excludes_lock_files_template_and_non_excel(self):
        self.report("a.xlsx")
        self.report("~$a.xlsx")
        self.report("ignored.txt")
        self.report("nested/b.xlsx")
        self.profile.template = str(self.report("template.xlsx"))
        self.assertEqual([path.name for path in discover_reports(self.profile)], ["a.xlsx"])
        self.profile.recursive = True
        self.assertEqual(len(discover_reports(self.profile)), 2)

    def test_recursive_pdf_paths_match(self):
        report = self.report("nested/a.xlsx")
        certs = self.root / "certs"
        (certs / "nested").mkdir(parents=True)
        pdf(certs / "nested" / "a.pdf")
        self.profile.recursive = True
        self.profile.certificate_dir = str(certs)
        self.assertEqual(match_certificate(report, self.profile), (certs / "nested" / "a.pdf").resolve())

    def test_fatal_excel_startup_is_recorded(self):
        self.report("a.xlsx")
        class BrokenExcel(FakeExcel):
            def __enter__(self):
                raise RuntimeError("Excel unavailable")
        with self.assertRaisesRegex(RuntimeError, "unavailable"):
            run_batch(self.profile, session_factory=BrokenExcel)
        manifest = json.loads(next(self.output.glob("*/batch.json")).read_text())
        self.assertEqual(manifest["fatal_error"], "Excel unavailable")
        self.assertEqual(manifest["not_processed"], 1)

    def test_pdf_encryption_rejected(self):
        approval, certificate, final = [self.root / name for name in ("approval.pdf", "cert.pdf", "final.pdf")]
        pdf(approval)
        pdf(certificate, encrypted=True)
        with self.assertRaisesRegex(ValidationError, "encrypted"):
            merge_pdfs(approval, certificate, final)
        self.assertFalse(final.exists())

    def test_corrupt_pdf_rejected(self):
        approval = self.root / "approval.pdf"
        approval.write_bytes(b"invalid")
        with self.assertRaises(Exception):
            merge_pdfs(approval, None, self.root / "final.pdf")

    def test_windows_filename_sanitization(self):
        self.assertEqual(safe_name("CON"), "_CON")
        self.assertEqual(safe_name("bad/name:*"), "bad_name__")
        self.assertEqual(safe_name("..."), "report")
        self.assertLessEqual(len(safe_name("x" * 300)), 65)
