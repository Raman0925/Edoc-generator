"""Batch orchestration with per-report isolation and auditable outputs."""

from __future__ import annotations

import hashlib
import json
import math
import re
import shutil
import tempfile
import threading
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from . import __version__
from .config import Profile, ValidationError, WORKBOOK_EXTENSIONS, atomic_json
from .excel import ExcelSession


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")[:65]
    if not value:
        value = "report"
    if value.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *[f"COM{i}" for i in range(1, 10)], *[f"LPT{i}" for i in range(1, 10)]}:
        value = f"_{value}"
    return value


def discover_reports(profile: Profile) -> list[Path]:
    folder = Path(profile.input_dir).expanduser().resolve()
    template = Path(profile.template).expanduser().resolve()
    files = folder.rglob("*") if profile.recursive else folder.iterdir()
    return sorted(
        (p.resolve() for p in files if p.is_file() and not p.is_symlink()
         and not p.name.startswith("~$") and p.suffix.lower() in WORKBOOK_EXTENSIONS
         and p.resolve() != template),
        key=lambda p: str(p).casefold(),
    )


def match_certificate(report: Path, profile: Profile) -> Path:
    root = Path(profile.certificate_dir).expanduser().resolve()
    relative = report.relative_to(Path(profile.input_dir).expanduser().resolve())
    folder = root / relative.parent if profile.recursive else root
    candidates = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"
                  and p.stem.casefold() == report.stem.casefold()] if folder.is_dir() else []
    if len(candidates) != 1:
        raise ValidationError(f"Expected exactly one certificate '{report.stem}.pdf' in {folder}; found {len(candidates)}.")
    return candidates[0]


def merge_pdfs(approval: Path, certificate: Path | None, output: Path) -> int:
    from pypdf import PdfReader, PdfWriter
    writer = PdfWriter()
    try:
        for path, label in ((approval, "Approval document"), (certificate, "Calibration certificate")):
            if path is None:
                continue
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                raise ValidationError(f"PDF '{path.name}' is encrypted. Provide an unencrypted copy.")
            if not reader.pages:
                raise ValidationError(f"PDF '{path.name}' contains no pages.")
            writer.append(reader, outline_item=label, import_outline=False)
        writer.add_metadata({"/Creator": f"eDoc Generator {__version__}"})
        count = len(writer.pages)
        with output.open("xb") as stream:
            writer.write(stream)
        if len(PdfReader(str(output)).pages) != count:
            raise ValidationError("Merged PDF page-count verification failed.")
        return count
    finally:
        writer.close()


def extract(session, book, profile: Profile) -> list:
    values = [session.read(book, mapping) for mapping in profile.mappings]
    for mapping, value in zip(profile.mappings, values):
        if mapping.required and (value is None or (isinstance(value, str) and not value.strip())):
            raise ValidationError(f"{mapping.label}: required source cell is empty.")
        if isinstance(value, float) and not math.isfinite(value):
            raise ValidationError(f"{mapping.label}: source value is not a finite number.")
    return values


def check_targets(session, book, profile: Profile) -> None:
    destinations: set[str] = set()
    for mapping in profile.mappings:
        resolved = session.check_target(book, mapping)
        if resolved in destinations:
            raise ValidationError("Two mappings write to the same merged cell. Use distinct destination cells.")
        destinations.add(resolved)


def preview(profile: Profile, sample: Path, session_factory=ExcelSession) -> dict:
    """Read a sample and validate destinations without modifying either workbook."""
    profile.validate_mappings()
    with session_factory() as session:
        with session.open(sample) as report:
            values = extract(session, report, profile)
            report_sheets = session.sheets(report)
        with session.open(Path(profile.template)) as template:
            check_targets(session, template, profile)
            template_sheets = session.sheets(template)
    return {"report": str(sample), "source_sheets": report_sheets, "target_sheets": template_sheets,
            "values": [{"label": m.label, "source": f"{m.source_sheet}!{m.source_cell}",
                        "target": f"{m.target_sheet}!{m.target_cell}", "mode": m.mode,
                        "value": str(v) if v is not None else "(blank)"}
                       for m, v in zip(profile.mappings, values)]}


@dataclass
class Result:
    source: str
    status: str
    output: str = ""
    error: str = ""
    pages: int = 0


@dataclass
class BatchResult:
    run_dir: str
    results: list[Result] = field(default_factory=list)
    cancelled: bool = False

    @property
    def succeeded(self) -> int:
        return sum(r.status == "success" for r in self.results)


def run_batch(profile: Profile, progress: Callable[[dict], None] | None = None,
              cancel: threading.Event | None = None, session_factory=ExcelSession) -> BatchResult:
    profile.validate()
    reports = discover_reports(profile)
    if not reports:
        raise ValidationError("No Excel reports found in the input folder.")
    emit = progress or (lambda _: None)
    cancel = cancel or threading.Event()
    root = Path(profile.output_dir).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix=datetime.now().strftime("edoc_%Y%m%d_%H%M%S_"), dir=root))
    batch = BatchResult(str(run))
    manifest = {"app_version": __version__, "started_utc": datetime.now(timezone.utc).isoformat(),
                "profile": profile.to_dict(), "total_reports": len(reports), "results": []}
    atomic_json(run / "batch.json", manifest)
    try:
        # Snapshot the template once; a mid-run template edit cannot change later outputs.
        with tempfile.TemporaryDirectory(prefix="edoc-template-") as template_temp:
            template_copy = Path(template_temp) / f"template{Path(profile.template).suffix.lower()}"
            shutil.copy2(Path(profile.template).expanduser().resolve(), template_copy)
            manifest["template_sha256"] = sha256(template_copy)
            with session_factory() as session:
                for index, report in enumerate(reports, start=1):
                    if cancel.is_set():
                        batch.cancelled = True
                        break
                    emit({"kind": "starting", "index": index, "total": len(reports), "source": str(report)})
                    try:
                        relative = report.relative_to(Path(profile.input_dir).expanduser().resolve())
                        key = hashlib.sha256(str(relative).encode()).hexdigest()[:10]
                        destination = run / f"{index:03d}_{safe_name(report.stem)}_{key}"
                        with tempfile.TemporaryDirectory(prefix=".pending-", dir=run) as temporary:
                            staging = Path(temporary)
                            approval_copy = staging / f"approval{template_copy.suffix}"
                            report_copy = staging / f"source{report.suffix.lower()}"
                            shutil.copy2(template_copy, approval_copy)
                            shutil.copy2(report, report_copy)
                            source_hash = sha256(report_copy)
                            approval_pdf = staging / "approval.pdf"
                            certificate_pdf = staging / "certificate.pdf"
                            with session.open(report_copy) as source_book:
                                values = extract(session, source_book, profile)
                                if profile.certificate_mode == "report":
                                    session.export_pdf(source_book, certificate_pdf, profile.report_sheet)
                            certificate_source = None
                            certificate_hash = None
                            if profile.certificate_mode == "pdf":
                                certificate_source = match_certificate(report, profile)
                                shutil.copy2(certificate_source, certificate_pdf)
                                certificate_hash = sha256(certificate_pdf)
                            with session.open(approval_copy, writable=True) as approval_book:
                                check_targets(session, approval_book, profile)
                                for mapping, value in zip(profile.mappings, values):
                                    session.write(approval_book, mapping, value)
                                session.calculate(approval_book)
                                session.save(approval_book)
                                session.export_pdf(approval_book, approval_pdf, profile.approval_sheet)
                            pages = merge_pdfs(approval_pdf,
                                               None if profile.certificate_mode == "approval_only" else certificate_pdf,
                                               staging / "final.pdf")
                            audit = {"source": str(report), "source_sha256": source_hash,
                                     "template_sha256": manifest["template_sha256"],
                                     "certificate_source": str(certificate_source) if certificate_source else None,
                                     "certificate_sha256": certificate_hash,
                                     "generated_utc": datetime.now(timezone.utc).isoformat(),
                                     "pages": pages, "final_sha256": sha256(staging / "final.pdf"),
                                     "values": [{**asdict(m), "value": v} for m, v in zip(profile.mappings, values)]}
                            atomic_json(staging / "audit.json", audit)
                            report_copy.unlink()
                            if not profile.keep_components:
                                approval_pdf.unlink()
                                certificate_pdf.unlink(missing_ok=True)
                            # Publish only after workbook, PDF and audit generation all succeed.
                            staging.rename(destination)
                        result = Result(str(report), "success", str(destination / "final.pdf"), pages=pages)
                    except Exception as exc:
                        result = Result(str(report), "failed", error=str(exc) or type(exc).__name__)
                    batch.results.append(result)
                    manifest["results"] = [asdict(item) for item in batch.results]
                    atomic_json(run / "batch.json", manifest)
                    emit({"kind": "result", "index": index, "total": len(reports), **asdict(result)})
    except Exception as exc:
        manifest["fatal_error"] = str(exc)
        raise
    finally:
        manifest.update({"cancelled": batch.cancelled, "finished_utc": datetime.now(timezone.utc).isoformat(),
                         "succeeded": batch.succeeded, "failed": sum(r.status == "failed" for r in batch.results),
                         "not_processed": len(reports) - len(batch.results)})
        atomic_json(run / "batch.json", manifest)
    return batch
