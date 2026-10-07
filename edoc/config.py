"""User-editable mappings, validated independently of Excel and the GUI."""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path


WORKBOOK_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}


class ValidationError(ValueError):
    """An actionable configuration or report error."""


def cell_address(value: str) -> str:
    address = value.strip().replace("$", "").upper()
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]{0,6})", address)
    if not match:
        raise ValidationError(f"Invalid cell '{value}'. Use one cell such as B4, not a range.")
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - ord("A") + 1
    if column > 16384 or int(match[2]) > 1048576:
        raise ValidationError(f"Cell '{value}' is outside Excel's worksheet limits.")
    return address


@dataclass
class Mapping:
    label: str = ""
    source_sheet: str = ""
    source_cell: str = ""
    target_sheet: str = ""
    target_cell: str = ""
    mode: str = "value"
    required: bool = True
    allow_formula_overwrite: bool = False

    def validate(self) -> None:
        for name in ("label", "source_sheet", "target_sheet"):
            value = getattr(self, name)
            if not isinstance(value, str) or not value.strip():
                raise ValidationError(f"Every mapping needs a {name.replace('_', ' ')}.")
            setattr(self, name, value.strip())
        if not isinstance(self.source_cell, str) or not isinstance(self.target_cell, str):
            raise ValidationError("Cell addresses must be text.")
        self.source_cell = cell_address(self.source_cell)
        self.target_cell = cell_address(self.target_cell)
        if self.mode not in {"value", "text"}:
            raise ValidationError("Mapping mode must be 'value' or 'text'.")
        if type(self.required) is not bool or type(self.allow_formula_overwrite) is not bool:
            raise ValidationError("Mapping checkboxes must be true or false.")


@dataclass
class Profile:
    schema_version: int = 1
    template: str = ""
    input_dir: str = ""
    output_dir: str = ""
    certificate_mode: str = "report"
    certificate_dir: str = ""
    approval_sheet: str = ""
    report_sheet: str = ""
    recursive: bool = False
    keep_components: bool = True
    mappings: list[Mapping] = field(default_factory=list)

    def validate_mappings(self) -> None:
        if not self.mappings:
            raise ValidationError("Add at least one cell mapping.")
        destinations: set[tuple[str, str]] = set()
        labels: set[str] = set()
        for mapping in self.mappings:
            mapping.validate()
            destination = (mapping.target_sheet.casefold(), mapping.target_cell)
            if destination in destinations:
                raise ValidationError(f"Two mappings write to {mapping.target_sheet}!{mapping.target_cell}.")
            if mapping.label.casefold() in labels:
                raise ValidationError(f"Mapping label '{mapping.label}' is used twice.")
            destinations.add(destination)
            labels.add(mapping.label.casefold())

    def validate(self) -> None:
        self.validate_mappings()
        if self.certificate_mode not in {"report", "pdf", "approval_only"}:
            raise ValidationError("Choose report export, existing PDF, or approval only.")
        for name in ("recursive", "keep_components"):
            if type(getattr(self, name)) is not bool:
                raise ValidationError(f"{name} must be true or false.")
        for name in ("template", "input_dir", "output_dir"):
            if not isinstance(getattr(self, name), str) or not getattr(self, name).strip():
                raise ValidationError(f"Choose a {name.replace('_', ' ')}.")
        template = Path(self.template).expanduser().resolve()
        source = Path(self.input_dir).expanduser().resolve()
        output = Path(self.output_dir).expanduser().resolve()
        if not template.is_file() or template.suffix.lower() not in WORKBOOK_EXTENSIONS:
            raise ValidationError("Select an existing Excel approval template.")
        if not source.is_dir():
            raise ValidationError("The input folder does not exist.")
        if output.exists() and not output.is_dir():
            raise ValidationError("The output location must be a folder.")
        if output == source or output.is_relative_to(source):
            raise ValidationError("Choose an output folder outside the input folder.")
        if self.certificate_mode == "pdf" and not Path(self.certificate_dir).expanduser().is_dir():
            raise ValidationError("Choose the folder containing matching certificate PDFs.")

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> Profile:
        if not isinstance(data, dict) or data.get("schema_version", 1) != 1:
            raise ValidationError("This is not a supported version 1 mapping profile.")
        try:
            values = dict(data)
            values["mappings"] = [Mapping(**item) for item in values.get("mappings", [])]
            profile = cls(**values)
            profile.validate_mappings()
            for name in ("template", "input_dir", "output_dir", "certificate_dir", "approval_sheet", "report_sheet"):
                if not isinstance(getattr(profile, name), str):
                    raise ValidationError(f"{name} must be text.")
            if profile.certificate_mode not in {"report", "pdf", "approval_only"}:
                raise ValidationError("Invalid certificate mode in profile.")
            if type(profile.recursive) is not bool or type(profile.keep_components) is not bool:
                raise ValidationError("Profile checkboxes must be true or false.")
            return profile
        except (TypeError, AttributeError) as exc:
            raise ValidationError(f"Invalid profile structure: {exc}") from exc


def load_profile(path: Path) -> Profile:
    try:
        return Profile.from_dict(json.loads(path.read_text(encoding="utf-8")))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValidationError(f"Could not load profile: {exc}") from exc


def atomic_json(path: Path, data: dict) -> None:
    """Replace settings/audit JSON only after a complete successful write."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".edoc-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2, default=str, allow_nan=False)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


def save_profile(path: Path, profile: Profile) -> None:
    profile.validate_mappings()
    atomic_json(path, profile.to_dict())
