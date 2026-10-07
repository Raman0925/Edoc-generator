import json
import tempfile
import unittest
from pathlib import Path

from edoc.config import Mapping, Profile, ValidationError, cell_address, load_profile, save_profile


def mapping(**changes):
    values = dict(label="Serial", source_sheet="Report", source_cell="B4", target_sheet="Approval", target_cell="F6")
    values.update(changes)
    return Mapping(**values)


class ConfigurationTests(unittest.TestCase):
    def test_absolute_addresses_normalized(self):
        self.assertEqual(cell_address(" $b$4 "), "B4")

    def test_excel_boundaries(self):
        self.assertEqual(cell_address("XFD1048576"), "XFD1048576")
        for address in ("XFE1", "A1048577", "A0", "A01", "A1:B2", "Sheet!A1", "=A1", "", "1A"):
            with self.subTest(address=address), self.assertRaises(ValidationError):
                cell_address(address)

    def test_empty_mapping_list_rejected(self):
        with self.assertRaises(ValidationError):
            Profile().validate_mappings()

    def test_duplicate_destination_case_insensitive(self):
        profile = Profile(mappings=[mapping(), mapping(label="Other", target_sheet="approval", target_cell="$F$6")])
        with self.assertRaisesRegex(ValidationError, "Two mappings"):
            profile.validate_mappings()

    def test_duplicate_label_rejected(self):
        with self.assertRaises(ValidationError):
            Profile(mappings=[mapping(), mapping(label="serial", target_cell="F7")]).validate_mappings()

    def test_bad_mode_rejected(self):
        with self.assertRaises(ValidationError):
            mapping(mode="formula").validate()

    def test_boolean_string_rejected(self):
        with self.assertRaises(ValidationError):
            mapping(required="false").validate()

    def test_profile_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "profile.json"
            profile = Profile(template="a.xlsx", mappings=[mapping(mode="text")])
            save_profile(path, profile)
            self.assertEqual(load_profile(path).to_dict(), profile.to_dict())
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_unknown_profile_version(self):
        with self.assertRaises(ValidationError):
            Profile.from_dict({"schema_version": 99})

    def test_malformed_profile(self):
        for data in ([1], {"mappings": None}, {"mappings": ["bad"]}, {"mappings": [{}]},
                     {"mappings": [vars(mapping())], "recursive": "false"},
                     {"mappings": [vars(mapping())], "template": 4}):
            with self.subTest(data=data), self.assertRaises(ValidationError):
                Profile.from_dict(data)

    def test_invalid_json_is_actionable(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "profile.json"
            path.write_text("{", encoding="utf-8")
            with self.assertRaisesRegex(ValidationError, "Could not load"):
                load_profile(path)

    def test_output_cannot_be_inside_input(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template = root / "template.xlsx"
            template.touch()
            for output in (root, root / "output"):
                profile = Profile(template=str(template), input_dir=str(root), output_dir=str(output), mappings=[mapping()])
                with self.assertRaisesRegex(ValidationError, "outside"):
                    profile.validate()

    def test_missing_pdf_folder_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "input").mkdir()
            (root / "template.xlsx").touch()
            profile = Profile(template=str(root / "template.xlsx"), input_dir=str(root / "input"),
                              output_dir=str(root / "output"), certificate_mode="pdf", certificate_dir=str(root / "missing"), mappings=[mapping()])
            with self.assertRaisesRegex(ValidationError, "certificate"):
                profile.validate()
