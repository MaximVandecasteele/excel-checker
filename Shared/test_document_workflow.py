import tempfile
import unittest
from pathlib import Path

from Shared.document_workflow import review_folder


class DocumentWorkflowTests(unittest.TestCase):
    def test_review_folder_delegates_document_specific_work_and_preserves_relative_paths(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            reference = root / "model.docx"
            input_folder = root / "students"
            output_folder = root / "reports"
            reference.touch()
            (input_folder / "Student A" / "Essay.docx").parent.mkdir(parents=True)
            (input_folder / "Student A" / "Essay.docx").write_text("essay", encoding="utf-8")
            (input_folder / "Student B").mkdir()
            (input_folder / "Student B" / "Essay.docx").write_text("other", encoding="utf-8")
            (input_folder / "ignored.txt").write_text("ignore", encoding="utf-8")
            (input_folder / "~$temporary.docx").touch()

            compared: list[tuple[Path, Path]] = []

            def compare(model: Path, student: Path) -> str:
                compared.append((model, student))
                return student.read_text(encoding="utf-8")

            def write_report(result: str, destination: Path) -> None:
                destination.write_text(result, encoding="utf-8")

            reports = review_folder(
                reference,
                input_folder,
                output_folder,
                file_extension=".docx",
                compare=compare,
                write_report=write_report,
            )

            self.assertEqual(len(compared), 2)
            self.assertEqual(
                [path.name for path in reports],
                ["Student_A_Essay_rapport.html", "Student_B_Essay_rapport.html"],
            )
            self.assertEqual(
                [path.read_text(encoding="utf-8") for path in reports],
                ["essay", "other"],
            )

    def test_review_folder_applies_case_insensitive_suffix_and_avoids_reference_file(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_folder = root / "students"
            input_folder.mkdir()
            reference = input_folder / "model_FINAL.xlsx"
            reference.touch()
            (input_folder / "task_FINAL.XLSX").touch()
            (input_folder / "task_draft.xlsx").touch()
            (input_folder / "other_FINAL.xlsx").touch()
            reports = review_folder(
                reference,
                input_folder,
                root / "reports",
                file_extension=".xlsx",
                compare=lambda _reference, student: student.name,
                write_report=lambda result, destination: destination.write_text(result),
                filename_suffix="final",
            )

            self.assertEqual([path.name for path in reports], [
                "other_FINAL_rapport.html",
                "task_FINAL_rapport.html",
            ])

    def test_review_folder_ignores_whitespace_in_suffix_and_file_stem(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_folder = root / "students"
            input_folder.mkdir()
            reference = root / "model.docx"
            reference.touch()
            for name in ("Test_Word1.docx", "Test_Word 1.docx", "Test_WORD1  .docx"):
                (input_folder / name).touch()
            (input_folder / "Test_Word2.docx").touch()

            reports = review_folder(
                reference,
                input_folder,
                root / "reports",
                file_extension=".docx",
                compare=lambda _reference, student: student.name,
                write_report=lambda result, destination: destination.write_text(result),
                filename_suffix="Word1  ",
            )

            self.assertEqual(
                [path.stem for path in reports],
                ["Test_Word_1_rapport", "Test_WORD1_rapport", "Test_Word1_rapport"],
            )

    def test_review_folder_can_match_text_anywhere_in_name_ignoring_case_and_spaces(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_folder = root / "students"
            input_folder.mkdir()
            reference = root / "model.docx"
            reference.touch()
            for name in (
                "Test_Word1_Final.docx",
                "Test_Word 1_Final.docx",
                "Test_WORD1_Final.docx",
                "Test_Word2_Final.docx",
            ):
                (input_folder / name).touch()

            reports = review_folder(
                reference,
                input_folder,
                root / "reports",
                file_extension=".docx",
                compare=lambda _reference, student: student.name,
                write_report=lambda result, destination: destination.write_text(result),
                filename_suffix="word 1",
                filename_match="contains",
            )

            self.assertEqual(
                [path.stem for path in reports],
                [
                    "Test_WORD1_Final_rapport",
                    "Test_Word1_Final_rapport",
                    "Test_Word_1_Final_rapport",
                ],
            )

    def test_review_folder_requires_a_dot_prefixed_extension(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            with self.assertRaises(ValueError):
                review_folder(
                    root / "model",
                    root,
                    root / "reports",
                    file_extension="docx",
                    compare=lambda _reference, _student: None,
                    write_report=lambda _result, _destination: None,
                )


if __name__ == "__main__":
    unittest.main()
