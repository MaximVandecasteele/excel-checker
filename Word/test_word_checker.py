import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.shared import Pt

from Word.word_checker import (
    GLOBAL_SCORING_PARTS,
    check_folder,
    compare_documents,
    reference_style_names,
    write_report,
)


class WordCheckerTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.reference = self.root / "model.docx"
        self.student = self.root / "student.docx"
        self._make_document(self.reference)

    def tearDown(self):
        self.temporary_directory.cleanup()

    @staticmethod
    def _make_document(path, text="Een voorbeeld.", style="Normal", double=False, empty=False):
        document = Document()
        paragraph = document.add_paragraph(style=style)
        paragraph.add_run("Een  voorbeeld." if double else text)
        if empty:
            document.add_paragraph("")
        document.save(path)

    def test_exact_document_match_has_no_differences_and_reports_every_style(self):
        self._make_document(self.student)
        result = compare_documents(self.reference, self.student)
        self.assertTrue(result.matches)
        self.assertEqual(reference_style_names(self.reference), ["Normal"])
        report = self.root / "report.html"
        write_report(
            result,
            report,
            {"Stijl: Normal": 2, **{name: 1 for name in GLOBAL_SCORING_PARTS}},
        )
        content = report.read_text(encoding="utf-8")
        self.assertIn("Geen verschillen gevonden; alle ingestelde punten zijn toegekend.", content)
        self.assertIn("Stijl: Normal", content)
        self.assertIn("Model: afwezig; leerling: afwezig.", content)
        self.assertIn('data-auto-award="true"', content)

    def test_reports_character_changes_double_spaces_and_empty_paragraphs(self):
        self._make_document(self.student, text="Een voorbeeld?", double=True, empty=True)
        result = compare_documents(self.reference, self.student)
        findings = result.sections["Stijl: Normal"]
        items = [finding.item for finding in findings]
        self.assertIn("Alinea 1 — tekenverschil", items)
        self.assertIn("Alinea 1 — dubbele spaties", items)
        self.assertIn("Alinea 2 — lege alinea ontbreekt of is extra", items)

    def test_compares_style_definitions_and_style_application(self):
        self._make_document(self.student)
        document = Document(self.student)
        document.styles["Normal"].font.size = Pt(18)
        document.paragraphs[0].style = "Title"
        document.save(self.student)
        result = compare_documents(self.reference, self.student)
        self.assertIn("Stijldefinitie", [f.item for f in result.sections["Stijl: Normal"]])
        self.assertIn(
            "Alinea 1 — toegepaste stijl",
            [f.item for f in result.sections["Stijl: Normal"]],
        )

    def test_compares_page_breaks_and_auto_numbering(self):
        reference_doc = Document()
        reference_doc.add_paragraph("First")
        reference_doc.add_paragraph("Second").runs[0].add_break(WD_BREAK.PAGE)
        reference_doc.add_paragraph("Numbered", style="List Number")
        reference_doc.save(self.reference)

        student_doc = Document()
        student_doc.add_paragraph("First")
        student_doc.add_paragraph("Second")
        student_doc.add_paragraph("Numbered")
        student_doc.save(self.student)

        result = compare_documents(self.reference, self.student)
        self.assertIn("Pagina-eindes", result.sections)
        self.assertIn("Automatische nummering", result.sections)

    def test_detects_generated_table_of_contents_field(self):
        reference_doc = Document()
        paragraph = reference_doc.add_paragraph()
        instruction = paragraph.add_run()._r
        instruction.addnext(OxmlElement("w:instrText"))
        instruction.getnext().text = " TOC \\o \"1-3\" "
        reference_doc.save(self.reference)
        self._make_document(self.student)
        result = compare_documents(self.reference, self.student)
        self.assertEqual(
            result.sections["Inhoudsopgave"][0].expected,
            "aanwezig",
        )

    def test_check_folder_writes_reports_for_docx_files(self):
        students = self.root / "students"
        students.mkdir()
        self._make_document(students / "leerling.docx")
        reports = check_folder(
            self.reference, students, self.root / "reports", {"Stijl: Normal": 2}
        )
        self.assertEqual([path.name for path in reports], ["leerling_rapport.html"])


if __name__ == "__main__":
    unittest.main()
