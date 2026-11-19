import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from Word.word_checker import (
    check_folder,
    compare_documents,
    reference_style_names,
    scoring_part_names,
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
        paragraph.add_run(text.replace(" ", "  ", 1) if double else text)
        if empty:
            document.add_paragraph("")
        document.save(path)

    def test_exact_document_match_has_no_differences_and_reports_every_style(self):
        self._make_document(self.student)
        result = compare_documents(self.reference, self.student)
        self.assertTrue(result.matches)
        self.assertEqual(reference_style_names(self.reference), ["Standaard"])
        report = self.root / "report.html"
        write_report(
            result,
            report,
            {
                name: 2 if name == "Stijl: Standaard" else 1
                for name in scoring_part_names(["Standaard"])
            },
        )
        content = report.read_text(encoding="utf-8")
        self.assertIn("Geen verschillen gevonden; alle ingestelde punten zijn toegekend.", content)
        self.assertIn("Stijl: Standaard", content)
        self.assertLess(
            content.index("Document opschonen"),
            content.index("Stijl: Standaard"),
        )
        self.assertIn("Model: afwezig; leerling: afwezig;", content)
        self.assertIn('data-auto-award="true"', content)

    def test_document_cleanup_is_first_in_points_distribution(self):
        names = scoring_part_names(["Standaard", "Kop 1"])
        self.assertEqual(
            names,
            [
                "Document opschonen",
                "Stijlen op juiste alinea's",
                "Stijl: Standaard",
                "Stijl: Kop 1",
                "Automatische nummering",
                "Inhoudsopgave",
                "Pagina-eindes",
            ],
        )

    def test_builtin_style_names_are_shown_in_dutch(self):
        document = Document()
        document.add_paragraph("Titel", style="Title")
        document.add_paragraph("Kop", style="Heading 1")
        document.add_paragraph("Tekst", style="Normal")
        localized_heading = document.styles.add_style("Heading_1", 1)
        document.add_paragraph("Aangepaste kop", style=localized_heading)
        toc_style = document.styles.add_style("TOC 1", 1)
        document.add_paragraph("Inhoud", style=toc_style)
        document.save(self.reference)

        self.assertEqual(
            reference_style_names(self.reference),
            ["Titel", "Kop 1", "Standaard"],
        )

    def test_used_styles_are_ordered_by_heading_hierarchy(self):
        document = Document()
        document.add_paragraph("Tekst", style="Normal")
        document.add_paragraph("Kop 2", style="Heading 2")
        document.add_paragraph("Anders", style="Quote")
        document.add_paragraph("Kop 1", style="Heading 1")
        document.add_paragraph("Subtitel", style="Subtitle")
        document.add_paragraph("Titel", style="Title")
        document.add_paragraph("Kop 3", style="Heading 3")
        document.save(self.reference)

        self.assertEqual(
            reference_style_names(self.reference),
            [
                "Titel",
                "Ondertitel",
                "Kop 1",
                "Kop 2",
                "Kop 3",
                "Standaard",
                "Citaat",
            ],
        )

    def test_reports_character_changes_double_spaces_and_empty_paragraphs(self):
        self._make_document(self.student, text="Een voorbeeld?", double=True, empty=True)
        result = compare_documents(self.reference, self.student)
        findings = result.sections["Stijl: Standaard"]
        items = [finding.item for finding in findings]
        self.assertIn("Alinea 1 — tekenverschil", items)
        cleanup_items = [finding.item for finding in result.sections["Document opschonen"]]
        self.assertIn("Alinea 1 — dubbele spaties", cleanup_items)
        self.assertIn("Lege alinea's", cleanup_items)

    def test_cleanup_differences_do_not_reduce_style_scores(self):
        reference = Document()
        reference.add_paragraph("Eerste alinea.")
        reference.add_paragraph("Tweede alinea.")
        reference.save(self.reference)

        student = Document()
        student.add_paragraph("Eerste  alinea.")
        student.add_paragraph("")
        student.add_paragraph("Tweede alinea.")
        student.save(self.student)

        result = compare_documents(self.reference, self.student)
        self.assertTrue(result.sections["Stijl: Standaard"] == [])
        self.assertEqual(
            [finding.item for finding in result.sections["Document opschonen"]],
            ["Lege alinea's", "Alinea 1 — dubbele spaties"],
        )

    def test_compares_style_definitions_and_style_application(self):
        self._make_document(self.student)
        document = Document(self.student)
        document.styles["Normal"].font.size = Pt(18)
        document.paragraphs[0].style = "Title"
        document.save(self.student)
        result = compare_documents(self.reference, self.student)
        self.assertIn("Stijldefinitie", [f.item for f in result.sections["Stijl: Standaard"]])
        self.assertIn(
            "Alinea 1 — toegepaste stijl",
            [f.item for f in result.sections["Stijlen op juiste alinea's"]],
        )
        self.assertNotIn(
            "Alinea 1 — toegepaste stijl",
            [f.item for f in result.sections["Stijl: Standaard"]],
        )

    def test_title_style_hints_must_be_removed_during_cleanup(self):
        reference = Document()
        reference.add_paragraph("Titel van document (1)", style="Heading 1")
        reference.add_paragraph("Tweede titel (2)", style="Heading 2")
        reference.save(self.reference)

        student = Document()
        student.add_paragraph("Titel van document", style="Heading 1")
        student.add_paragraph("Tweede titel (2)", style="Heading 2")
        student.save(self.student)

        result = compare_documents(self.reference, self.student)
        cleanup_findings = result.sections["Document opschonen"]
        self.assertEqual(
            [(finding.expected, finding.actual) for finding in cleanup_findings],
            [("(2)", "stijlaanwijzing (2) is niet verwijderd")],
        )
        self.assertEqual(result.sections["Stijl: Kop 1"], [])
        self.assertEqual(result.sections["Stijl: Kop 2"], [])
        self.assertEqual(result.sections["Stijlen op juiste alinea's"], [])

    def test_style_assignment_is_reported_in_its_own_scoring_category(self):
        reference = Document()
        reference.add_paragraph("Hoofdtitel", style="Heading 1")
        reference.save(self.reference)
        student = Document()
        student.add_paragraph("Hoofdtitel", style="Normal")
        student.save(self.student)

        result = compare_documents(self.reference, self.student)
        self.assertEqual(len(result.sections["Stijlen op juiste alinea's"]), 1)
        self.assertEqual(result.sections["Stijlen op juiste alinea's"][0].expected, "Kop 1")
        self.assertEqual(result.sections["Stijlen op juiste alinea's"][0].actual, "Standaard")
        self.assertEqual(result.sections["Stijl: Kop 1"], [])

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

    def test_compares_toc_entries_and_ignores_toc_style_differences(self):
        def add_toc(path, entries, font_size):
            document = Document()
            document.add_paragraph("Documenttitel", style="Heading 1")
            toc_style = document.styles.add_style("TOC 1", WD_STYLE_TYPE.PARAGRAPH)
            toc_style.font.size = Pt(font_size)
            paragraph = document.add_paragraph(style=toc_style)
            run = paragraph.add_run()
            begin = OxmlElement("w:fldChar")
            begin.set(qn("w:fldCharType"), "begin")
            instruction = OxmlElement("w:instrText")
            instruction.text = ' TOC \\o "1-3" '
            separate = OxmlElement("w:fldChar")
            separate.set(qn("w:fldCharType"), "separate")
            run._r.extend((begin, instruction, separate))
            for index, entry in enumerate(entries):
                entry_paragraph = paragraph if index == 0 else document.add_paragraph(style=toc_style)
                entry_paragraph.add_run(entry)
            end = OxmlElement("w:fldChar")
            end.set(qn("w:fldCharType"), "end")
            document.paragraphs[-1].add_run()._r.append(end)
            document.add_paragraph("Documenttekst")
            document.save(path)

        add_toc(self.reference, ["Hoofdstuk 1\t1", "Hoofdstuk 2\t2"], 10)
        add_toc(self.student, ["Hoofdstuk 1\t1", "Extra hoofdstuk\t3"], 18)
        result = compare_documents(self.reference, self.student)

        toc_findings = result.sections["Inhoudsopgave"]
        self.assertIn("Titel ontbreekt", [finding.item for finding in toc_findings])
        self.assertIn("Extra titel", [finding.item for finding in toc_findings])
        self.assertNotIn("Stijl: Inhoudsopgave 1", result.sections)
        self.assertNotIn(
            "Stijl: Inhoudsopgave 1",
            reference_style_names(self.reference),
        )
        self.assertNotIn("tekenverschil", [
            finding.item for findings in result.sections.values() for finding in findings
        ])

    def test_check_folder_writes_reports_for_docx_files(self):
        students = self.root / "students"
        students.mkdir()
        self._make_document(students / "leerling.docx")
        reports = check_folder(
            self.reference, students, self.root / "reports", {"Stijl: Standaard": 2}
        )
        self.assertEqual([path.name for path in reports], ["leerling_rapport.html"])


if __name__ == "__main__":
    unittest.main()
