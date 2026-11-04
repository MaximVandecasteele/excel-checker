import tempfile
import unittest
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import Rule
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Side
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.filters import FilterColumn, Filters, SortCondition, SortState
from openpyxl.worksheet.table import Table

from excel_checker import (
    _column_dimension_signatures,
    _conditional_rule_signature,
    _cell_value,
    _format_column_differences,
    _format_row_differences,
    _format_cf_rule,
    _normalize_number_format,
    _row_dimension_signatures,
    _side,
    compare_workbooks,
)


class WorkbookComparisonTests(unittest.TestCase):
    def test_missing_border_side_is_supported(self):
        self.assertIsNone(_side(None))
        self.assertIsNone(_side(Side()))

    def test_empty_cells_with_different_internal_types_are_equal(self):
        workbook = Workbook()
        regular_blank = workbook.active["A1"]
        typed_blank = workbook.active["B1"]
        typed_blank.data_type = "str"
        self.assertEqual(_cell_value(regular_blank), _cell_value(typed_blank))

    def test_escaped_literal_number_formats_are_normalized(self):
        self.assertEqual(
            _normalize_number_format(r"dd\ mmm\ yyyy"),
            _normalize_number_format("dd mmm yyyy"),
        )
        self.assertEqual(
            _normalize_number_format(r"\€\ #,##0.00"),
            _normalize_number_format("€ #,##0.00"),
        )

    def test_conditional_formatting_summary_is_concise(self):
        rule = Rule(
            type="expression",
            formula=['F5="Geblokkeerd"'],
            dxf=DifferentialStyle(
                fill=PatternFill(fill_type="solid", fgColor="FFFECACA")
            ),
        )
        summary = _format_cf_rule("F5:F64", rule)
        self.assertIn('F5:F64 — formuleregel; voorwaarde: F5="Geblokkeerd"', summary)
        self.assertIn("achtergrond #FECACA", summary)
        self.assertNotIn("dxfId", summary)
        self.assertNotIn("ConditionalFormatting", summary)

    def test_conditional_formatting_ignores_internal_ids_and_priorities(self):
        formatting = DifferentialStyle(
            fill=PatternFill(fill_type="solid", fgColor="FFFECACA")
        )
        first = Rule(
            type="expression", formula=["A1>0"], dxfId=1, priority=1, dxf=formatting
        )
        second = Rule(
            type="expression", formula=["A1>0"], dxfId=4, priority=8, dxf=formatting
        )
        self.assertEqual(
            _conditional_rule_signature(first),
            _conditional_rule_signature(second),
        )

    def test_conditional_formatting_difference_reports_essentials(self):
        for path, color in ((self.reference, "FFFECACA"), (self.student, "FFDCFCE7")):
            workbook = Workbook()
            sheet = workbook.active
            rule = Rule(
                type="expression",
                formula=['F5="Geblokkeerd"'],
                dxf=DifferentialStyle(
                    fill=PatternFill(fill_type="solid", fgColor=color)
                ),
            )
            sheet.conditional_formatting.add("F5:F64", rule)
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item == "Voorwaardelijke opmaak"
        )
        self.assertIn("F5:F64", difference.expected)
        self.assertIn("achtergrond #FECACA", difference.expected)
        self.assertIn("achtergrond #DCFCE7", difference.actual)
        self.assertNotIn("ConditionalFormatting", difference.expected)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.reference = self.directory / "model.xlsx"
        self.student = self.directory / "leerling.xlsx"

    def tearDown(self):
        self.temp.cleanup()

    def _workbook_with_chart(self, path, chart_title="Verkoop", anchor="D1"):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Gegevens"
        sheet.append(["Maand", "Aantal"])
        sheet.append(["Jan", 10])
        sheet.append(["Feb", 20])
        chart = BarChart()
        chart.title = chart_title
        chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=3), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=3))
        sheet.add_chart(chart, anchor)
        workbook.save(path)

    def test_chart_comparison_ignores_changed_source_cell_values(self):
        for path, values in (
            (self.reference, (10, 20)),
            (self.student, (100, 200)),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Gegevens"
            sheet.append(["Maand", "Aantal"])
            sheet.append(["Jan", values[0]])
            sheet.append(["Feb", values[1]])
            chart = BarChart()
            chart.title = "Verkoop"
            chart.add_data(
                Reference(sheet, min_col=2, min_row=1, max_row=3),
                titles_from_data=True,
            )
            chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=3))
            sheet.add_chart(chart, "D1")
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        self.assertFalse(any(item.item == "Grafieken" for item in result.differences))

    def test_chart_position_is_ignored(self):
        self._workbook_with_chart(self.reference, anchor="D1")
        self._workbook_with_chart(self.student, anchor="J20")
        self.assertTrue(compare_workbooks(self.reference, self.student).matches)

    def test_chart_default_settings_survive_excel_round_trip(self):
        self._workbook_with_chart(self.reference)
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Gegevens"
        sheet.append(["Maand", "Aantal"])
        sheet.append(["Jan", 10])
        sheet.append(["Feb", 20])
        chart = BarChart()
        chart.title = "Verkoop"
        chart.add_data(
            Reference(sheet, min_col=2, min_row=1, max_row=3),
            titles_from_data=True,
        )
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=3))
        chart.x_axis.crosses = "autoZero"
        chart.y_axis.crosses = "autoZero"
        chart.dLbls = DataLabelList(
            showLegendKey=False,
            showVal=False,
            showCatName=False,
            showSerName=False,
            showPercent=False,
            showBubbleSize=False,
        )
        sheet.add_chart(chart, "D1")
        workbook.save(self.student)
        self.assertTrue(compare_workbooks(self.reference, self.student).matches)

    def test_chart_configuration_is_compared(self):
        self._workbook_with_chart(self.reference, chart_title="Verkoop")
        self._workbook_with_chart(self.student, chart_title="Omzet")
        result = compare_workbooks(self.reference, self.student)
        difference = next(item for item in result.differences if item.item == "Grafieken")
        self.assertIn('titel "Verkoop"', difference.expected)
        self.assertIn('titel "Omzet"', difference.actual)
        self.assertNotIn("chartSpace", difference.expected)
        self.assertNotIn("Parameters:", difference.expected)

    def test_chart_source_range_difference_is_readable(self):
        self._workbook_with_chart(self.reference)
        self._workbook_with_chart(self.student)
        workbook = load_workbook(self.student)
        workbook.active._charts[0].series[0].val.numRef.f = "'Gegevens'!$B$2:$B$2"
        workbook.save(self.student)

        result = compare_workbooks(self.reference, self.student)
        difference = next(item for item in result.differences if item.item == "Grafieken")
        self.assertIn("waarden Gegevens!$B$2:$B$3", difference.expected)
        self.assertIn("waarden Gegevens!$B$2:$B$2", difference.actual)

    def _workbook_with_sorted_filtered_data(self, path, rows, apply_filter, hidden_rows=()):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data"
        sheet.append(["Categorie", "Omzet"])
        for row in rows:
            sheet.append(row)
        if apply_filter:
            sheet.auto_filter.ref = "A1:B4"
            sheet.auto_filter.filterColumn = [
                FilterColumn(colId=0, filters=Filters(filter=["Rowing Machine"]))
            ]
            sheet.auto_filter.sortState = SortState(
                ref="A2:B4",
                sortCondition=[SortCondition(ref="B1:B4", descending=True)],
            )
            for row in hidden_rows:
                sheet.row_dimensions[row].hidden = True
        workbook.save(path)

    def test_correct_sort_and_missing_filter_are_reported_without_cell_noise(self):
        source_rows = [
            ("Treadmill", 10),
            ("Rowing Machine", 20),
            ("Rowing Machine", 30),
        ]
        sorted_rows = sorted(source_rows, key=lambda row: row[1], reverse=True)
        self._workbook_with_sorted_filtered_data(
            self.reference, source_rows, apply_filter=True, hidden_rows=(2, 4)
        )
        self._workbook_with_sorted_filtered_data(
            self.student, sorted_rows, apply_filter=False
        )

        result = compare_workbooks(self.reference, self.student)
        self.assertFalse(any("inhoud/formule" in item.item for item in result.differences))
        self.assertFalse(any(item.item == "Sortering" for item in result.differences))
        filter_difference = next(item for item in result.differences if item.item == "Filter")
        self.assertIn("Categorie: waarde in Rowing Machine", filter_difference.expected)
        self.assertEqual(filter_difference.actual, "Geen filter ingesteld")
        self.assertFalse(any(item.item == "Rijhoogtes/-instellingen" for item in result.differences))

    def test_incorrect_sort_is_reported_once_for_reordered_rows(self):
        source_rows = [
            ("Treadmill", 10),
            ("Rowing Machine", 20),
            ("Rowing Machine", 30),
        ]
        self._workbook_with_sorted_filtered_data(
            self.reference, source_rows, apply_filter=True
        )
        self._workbook_with_sorted_filtered_data(
            self.student, source_rows, apply_filter=False
        )

        result = compare_workbooks(self.reference, self.student)
        self.assertFalse(any("inhoud/formule" in item.item for item in result.differences))
        sort_differences = [item for item in result.differences if item.item == "Sortering"]
        self.assertEqual(len(sort_differences), 1)
        self.assertIn("Omzet (aflopend)", sort_differences[0].expected)

    def test_changed_rows_are_not_mistaken_for_sorting_only(self):
        source_rows = [
            ("Treadmill", 10),
            ("Rowing Machine", 30),
            ("Rowing Machine", 20),
            ("Airbike", 40),
        ]
        edited_rows = [
            ("Treadmill", 10),
            ("Rowing Machine", 20),
            ("Rowing Machine", 25),
            ("Airbike", 40),
        ]
        self._workbook_with_sorted_filtered_data(
            self.reference, source_rows, apply_filter=True
        )
        self._workbook_with_sorted_filtered_data(
            self.student, edited_rows, apply_filter=False
        )

        result = compare_workbooks(self.reference, self.student)
        self.assertTrue(any("inhoud/formule" in item.item for item in result.differences))
        self.assertTrue(any(item.item == "Sortering" for item in result.differences))

    def test_matching_filter_criteria_are_not_reported(self):
        sorted_rows = [
            ("Rowing Machine", 30),
            ("Airbike", 20),
            ("Treadmill", 10),
        ]
        self._workbook_with_sorted_filtered_data(
            self.reference, sorted_rows, apply_filter=True
        )
        self._workbook_with_sorted_filtered_data(
            self.student, sorted_rows, apply_filter=True
        )

        result = compare_workbooks(self.reference, self.student)
        self.assertFalse(any(item.item == "Filter" for item in result.differences))
        self.assertFalse(any(item.item == "Sortering" for item in result.differences))

    def test_sort_is_checked_only_for_rows_matching_the_model_filter(self):
        expected_rows = [
            ("Treadmill", 50),
            ("Rowing Machine", 10),
            ("Rowing Machine", 30),
            ("Airbike", 40),
        ]
        student_rows = [
            ("Treadmill", 50),
            ("Rowing Machine", 30),
            ("Airbike", 40),
            ("Rowing Machine", 10),
        ]
        self._workbook_with_sorted_filtered_data(
            self.reference, expected_rows, apply_filter=True
        )
        self._workbook_with_sorted_filtered_data(
            self.student, student_rows, apply_filter=True
        )

        result = compare_workbooks(self.reference, self.student)
        self.assertFalse(any(item.item == "Sortering" for item in result.differences))

    def test_grouped_and_individual_column_widths_are_equivalent(self):
        individual = Workbook()
        individual_sheet = individual.active
        individual_sheet.column_dimensions["E"].width = 15
        individual_sheet.column_dimensions["F"].width = 15
        individual.save(self.reference)

        grouped = Workbook()
        grouped_sheet = grouped.active
        grouped_sheet.column_dimensions.group("E", "F", outline_level=0)
        grouped_sheet.column_dimensions["E"].width = 15
        grouped.save(self.student)

        expected = load_workbook(self.reference).active
        actual = load_workbook(self.student).active
        self.assertEqual(
            _column_dimension_signatures(expected),
            _column_dimension_signatures(actual),
        )
        self.assertTrue(compare_workbooks(self.reference, self.student).matches)

    def test_table_list_string_values_do_not_break_comparison(self):
        for path in (self.reference, self.student):
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Name", "Value"])
            sheet.append(["Item", 5])
            sheet.add_table(Table(displayName="Items", ref="A1:B2"))
            workbook.save(path)

        loaded = load_workbook(self.reference)
        table_list = loaded.active.tables
        self.assertIsInstance(table_list["Items"], Table)
        self.assertIsInstance(table_list.items()[0][1], str)
        result = compare_workbooks(self.reference, self.student)
        self.assertIsNone(result.error)

    def test_column_width_difference_names_only_the_affected_column(self):
        for path, width in ((self.reference, 15), (self.student, 20)):
            workbook = Workbook()
            workbook.active.column_dimensions["C"].width = width
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item == "Kolombreedtes/-instellingen"
        )
        self.assertIn("C:", difference.expected)
        self.assertIn("C:", difference.actual)
        self.assertNotIn("A:", difference.expected)

    def test_column_ranges_are_compacted_in_reports(self):
        expected = {
            column: (15.0, False, 0, False, False)
            for column in range(5, 8)
        }
        actual = {
            column: (18.0, False, 0, False, False)
            for column in range(5, 8)
        }
        expected_text, actual_text = _format_column_differences(expected, actual)
        self.assertEqual(expected_text, "E–G: breedte 15.0")
        self.assertEqual(actual_text, "E–G: breedte 18.0")

    def test_row_ranges_are_compacted_in_reports(self):
        expected = {}
        actual = {
            row: (20.0, False, 0, False)
            for row in range(4, 8)
        }
        expected_text, actual_text = _format_row_differences(expected, actual)
        self.assertEqual(expected_text, "Rijen 4–7: standaard")
        self.assertEqual(actual_text, "Rijen 4–7: hoogte 20.0")

    def test_default_width_range_is_not_reported_as_changed(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.sheet_format.defaultColWidth = 8.7109375
        column = sheet.column_dimensions["H"]
        column.min = 8
        column.max = 16384
        column.width = 8.7109375
        workbook.save(self.reference)
        loaded_sheet = load_workbook(self.reference).active
        self.assertEqual(_column_dimension_signatures(loaded_sheet), {})

    def test_default_row_heights_and_contiguous_ranges_are_normalized(self):
        workbook = Workbook()
        sheet = workbook.active
        for row in range(2, 6):
            sheet.row_dimensions[row].height = 15
            sheet.row_dimensions[row].hidden = True
        workbook.save(self.reference)
        loaded_sheet = load_workbook(self.reference).active
        self.assertEqual(
            _row_dimension_signatures(loaded_sheet),
            {row: (None, True, 0, False) for row in range(2, 6)},
        )

    def test_formula_and_formatting_are_compared(self):
        for path, formula, bold in (
            (self.reference, "=SUM(A1:A2)", True),
            (self.student, "=A1+A2", False),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = 2
            sheet["A2"] = 3
            sheet["A3"] = formula
            font = copy(sheet["A3"].font)
            font.bold = bold
            sheet["A3"].font = font
            workbook.save(path)
        result = compare_workbooks(self.reference, self.student)
        labels = [difference.item for difference in result.differences]
        self.assertIn("A3 inhoud/formule", labels)
        self.assertTrue(any(label.startswith("Opmaak: lettertype") for label in labels))

    def test_copied_formula_errors_in_a_row_are_grouped(self):
        for path, multiplier in ((self.reference, 2), (self.student, 3)):
            workbook = Workbook()
            sheet = workbook.active
            for column in range(2, 6):
                previous = chr(ord("A") + column - 2)
                current = chr(ord("A") + column - 1)
                sheet[f"{current}2"] = f"={previous}2*{multiplier}"
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formula_differences = [
            item for item in result.differences
            if "inhoud/formule" in item.item
        ]
        self.assertEqual(len(formula_differences), 1)
        self.assertIn("B2:E2", formula_differences[0].item)
        self.assertIn("4 cellen", formula_differences[0].item)
        self.assertEqual(formula_differences[0].expected, "=A2*2")
        self.assertEqual(formula_differences[0].actual, "=A2*3")

    def test_copied_formula_errors_in_a_column_are_grouped(self):
        for path, multiplier in ((self.reference, 2), (self.student, 3)):
            workbook = Workbook()
            sheet = workbook.active
            for row in range(2, 6):
                sheet[f"B{row}"] = f"=A{row}*{multiplier}"
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formula_differences = [
            item for item in result.differences
            if "inhoud/formule" in item.item
        ]
        self.assertEqual(len(formula_differences), 1)
        self.assertIn("B2:B5", formula_differences[0].item)
        self.assertIn("4 cellen", formula_differences[0].item)
        self.assertIn("kolom", formula_differences[0].item)

    def test_isolated_formula_error_is_not_grouped(self):
        for path, formula in (
            (self.reference, "=A1+1"),
            (self.student, "=A1+2"),
        ):
            workbook = Workbook()
            workbook.active["B1"] = formula
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formula_differences = [
            item for item in result.differences
            if "inhoud/formule" in item.item
        ]
        self.assertEqual(len(formula_differences), 1)
        self.assertEqual(formula_differences[0].item, "B1 inhoud/formule")

    def test_large_contiguous_cell_range_is_summarized(self):
        for path, offset in ((self.reference, 0), (self.student, 100)):
            workbook = Workbook()
            sheet = workbook.active
            for row in range(1, 6):
                for column in range(1, 6):
                    sheet.cell(row, column, row * 10 + column + offset)
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        content_differences = [
            item for item in result.differences
            if "inhoud/formule" in item.item or "Inhoud/formules" in item.item
        ]
        self.assertEqual(len(content_differences), 1)
        self.assertIn("A1:E5", content_differences[0].item)
        self.assertIn("25 afwijkende cellen", content_differences[0].item)
        self.assertIn("A1:", content_differences[0].expected)
        self.assertIn("A1:", content_differences[0].actual)

    def test_small_separate_cell_differences_remain_individual(self):
        for path, value in ((self.reference, 1), (self.student, 2)):
            workbook = Workbook()
            workbook.active["A1"] = value
            workbook.active["C3"] = value
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        content_differences = [
            item for item in result.differences
            if item.item.endswith("inhoud/formule")
        ]
        self.assertEqual(len(content_differences), 2)

    def test_copied_formatting_error_in_a_row_is_grouped(self):
        for path, color in ((self.reference, "FFFF0000"), (self.student, "FF0000FF")):
            workbook = Workbook()
            sheet = workbook.active
            for column in range(1, 5):
                sheet.cell(1, column).fill = PatternFill(
                    fill_type="solid", fgColor=color
                )
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formatting_differences = [
            item for item in result.differences
            if item.item.startswith("Opmaak:")
        ]
        self.assertEqual(len(formatting_differences), 1)
        self.assertIn("A1:D1", formatting_differences[0].item)
        self.assertIn("4 cellen", formatting_differences[0].item)
        self.assertIn("opvulling", formatting_differences[0].item)

    def test_copied_formatting_error_in_a_column_is_grouped(self):
        for path, bold in ((self.reference, False), (self.student, True)):
            workbook = Workbook()
            sheet = workbook.active
            for row in range(1, 5):
                sheet.cell(row, 2).font = Font(bold=bold)
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formatting_differences = [
            item for item in result.differences
            if item.item.startswith("Opmaak:")
        ]
        self.assertEqual(len(formatting_differences), 1)
        self.assertIn("B1:B4", formatting_differences[0].item)
        self.assertIn("4 cellen", formatting_differences[0].item)
        self.assertIn("lettertype", formatting_differences[0].item)

    def test_same_formatting_difference_in_separate_ranges_is_one_note(self):
        for path, color in ((self.reference, "FFFF0000"), (self.student, "FF0000FF")):
            workbook = Workbook()
            sheet = workbook.active
            for coordinate in ("A1", "B1", "D1", "E1", "A3", "B3"):
                sheet[coordinate].fill = PatternFill(
                    fill_type="solid", fgColor=color
                )
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formatting_differences = [
            item for item in result.differences
            if item.item.startswith("Opmaak:")
        ]
        self.assertEqual(len(formatting_differences), 1)
        self.assertIn("6 cellen", formatting_differences[0].item)
        self.assertIn("A1:B1", formatting_differences[0].item)
        self.assertIn("D1:E1", formatting_differences[0].item)
        self.assertIn("A3:B3", formatting_differences[0].item)
        self.assertIn("#FF0000", formatting_differences[0].expected)
        self.assertIn("#0000FF", formatting_differences[0].actual)

    def test_theme_fill_difference_is_reported_in_readable_terms(self):
        for path, theme, tint in (
            (self.reference, 3, 0.9),
            (self.student, 4, 0.8),
        ):
            workbook = Workbook()
            workbook.active["A1"].fill = PatternFill(
                fill_type="solid",
                fgColor=Color(type="theme", theme=theme, tint=tint),
            )
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item.startswith("Opmaak: opvulling")
        )
        self.assertIn("themakleur 3 (tint 0.90)", difference.expected)
        self.assertIn("themakleur 4 (tint 0.80)", difference.actual)
        self.assertNotIn("'theme'", difference.expected)

    def test_isolated_formatting_error_is_not_grouped(self):
        for path, color in ((self.reference, "FFFF0000"), (self.student, "FF0000FF")):
            workbook = Workbook()
            workbook.active["B1"].fill = PatternFill(
                fill_type="solid", fgColor=color
            )
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        formatting_differences = [
            item for item in result.differences
            if item.item.startswith("Opmaak:")
        ]
        self.assertEqual(len(formatting_differences), 1)
        self.assertEqual(
            formatting_differences[0].item,
            "Opmaak: opvulling (1 cel; bereiken: B1)",
        )


if __name__ == "__main__":
    unittest.main()
