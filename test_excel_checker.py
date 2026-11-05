import tempfile
import unittest
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import Rule
from openpyxl.styles import PatternFill, Font, Border, Alignment
from openpyxl.styles.colors import Color
from openpyxl.styles.borders import Side
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.filters import FilterColumn, Filters, SortCondition, SortState
from openpyxl.worksheet.table import Table

from excel_checker import (
    ComparisonResult,
    Difference,
    _column_dimension_signatures,
    _conditional_rule_signature,
    _cell_value,
    _format_column_differences,
    _format_row_differences,
    _format_cf_rule,
    _normalize_number_format,
    _row_dimension_signatures,
    _side,
    _tinted_rgb,
    compare_workbooks,
    write_report,
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

    def test_builtin_number_formats_are_reported_by_category_name(self):
        for path, number_format in (
            (self.reference, '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)'),
            (self.student, '"€" #,##0.00'),
        ):
            workbook = Workbook()
            workbook.active["A1"] = 123.45
            workbook.active["A1"].number_format = number_format
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item.startswith("Opmaak: getalnotatie")
        )
        self.assertEqual(difference.expected, "getalnotatie: financieel")
        self.assertEqual(difference.actual, "getalnotatie: valuta")

    def test_same_builtin_number_format_category_includes_effective_formats(self):
        for path, number_format in (
            (self.reference, "0.00%"),
            (self.student, "0%"),
        ):
            workbook = Workbook()
            workbook.active["A1"] = 0.1234
            workbook.active["A1"].number_format = number_format
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item.startswith("Opmaak: getalnotatie")
        )
        self.assertEqual(difference.expected, "getalnotatie: percentage : 0.00%")
        self.assertEqual(difference.actual, "getalnotatie: percentage : 0%")

    def test_custom_number_formats_are_reported_in_full(self):
        custom_format = '"€" #00000000000000,##0.00'
        for path, number_format in (
            (self.reference, '"€" #,##0.0'),
            (self.student, custom_format),
        ):
            workbook = Workbook()
            workbook.active["A1"] = 123.45
            workbook.active["A1"].number_format = number_format
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item.startswith("Opmaak: getalnotatie")
        )
        self.assertEqual(difference.expected, 'getalnotatie: \'"€" #,##0.0\'')
        self.assertEqual(difference.actual, f"getalnotatie: {custom_format!r}")

    def test_localized_default_currency_and_financial_formats_use_category_names(self):
        examples = (
            (
                '_ "€" * #,##0.00_ ;_ "€" * -#,##0.00_ ;_ "€" * "-"??_ ;_ @_',
                '_ "€" * #00000000000000,##0.00_ ;_ "€" * -#,##0.00_ ;_ "€" * "-"??_ ;_ @_',
                "financieel",
            ),
            ('"€" #,##0.00', '"€" #00000000000000,##0.00', "valuta"),
        )
        for default_format, custom_format, expected_name in examples:
            with self.subTest(format=default_format):
                for path, number_format in (
                    (self.reference, default_format),
                    (self.student, custom_format),
                ):
                    workbook = Workbook()
                    workbook.active["A1"] = 123.45
                    workbook.active["A1"].number_format = number_format
                    workbook.save(path)

                result = compare_workbooks(self.reference, self.student)
                difference = next(
                    item for item in result.differences
                    if item.item.startswith("Opmaak: getalnotatie")
                )
                self.assertEqual(
                    difference.expected,
                    f"getalnotatie: {expected_name}",
                )
                self.assertIn(custom_format, difference.actual)

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

    def test_row_heights_and_column_widths_allow_a_twenty_percent_difference(self):
        for path, row_height, column_width in (
            (self.reference, 20, 15),
            (self.student, 24, 18),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet.row_dimensions[2].height = row_height
            sheet.column_dimensions["C"].width = column_width
            workbook.save(path)

        self.assertTrue(compare_workbooks(self.reference, self.student).matches)

    def test_explicit_default_row_settings_within_tolerance_are_not_reported(self):
        for path, height in ((self.reference, None), (self.student, 18)):
            workbook = Workbook()
            workbook.active["A1"] = "Verkoop"
            if height is not None:
                workbook.active.row_dimensions[3].height = height
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        self.assertFalse(any(
            item.item == "Rijhoogtes/-instellingen"
            for item in result.differences
        ))

    def test_row_heights_and_column_widths_report_differences_over_twenty_percent(self):
        for path, row_height, column_width in (
            (self.reference, 20, 15),
            (self.student, 24.5, 18.1),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet.row_dimensions[2].height = row_height
            sheet.column_dimensions["C"].width = column_width
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference_names = {item.item for item in result.differences}
        self.assertIn("Rijhoogtes/-instellingen", difference_names)
        self.assertIn("Kolombreedtes/-instellingen", difference_names)

    def test_equal_model_dimensions_must_remain_equal_in_student_workbook(self):
        for path, row_heights, column_widths in (
            (self.reference, (20, 20), (15, 15)),
            (self.student, (21, 20), (16, 15)),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet.row_dimensions[2].height = row_heights[0]
            sheet.row_dimensions[3].height = row_heights[1]
            sheet.column_dimensions["C"].width = column_widths[0]
            sheet.column_dimensions["D"].width = column_widths[1]
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        row_difference = next(
            item for item in result.differences
            if item.item == "Rijhoogtes/-instellingen"
        )
        column_difference = next(
            item for item in result.differences
            if item.item == "Kolombreedtes/-instellingen"
        )
        self.assertIn("Rij 2", row_difference.expected)
        self.assertIn("Rij 3", row_difference.expected)
        self.assertIn("C:", column_difference.expected)
        self.assertIn("D:", column_difference.expected)
        self.assertIn("gelijke modelafmetingen", row_difference.actual)
        self.assertIn("gelijke modelafmetingen", column_difference.actual)

    def test_different_model_dimensions_do_not_have_to_match_each_other(self):
        for path, widths in (
            (self.reference, (15, 20)),
            (self.student, (16, 21)),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet.column_dimensions["C"].width = widths[0]
            sheet.column_dimensions["D"].width = widths[1]
            workbook.save(path)

        self.assertTrue(compare_workbooks(self.reference, self.student).matches)

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

    def test_separate_dimension_changes_are_reported_on_separate_lines(self):
        expected = {
            2: (15.0, False, 0, False, False),
            4: (20.0, False, 0, False, False),
        }
        actual = {
            2: (18.0, False, 0, False, False),
            4: (24.0, False, 0, False, False),
        }
        expected_columns, actual_columns = _format_column_differences(expected, actual)
        expected_rows, actual_rows = _format_row_differences(
            {2: (15.0, False, 0, False), 4: (20.0, False, 0, False)},
            {2: (18.0, False, 0, False), 4: (24.0, False, 0, False)},
        )

        self.assertEqual(expected_columns, "B: breedte 15.0\nD: breedte 20.0")
        self.assertEqual(actual_columns, "B: breedte 18.0\nD: breedte 24.0")
        self.assertEqual(expected_rows, "Rij 2: hoogte 15.0\nRij 4: hoogte 20.0")
        self.assertEqual(actual_rows, "Rij 2: hoogte 18.0\nRij 4: hoogte 24.0")

    def test_dimension_report_preserves_line_breaks_in_html(self):
        for path, widths in (
            (self.reference, (10, 20)),
            (self.student, (13, 26)),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet.column_dimensions["B"].width = widths[0]
            sheet.column_dimensions["D"].width = widths[1]
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item == "Kolombreedtes/-instellingen"
        )
        self.assertIn("\n", difference.expected)
        report_path = self.directory / "rapport.html"
        write_report(result, report_path)
        report = report_path.read_text(encoding="utf-8")
        self.assertIn("B: breedte 10.0\nD: breedte 20.0", report)
        self.assertIn("white-space:pre-wrap", report)

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
        font_difference = next(
            item for item in result.differences
            if item.item.startswith("Opmaak: lettertype")
        )
        self.assertIn("vet", font_difference.expected)
        self.assertIn("normaal", font_difference.actual)
        self.assertNotIn("True", font_difference.expected + font_difference.actual)
        self.assertNotIn("False", font_difference.expected + font_difference.actual)

    def test_font_report_names_only_readable_changed_properties(self):
        for path, font in (
            (self.reference, Font(name="Arial", sz=11)),
            (self.student, Font(name="Calibri", sz=14, bold=True, color="FFFF0000")),
        ):
            workbook = Workbook()
            workbook.active["A1"] = "Tekst"
            workbook.active["A1"].font = font
            workbook.save(path)

        result = compare_workbooks(self.reference, self.student)
        difference = next(
            item for item in result.differences
            if item.item.startswith("Opmaak: lettertype")
        )
        self.assertEqual(
            difference.expected,
            "lettertype: Arial, 11 pt, normaal, standaardkleur",
        )
        self.assertEqual(
            difference.actual,
            "lettertype: Calibri, 14 pt, vet, #FF0000",
        )
        self.assertNotIn("True", difference.expected + difference.actual)
        self.assertNotIn("False", difference.expected + difference.actual)

    def test_border_report_names_only_changed_edges_in_plain_language(self):
        for path, border in (
            (self.reference, Border(left=Side(style="thin", color="FFFF0000"))),
            (
                self.student,
                Border(
                    left=Side(style="double", color="FF0000FF"),
                    bottom=Side(style="dashed", color="FF00AA00"),
                ),
            ),
        ):
            workbook = Workbook()
            workbook.active["A1"] = "Tekst"
            workbook.active["A1"].border = border
            workbook.save(path)

        difference = next(
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: rand")
        )
        self.assertIn("linkerrand stijl: dun", difference.expected)
        self.assertIn("linkerrand kleur: #FF0000", difference.expected)
        self.assertIn("linkerrand stijl: dubbel", difference.actual)
        self.assertIn("linkerrand kleur: #0000FF", difference.actual)
        self.assertIn("onderrand: geen", difference.expected)
        self.assertIn("onderrand stijl: gestreept", difference.actual)
        self.assertIn("onderrand kleur: #00AA00", difference.actual)
        self.assertEqual(difference.actual.count("\n"), 3)
        self.assertNotIn("True", difference.expected + difference.actual)
        self.assertNotIn("False", difference.expected + difference.actual)
        self.assertNotIn("None", difference.expected + difference.actual)

    def test_border_report_only_mentions_properties_that_changed(self):
        for path, border in (
            (self.reference, Border(top=Side(style="thin", color="FFFF0000"))),
            (self.student, Border(top=Side(style="thick", color="FFFF0000"))),
        ):
            workbook = Workbook()
            workbook.active["A1"] = "Tekst"
            workbook.active["A1"].border = border
            workbook.save(path)

        difference = next(
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: rand")
        )
        self.assertEqual(difference.expected, "bovenrand stijl: dun")
        self.assertEqual(difference.actual, "bovenrand stijl: dik")
        self.assertNotIn("kleur", difference.expected + difference.actual)

    def test_border_differences_are_displayed_on_separate_html_lines(self):
        for path, border in (
            (self.reference, Border(top=Side(style="thin", color="FFFF0000"))),
            (
                self.student,
                Border(
                    top=Side(style="thick", color="FF0000FF"),
                    bottom=Side(style="dashed", color="FF00AA00"),
                ),
            ),
        ):
            workbook = Workbook()
            workbook.active["A1"].border = border
            workbook.save(path)

        report_path = self.directory / "randen.html"
        write_report(compare_workbooks(self.reference, self.student), report_path)
        report = report_path.read_text(encoding="utf-8")
        self.assertIn(
            "bovenrand stijl: dik\n"
            "bovenrand kleur: #0000FF\n"
            "onderrand stijl: gestreept\n"
            "onderrand kleur: #00AA00",
            report,
        )
        self.assertIn("white-space:pre-wrap", report)

    def test_border_color_only_change_does_not_repeat_unchanged_style(self):
        for path, color in (
            (self.reference, "FFFF0000"),
            (self.student, "FF0000FF"),
        ):
            workbook = Workbook()
            workbook.active["A1"].border = Border(
                top=Side(style="thin", color=color)
            )
            workbook.save(path)

        difference = next(
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: rand")
        )
        self.assertEqual(difference.expected, "bovenrand kleur: #FF0000")
        self.assertEqual(difference.actual, "bovenrand kleur: #0000FF")
        self.assertNotIn("stijl", difference.expected + difference.actual)

    def test_alignment_report_uses_readable_changed_properties(self):
        for path, alignment in (
            (
                self.reference,
                Alignment(horizontal="center", vertical="center", wrap_text=False),
            ),
            (
                self.student,
                Alignment(horizontal="right", vertical="center", wrap_text=True),
            ),
        ):
            workbook = Workbook()
            workbook.active["A1"] = "Tekst"
            workbook.active["A1"].alignment = alignment
            workbook.save(path)

        difference = next(
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: uitlijning")
        )
        self.assertEqual(
            difference.expected,
            "horizontaal: centreren\ngeen tekstterugloop",
        )
        self.assertEqual(
            difference.actual,
            "horizontaal: rechts\ntekstterugloop",
        )
        self.assertNotIn("verticaal", difference.expected + difference.actual)
        self.assertNotIn("True", difference.expected + difference.actual)
        self.assertNotIn("False", difference.expected + difference.actual)
        self.assertNotIn("None", difference.expected + difference.actual)

    def test_contiguous_rectangle_border_changes_are_reported_as_one_range(self):
        for path, has_border in (
            (self.reference, True),
            (self.student, False),
        ):
            workbook = Workbook()
            sheet = workbook.active
            if has_border:
                for row in range(1, 6):
                    for column in range(1, 11):
                        sheet.cell(row, column).border = Border(
                            left=Side(style="thin", color="FFFF0000") if column == 1 else None,
                            right=Side(style="thin", color="FFFF0000") if column == 10 else None,
                            top=Side(style="thin", color="FFFF0000") if row == 1 else None,
                            bottom=Side(style="thin", color="FFFF0000") if row == 5 else None,
                        )
            workbook.save(path)

        border_differences = [
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: rand")
        ]
        self.assertEqual(len(border_differences), 1)
        difference = border_differences[0]
        self.assertIn("bereik A1:J5", difference.item)
        self.assertIn("bovenrand stijl: dun", difference.expected)
        self.assertEqual(difference.actual.count("\n"), 3)
        self.assertIn("bovenrand: geen", difference.actual)

    def test_theme_font_colors_are_reported_as_rgb(self):
        for path, color in (
            (self.reference, "FF000000"),
            (self.student, Color(type="theme", theme=4, tint=0.5)),
        ):
            workbook = Workbook()
            workbook.active["A1"] = "Tekst"
            workbook.active["A1"].font = Font(color=color)
            workbook.save(path)

        difference = next(
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: lettertype")
        )
        self.assertRegex(difference.actual, r"lettertype: #[0-9A-F]{6}")
        self.assertNotIn("themakleur", difference.actual)
        self.assertNotIn("tint", difference.actual)

    def test_theme_tint_matches_excel_palette_rgb_values(self):
        self.assertEqual(_tinted_rgb("4472C4", 0.8), "D9E1F2")
        self.assertEqual(_tinted_rgb("4472C4", -0.25), "305496")
        self.assertEqual(_tinted_rgb("FF0000", 0.5), "FF8080")
        self.assertEqual(_tinted_rgb("AD987D", 0.269), "C2B39E")
        self.assertEqual(_tinted_rgb("85D5A2", -0.141), "61C985")
        self.assertEqual(_tinted_rgb("BAED9D", -0.254), "7CDE47")
        self.assertEqual(_tinted_rgb("8077C7", -0.639), "262152")

    def test_theme_fill_colors_are_reported_as_rgb(self):
        for path, fill in (
            (self.reference, PatternFill()),
            (
                self.student,
                PatternFill(
                    fill_type="solid",
                    fgColor=Color(type="theme", theme=4, tint=0.5),
                ),
            ),
        ):
            workbook = Workbook()
            workbook.active["A1"] = "Tekst"
            workbook.active["A1"].fill = fill
            workbook.save(path)

        difference = next(
            item for item in compare_workbooks(
                self.reference, self.student
            ).differences
            if item.item.startswith("Opmaak: opvulling")
        )
        self.assertRegex(difference.actual, r"#[0-9A-F]{6}")
        self.assertNotIn("themakleur", difference.actual)
        self.assertNotIn("tint", difference.actual)

    def test_html_report_displays_color_swatch_for_font_and_fill(self):
        for path, color in (
            (self.reference, "FFFF0000"),
            (self.student, "FF0000FF"),
        ):
            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = "Tekst"
            sheet["A1"].font = Font(color=color)
            sheet["A1"].fill = PatternFill(fill_type="solid", fgColor=color)
            workbook.save(path)

        report_path = self.directory / "kleuren.html"
        write_report(compare_workbooks(self.reference, self.student), report_path)
        report = report_path.read_text(encoding="utf-8")
        self.assertIn(
            '<span class="color-swatch" style="background-color:#FF0000"',
            report,
        )
        self.assertIn(
            '<span class="color-swatch" style="background-color:#0000FF"',
            report,
        )
        self.assertIn(
            ".color-swatch{display:inline-block;width:50px;height:50px",
            report,
        )

    def test_html_report_groups_differences_by_sheet_and_category(self):
        result = ComparisonResult(
            self.student,
            [
                Difference("Data", "A1 inhoud/formule", "'model'", "'leerling'",
                           "Invoer data (inhoud)"),
                Difference("Data", "Opmaak: lettertype", "vet", "normaal",
                           "Opmaak"),
                Difference("Data", "Opmaak: getalnotatie (1 cel; bereiken: B2)",
                           "percentage", "valuta", "Opmaak"),
                Difference("Berekeningen", "C2 inhoud/formule", "=A2+B2", "=A2-B2",
                           "Functies / berekeningen"),
            ],
        )
        report_path = self.directory / "gegroepeerd.html"

        write_report(result, report_path)

        report = report_path.read_text(encoding="utf-8")
        self.assertIn("<h2>Data</h2>", report)
        self.assertIn("<h2>Berekeningen</h2>", report)
        self.assertIn("<h3>Opmaak</h3>", report)
        self.assertIn("<h4>Lettertype</h4>", report)
        self.assertIn("<h4>Getalnotatie</h4>", report)
        self.assertIn("<td>lettertype</td>", report)
        self.assertIn("<td>getalnotatie (1 cel; bereiken: B2)</td>", report)
        self.assertNotIn("<td>Opmaak: lettertype</td>", report)
        self.assertIn("<h3>Invoer data (inhoud)</h3>", report)
        self.assertIn("<h3>Functies / berekeningen</h3>", report)
        self.assertLess(report.index("<h3>Opmaak</h3>"),
                        report.index("<h3>Invoer data (inhoud)</h3>"))
        self.assertIn("</section><hr><section", report)
        self.assertNotIn("<th>Werkblad</th>", report)

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
        self.assertRegex(difference.expected, r"#[0-9A-F]{6}")
        self.assertRegex(difference.actual, r"#[0-9A-F]{6}")
        self.assertNotIn("themakleur", difference.expected + difference.actual)
        self.assertNotIn("tint", difference.expected + difference.actual)
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
