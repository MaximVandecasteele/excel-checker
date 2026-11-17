import json
import threading
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from Excel.excel_checker import (
    ComparisonResult,
    Difference,
    compare_workbooks,
    write_report,
)
from Shared.report_server import make_report_server


def test_differences_on_renamed_sheet_use_reference_name_for_scoring(tmp_path):
    reference_path = tmp_path / "reference.xlsx"
    student_path = tmp_path / "student.xlsx"
    report_path = tmp_path / "report.html"

    reference = Workbook()
    reference_sheet = reference.active
    reference_sheet.title = "Hype O'Dream 2025"
    reference_sheet["A1"] = "Input"
    reference_sheet["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor="FFFF0000",
    )
    reference_sheet["B1"] = "=1+1"
    reference.save(reference_path)

    student = Workbook()
    student.active.title = "Hype'O Dream 2025"
    student.save(student_path)

    result = compare_workbooks(reference_path, student_path)

    assert result.error is None
    assert result.differences
    assert {difference.sheet for difference in result.differences} == {
        "Hype'O Dream 2025"
    }
    assert any(
        difference.category == "Functies / berekeningen"
        and difference.item == "B1 inhoud/formule"
        for difference in result.differences
    )

    write_report(
        result,
        report_path,
        {
            "Hype O'Dream 2025": {
                "Opmaak": 2,
                "Functies / berekeningen": 4,
                "Invoer data (inhoud)": 2,
            }
        },
    )

    report = report_path.read_text(encoding="utf-8")
    assert "verschil(len) gevonden" in report
    assert (
        'aria-label="Hype O&#x27;Dream 2025 — Functies / berekeningen behaalde punten"'
        in report
    )
    assert "B1 inhoud/formule" in report
    assert 'data-auto-award="true"' not in report


def test_auto_awarded_scores_are_not_reused_after_report_changes(tmp_path):
    report_path = tmp_path / "student.html"
    scoring_parts = {"Data": {"Opmaak": 2}}
    write_report(
        ComparisonResult(tmp_path / "student.xlsx", []),
        report_path,
        scoring_parts,
    )
    original_revision = report_path.read_text(encoding="utf-8").split(
        'data-score-revision="', 1
    )[1].split('"', 1)[0]

    server = make_report_server(tmp_path)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base_url = f"http://127.0.0.1:{server.server_port}"
    try:
        request = Request(
            f"{base_url}/api/scores/{report_path.name}",
            data=json.dumps({
                "scores": ["2"],
                "revision": original_revision,
            }).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urlopen(request) as response:
            assert json.load(response) == {"saved": True}

        write_report(
            ComparisonResult(
                tmp_path / "student.xlsx",
                [Difference("Data", "Opmaak: lettertype", "vet", "normaal", "Opmaak")],
            ),
            report_path,
            scoring_parts,
        )
        changed_revision = report_path.read_text(encoding="utf-8").split(
            'data-score-revision="', 1
        )[1].split('"', 1)[0]
        assert changed_revision != original_revision

        with urlopen(f"{base_url}/api/scores/{report_path.name}") as response:
            assert json.load(response) == {"scores": [], "complete": False}

        stale_request = Request(
            f"{base_url}/api/scores/{report_path.name}",
            data=json.dumps({
                "scores": ["2"],
                "revision": original_revision,
            }).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            urlopen(stale_request)
        except HTTPError as error:
            assert error.code == 409
        else:
            raise AssertionError("A score for an outdated report was accepted.")
    finally:
        server.shutdown()
        server.server_close()
        thread.join()
