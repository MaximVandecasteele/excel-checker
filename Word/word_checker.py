from __future__ import annotations

import difflib
import hashlib
import html
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from Shared.document_workflow import review_folder


GLOBAL_SCORING_PARTS = (
    "Automatische nummering",
    "Inhoudsopgave",
    "Pagina-eindes",
)


@dataclass(frozen=True)
class Finding:
    item: str
    expected: str
    actual: str


@dataclass
class ComparisonResult:
    student_file: Path
    sections: dict[str, list[Finding]] = field(default_factory=dict)
    summaries: dict[str, str] = field(default_factory=dict)
    error: str | None = None

    @property
    def matches(self) -> bool:
        return not any(self.sections.values()) and self.error is None


def _document_paragraphs(document: Any) -> list[Paragraph]:
    return [
        Paragraph(element, document)
        for element in document._element.body.iter(qn("w:p"))
    ]


def _style_name(style: Any) -> str:
    return style.name if style is not None else "Geen stijl"


def _used_style_names(document: Any) -> list[str]:
    names: list[str] = []
    for paragraph in _document_paragraphs(document):
        run_style_names = [
            _style_name(run.style)
            for run in paragraph.runs
            if run._r.rPr is not None and run._r.rPr.rStyle is not None
        ]
        for name in [_style_name(paragraph.style), *run_style_names]:
            if name not in names:
                names.append(name)
    return names or ["Normal"]


def reference_style_names(path: Path) -> list[str]:
    document = Document(path)
    return _used_style_names(document)


def _numbering_signature(paragraph: Paragraph) -> tuple[Any, ...] | None:
    num_properties = paragraph._p.pPr
    num_properties = num_properties.numPr if num_properties is not None else None
    if num_properties is None:
        style = paragraph.style
        while style is not None:
            style_properties = style._element.pPr
            num_properties = (
                style_properties.numPr if style_properties is not None else None
            )
            if num_properties is not None:
                break
            style = style.base_style
    if num_properties is None or num_properties.numId is None:
        return None

    numbering = paragraph.part.numbering_part.element
    num_id = num_properties.numId.val
    level = num_properties.ilvl.val if num_properties.ilvl is not None else 0
    number = next(
        (item for item in numbering.findall(qn("w:num"))
         if item.get(qn("w:numId")) == str(num_id)),
        None,
    )
    if number is None:
        return ("onbekende nummering", level)
    abstract_id = number.find(qn("w:abstractNumId"))
    abstract = next(
        (item for item in numbering.findall(qn("w:abstractNum"))
         if abstract_id is not None
         and item.get(qn("w:abstractNumId")) == abstract_id.get(qn("w:val"))),
        None,
    )
    level_definition = None
    if abstract is not None:
        level_definition = next(
            (item for item in abstract.findall(qn("w:lvl"))
             if item.get(qn("w:ilvl")) == str(level)),
            None,
        )
    if level_definition is None:
        return ("nummering zonder niveaudefinitie", level)

    def value(tag: str) -> str | None:
        element = level_definition.find(qn(f"w:{tag}"))
        return element.get(qn("w:val")) if element is not None else None

    override = next(
        (item for item in number.findall(qn("w:lvlOverride"))
         if item.get(qn("w:ilvl")) == str(level)),
        None,
    )
    start_override = None
    if override is not None:
        override_level = override.find(qn("w:lvl"))
        if override_level is not None:
            level_definition = override_level
        start = override.find(qn("w:startOverride"))
        start_override = start.get(qn("w:val")) if start is not None else None
    return (
        level, value("start"), start_override, value("numFmt"), value("lvlText"),
        value("lvlJc"), value("suff"), value("lvlRestart"),
    )


def _style_signature(
    style: Any, seen: frozenset[str] = frozenset()
) -> tuple[Any, ...]:
    if style.style_id in seen:
        return ("style inheritance cycle", style.style_id)
    font = style.font
    color = font.color
    font_properties = (
        font.name, font.size.pt if font.size is not None else None, font.bold,
        font.italic, font.underline, font.strike, font.all_caps, font.small_caps,
        font.subscript, font.superscript,
        str(color.rgb) if color is not None and color.rgb is not None else None,
        str(color.theme_color) if color is not None and color.theme_color is not None else None,
        getattr(color, "theme_brightness", None) if color is not None else None,
    )
    paragraph_properties: tuple[Any, ...] = ()
    if style.type == WD_STYLE_TYPE.PARAGRAPH:
        paragraph = style.paragraph_format
        tabs = tuple(
            (str(tab.position), str(tab.alignment), str(tab.leader))
            for tab in paragraph.tab_stops
        )
        paragraph_properties = (
            str(paragraph.alignment) if paragraph.alignment is not None else None,
            str(paragraph.left_indent), str(paragraph.right_indent),
            str(paragraph.first_line_indent), str(paragraph.space_before),
            str(paragraph.space_after), str(paragraph.line_spacing),
            paragraph.keep_together, paragraph.keep_with_next,
            paragraph.page_break_before, paragraph.widow_control, tabs,
        )
    base_style = style.base_style
    return (
        str(style.type), _style_name(base_style),
        _style_signature(base_style, seen | {style.style_id}) if base_style else None,
        font_properties,
        paragraph_properties,
    )


def _style_definitions(document: Any, names: list[str]) -> dict[str, tuple[Any, ...] | None]:
    result: dict[str, tuple[Any, ...] | None] = {}
    for name in names:
        style = next((item for item in document.styles if item.name == name), None)
        result[name] = _style_signature(style) if style is not None else None
    return result


def _has_generated_toc(document: Any) -> bool:
    body = document._element.body
    return any(
        re.search(r"\bTOC\b", instruction.text or "", flags=re.IGNORECASE)
        for instruction in body.iter(qn("w:instrText"))
    ) or any(
        re.search(r"\bTOC\b", field.get(qn("w:instr")) or "", flags=re.IGNORECASE)
        for field in body.iter(qn("w:fldSimple"))
    )


def _page_breaks(paragraphs: list[Paragraph]) -> tuple[tuple[int, str], ...]:
    breaks: list[tuple[int, str]] = []
    for index, paragraph in enumerate(paragraphs):
        page_break_before = paragraph.paragraph_format.page_break_before
        style = paragraph.style
        while page_break_before is None and style is not None:
            page_break_before = style.paragraph_format.page_break_before
            style = style.base_style
        if page_break_before:
            breaks.append((index, "begin van alinea"))
        for run in paragraph.runs:
            for element in run._r.iter(qn("w:br")):
                if element.get(qn("w:type")) == "page":
                    breaks.append((index, "in alinea"))
    return tuple(breaks)


def _section_breaks(document: Any, paragraphs: list[Paragraph]) -> tuple[tuple[int, str], ...]:
    breaks = []
    for index, paragraph in enumerate(paragraphs):
        properties = paragraph._p.pPr
        section_properties = properties.sectPr if properties is not None else None
        if section_properties is not None:
            section_type = section_properties.find(qn("w:type"))
            breaks.append(
                (index, section_type.get(qn("w:val")) if section_type is not None else "nextPage")
            )
    final_section = document._element.body.sectPr
    if final_section is not None:
        section_type = final_section.find(qn("w:type"))
        breaks.append(
            (
                len(paragraphs),
                section_type.get(qn("w:val")) if section_type is not None else "nextPage",
            )
        )
    return tuple(breaks)


def _add(
    result: ComparisonResult,
    section: str,
    item: str,
    expected: Any,
    actual: Any,
) -> None:
    result.sections.setdefault(section, []).append(
        Finding(item, str(expected), str(actual))
    )


def compare_documents(reference_path: Path, student_path: Path) -> ComparisonResult:
    result = ComparisonResult(student_path)
    try:
        reference = Document(reference_path)
        student = Document(student_path)
    except Exception as exc:
        result.error = f"Bestand kon niet worden gelezen: {exc}"
        return result

    reference_paragraphs = _document_paragraphs(reference)
    student_paragraphs = _document_paragraphs(student)
    reference_styles = _used_style_names(reference)
    result.sections = {
        **{f"Stijl: {name}": [] for name in reference_styles},
        **{name: [] for name in GLOBAL_SCORING_PARTS},
    }
    student_styles = _used_style_names(student)
    reference_definitions = _style_definitions(reference, reference_styles)
    student_definitions = _style_definitions(student, reference_styles)

    for style_name in reference_styles:
        section = f"Stijl: {style_name}"
        if style_name not in student_styles:
            _add(result, section, "Stijl gebruikt", "aanwezig", "ontbreekt")
        if reference_definitions[style_name] != student_definitions[style_name]:
            _add(
                result, section, "Stijldefinitie",
                reference_definitions[style_name] or "ontbreekt",
                student_definitions[style_name] or "ontbreekt",
            )

    for style_name in student_styles:
        if style_name not in reference_styles:
            _add(
                result, f"Stijl: {style_name}", "Extra stijl gebruikt",
                "niet gebruikt in model", "gebruikt in leerlingbestand",
            )

    for index in range(max(len(reference_paragraphs), len(student_paragraphs))):
        expected = reference_paragraphs[index] if index < len(reference_paragraphs) else None
        actual = student_paragraphs[index] if index < len(student_paragraphs) else None
        section = f"Stijl: {_style_name(expected.style) if expected is not None else '[extra alinea]'}"
        location = f"Alinea {index + 1}"
        if expected is None or actual is None:
            paragraph = expected if expected is not None else actual
            section = f"Stijl: {_style_name(paragraph.style)}"
            is_empty = paragraph is not None and not paragraph.text.strip()
            item = (
                f"{location} — lege alinea ontbreekt of is extra"
                if is_empty
                else f"{location} ontbreekt of is extra"
            )
            _add(
                result, section, item,
                (
                    "lege alinea ontbreekt" if expected is not None and is_empty
                    else expected.text if expected is not None else "geen alinea"
                ),
                (
                    "lege alinea toegevoegd" if actual is not None and is_empty
                    else actual.text if actual is not None else "geen alinea"
                ),
            )
            continue
        if expected.text != actual.text:
            _add(result, section, f"{location} — tekenverschil", expected.text, actual.text)
        expected_style = _style_name(expected.style)
        actual_style = _style_name(actual.style)
        if expected_style != actual_style:
            _add(
                result, f"Stijl: {expected_style}",
                f"{location} — toegepaste stijl", expected_style, actual_style,
            )

        expected_numbering = _numbering_signature(expected)
        actual_numbering = _numbering_signature(actual)
        if expected_numbering != actual_numbering:
            _add(
                result, "Automatische nummering", location,
                expected_numbering or "geen automatische nummering",
                actual_numbering or "geen automatische nummering",
            )

        expected_double_spaces = tuple(
            len(match.group()) for match in re.finditer(r" {2,}", expected.text)
        )
        actual_double_spaces = tuple(
            len(match.group()) for match in re.finditer(r" {2,}", actual.text)
        )
        if expected_double_spaces != actual_double_spaces:
            _add(
                result, section, f"{location} — dubbele spaties",
                expected_double_spaces, actual_double_spaces,
            )
        if (not expected.text.strip()) != (not actual.text.strip()):
            _add(
                result, section, f"{location} — lege alinea",
                "leeg" if not expected.text.strip() else "niet leeg",
                "leeg" if not actual.text.strip() else "niet leeg",
            )

    reference_toc = _has_generated_toc(reference)
    student_toc = _has_generated_toc(student)
    result.summaries["Inhoudsopgave"] = (
        f"Model: {'aanwezig' if reference_toc else 'afwezig'}; "
        f"leerling: {'aanwezig' if student_toc else 'afwezig'}."
    )
    if reference_toc != student_toc:
        _add(
            result, "Inhoudsopgave", "Automatisch gegenereerde inhoudsopgave",
            "aanwezig" if reference_toc else "afwezig",
            "aanwezig" if student_toc else "afwezig",
        )

    reference_breaks = _page_breaks(reference_paragraphs)
    student_breaks = _page_breaks(student_paragraphs)
    reference_section_breaks = _section_breaks(reference, reference_paragraphs)
    student_section_breaks = _section_breaks(student, student_paragraphs)
    result.summaries["Pagina-eindes"] = (
        f"Pagina-eindes: model {reference_breaks or 'geen'}, leerling "
        f"{student_breaks or 'geen'}; sectie-eindes: model "
        f"{reference_section_breaks}, leerling {student_section_breaks}."
    )
    if reference_breaks != student_breaks:
        _add(
            result, "Pagina-eindes", "Pagina-eindes (alinea-posities)",
            reference_breaks or "geen", student_breaks or "geen",
        )
    if reference_section_breaks != student_section_breaks:
        _add(
            result, "Pagina-eindes", "Sectie-eindes (alinea-posities en type)",
            reference_section_breaks, student_section_breaks,
        )
    reference_numbering = [
        _numbering_signature(paragraph) for paragraph in reference_paragraphs
    ]
    student_numbering = [
        _numbering_signature(paragraph) for paragraph in student_paragraphs
    ]
    result.summaries["Automatische nummering"] = (
        f"Genummerde alinea's: model "
        f"{sum(item is not None for item in reference_numbering)}, leerling "
        f"{sum(item is not None for item in student_numbering)}."
    )

    return result


def _char_diff(expected: str, actual: str) -> str:
    parts: list[str] = []
    for operation, start_expected, end_expected, start_actual, end_actual in (
        difflib.SequenceMatcher(a=expected, b=actual, autojunk=False).get_opcodes()
    ):
        if operation == "equal":
            parts.append(html.escape(expected[start_expected:end_expected]))
        else:
            if operation in {"delete", "replace"}:
                parts.append(
                    f'<del>{html.escape(expected[start_expected:end_expected])}</del>'
                )
            if operation in {"insert", "replace"}:
                parts.append(
                    f'<ins>{html.escape(actual[start_actual:end_actual])}</ins>'
                )
    return "".join(parts) or '<span class="space">∅</span>'


def write_report(
    result: ComparisonResult,
    output_path: Path,
    scoring_parts: dict[str, float] | None = None,
) -> None:
    scoring_parts = scoring_parts or {}
    scored = [
        (section, maximum)
        for section, maximum in scoring_parts.items()
        if maximum > 0
    ]
    revision = hashlib.sha256(
        json.dumps(
            {
                "error": result.error,
                "sections": {
                    name: [(finding.item, finding.expected, finding.actual) for finding in findings]
                    for name, findings in result.sections.items()
                },
                "summaries": result.summaries,
                "scoring_parts": scoring_parts,
            },
            ensure_ascii=False,
            sort_keys=True,
        ).encode("utf-8")
    ).hexdigest()

    score_fields = []
    for index, (section, maximum) in enumerate(scored):
        automatic = result.error is None and not result.sections.get(section)
        value = f' value="{html.escape(str(maximum), quote=True)}"' if automatic else ""
        auto_attribute = ' data-auto-award="true"' if automatic else ""
        score_fields.append(
            f'<span class="score"><label for="score-{index}">Behaald</label> '
            f'<input id="score-{index}" type="number" min="0" '
            f'max="{html.escape(str(maximum), quote=True)}" step="any" '
            f'data-score-field="earned" data-maximum="{html.escape(str(maximum), quote=True)}"'
            f'{value}{auto_attribute}> / {html.escape(str(maximum))} punten</span>'
        )

    if result.error:
        detail = f'<p class="error">{html.escape(result.error)}</p>'
    else:
        sections = []
        section_names = list(dict.fromkeys([*scoring_parts, *result.sections]))
        for section in section_names:
            findings = result.sections.get(section, [])
            if not findings:
                sections.append(
                    f'<section><h2>{html.escape(section)}</h2>'
                    '<p class="ok">Geen verschillen gevonden. '
                    f'{html.escape(result.summaries.get(section, ""))}</p></section>'
                )
                continue
            lines = "".join(
                "<tr><th>" + html.escape(finding.item) + "</th><td>"
                + _char_diff(finding.expected, finding.actual)
                + '<details><summary>Toon waarden</summary><p>Model: <code>'
                + html.escape(finding.expected) + "</code></p><p>Leerling: <code>"
                + html.escape(finding.actual) + "</code></p></details></td></tr>"
                for finding in findings
            )
            sections.append(
                f"<section><h2>{html.escape(section)}</h2><table>"
                f"<thead><tr><th>Controle</th><th>Vergelijking (rood = model, groen = leerling)</th></tr></thead>"
                f"<tbody>{lines}</tbody></table></section>"
            )
        detail = "".join(sections)
        if result.matches:
            detail = (
                '<p class="ok">Geen verschillen gevonden; alle ingestelde punten zijn toegekend.</p>'
                + detail
            )

    score_section = ""
    if scored:
        score_rows = "".join(
            f"<li>{html.escape(section)} {field}</li>"
            for (section, _maximum), field in zip(scored, score_fields)
        )
        score_section = (
            f'<main class="grading" data-score-revision="{revision}" '
            f'data-score-file="{html.escape(output_path.name, quote=True)}">'
            f"<h2>Punten</h2><ul>{score_rows}</ul>"
            '<p>Totaal: <output id="score-total">0</output> punten</p>'
            '<p id="score-status" aria-live="polite"></p></main>'
        )

    document = f"""<!doctype html>
<html lang="nl"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>{html.escape(result.student_file.stem)} — Word Checker</title>
<style>
body{{font:16px/1.5 system-ui,sans-serif;max-width:1100px;margin:2rem auto;padding:0 1rem;color:#222}}
h1,h2{{line-height:1.2}} section{{margin:2rem 0}} table{{border-collapse:collapse;width:100%}}
th,td{{border:1px solid #ccc;padding:.6rem;text-align:left;vertical-align:top;white-space:pre-wrap}}
th{{background:#f2f4f7}} code{{white-space:pre-wrap;overflow-wrap:anywhere}}
del{{background:#ffd9d9;color:#8b0000;text-decoration:line-through}}
ins{{background:#d9f7df;color:#075b19;text-decoration:none}} .space{{color:#777}}
.ok{{padding:1rem;background:#e6f4ea}} .error{{padding:1rem;background:#fce8e6}}
.score input{{width:5rem}} details{{margin-top:.5rem}} main.grading{{border-top:2px solid #ddd}}
@media print{{details{{display:none}}}}
</style>
<h1>{html.escape(result.student_file.stem)}</h1>
<p>Word-document vergeleken met de modeloplossing.</p>
{score_section}{detail}
<script>
(() => {{
 const fields = [...document.querySelectorAll('[data-score-field="earned"]')];
 const total = document.getElementById("score-total");
 const status = document.getElementById("score-status");
 function update() {{
   let sum = 0, complete = true;
   fields.forEach(field => {{
     const max = Number(field.dataset.maximum);
     const value = Number(field.value);
     if (field.value === "" || !Number.isFinite(value) || value < 0 || value > max) complete = false;
     else sum += value;
   }});
   if (total) total.value = String(sum);
   if (status) status.textContent = complete ? "Alle scores zijn ingevuld." : "Niet ingevulde scores tellen als 0.";
 }}
 fields.forEach(field => field.addEventListener("input", update));
 update();
}})();
</script></html>"""
    output_path.write_text(document, encoding="utf-8")


def check_folder(
    reference_path: Path,
    student_folder: Path,
    output_folder: Path,
    scoring_parts: dict[str, float] | None = None,
    filename_suffix: str = "",
) -> list[Path]:
    return review_folder(
        reference_path,
        student_folder,
        output_folder,
        file_extension=".docx",
        compare=compare_documents,
        write_report=lambda result, destination: write_report(
            result, destination, scoring_parts
        ),
        filename_suffix=filename_suffix,
    )
