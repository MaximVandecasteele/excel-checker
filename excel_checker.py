from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import (
    get_column_letter,
    column_index_from_string,
    range_boundaries,
)
from openpyxl.formula.translate import Translator


@dataclass(frozen=True)
class Difference:
    sheet: str
    item: str
    expected: str
    actual: str


@dataclass
class ComparisonResult:
    student_file: Path
    differences: list[Difference]
    error: str | None = None

    @property
    def matches(self) -> bool:
        return not self.differences and self.error is None


def _color(color: Any) -> Any:
    if color is None:
        return None
    if color.type == "rgb":
        value = color.rgb
    elif color.type == "indexed":
        value = color.indexed
    elif color.type == "theme":
        value = color.theme
    elif color.type == "auto":
        value = color.auto
    else:
        value = None
    return (color.type, value, color.tint)


def _side(side: Any) -> Any:
    if side is None:
        return None
    color = _color(side.color)
    if side.style is None and color is None:
        return None
    return (
        side.style,
        color,
    )


def _normalize_number_format(number_format: str) -> str:
    normalized: list[str] = []
    index = 0
    quoted = False
    format_tokens = "0#?@;[]*_\\"
    while index < len(number_format):
        character = number_format[index]
        if character == '"':
            quoted = not quoted
            normalized.append(character)
        elif character == "\\" and index + 1 < len(number_format):
            escaped = number_format[index + 1]
            if quoted or escaped in format_tokens:
                normalized.extend((character, escaped))
            else:
                normalized.append(escaped)
            index += 1
        else:
            normalized.append(character)
        index += 1
    return "".join(normalized)


def _style(cell: Any) -> dict[str, Any]:
    font = cell.font
    fill = cell.fill
    alignment = cell.alignment
    border = cell.border
    protection = cell.protection
    return {
        "lettertype": (
            font.name, font.sz, font.b, font.i, font.u, font.strike,
            _color(font.color), font.vertAlign, font.outline, font.shadow,
            font.charset, font.family, font.scheme,
        ),
        "opvulling": (
            fill.fill_type, _color(fill.fgColor), _color(fill.bgColor),
            fill.patternType,
        ),
        "rand": (
            _side(border.left), _side(border.right), _side(border.top),
            _side(border.bottom), _side(border.diagonal),
            border.diagonalUp, border.diagonalDown, border.outline,
        ),
        "uitlijning": (
            alignment.horizontal, alignment.vertical, alignment.textRotation,
            alignment.wrapText, alignment.shrinkToFit, alignment.indent,
            alignment.relativeIndent, alignment.justifyLastLine,
            alignment.readingOrder,
        ),
        "getalnotatie": _normalize_number_format(cell.number_format),
        "celbeveiliging": (protection.locked, protection.hidden),
    }


def _cell_value(cell: Any) -> tuple[Any, str]:
    value = cell.value
    if value is None:
        return None, "blank"
    if isinstance(value, str) and cell.data_type == "f":
        return value, "formule"
    return value, cell.data_type


def _formula_pattern(formula: str, coordinate: str) -> str:
    try:
        return Translator(formula, origin=coordinate).translate_formula("MZ500000")
    except Exception:
        return f"{coordinate}:{formula}"


def _formula_difference_key(
    expected_formula: str,
    actual_formula: str,
    coordinate: str,
) -> tuple[str, str]:
    return (
        _formula_pattern(expected_formula, coordinate),
        _formula_pattern(actual_formula, coordinate),
    )


def _coalesce_formula_differences(
    sheet: str,
    formula_changes: dict[tuple[int, int], tuple[str, str, tuple[str, str]]],
) -> list[Difference]:
    grouped: list[Difference] = []
    consumed: set[tuple[int, int]] = set()

    for row_first in (True, False):
        lines: dict[int, list[tuple[int, int]]] = {}
        for coordinate in formula_changes:
            row, column = coordinate
            line, position = (row, column) if row_first else (column, row)
            lines.setdefault(line, []).append((position, coordinate))

        candidates = []
        for entries in lines.values():
            entries.sort()
            run: list[tuple[int, int]] = []
            previous_position = None
            previous_key = None
            for position, coordinate in entries:
                key = formula_changes[coordinate][2]
                if previous_position is None or position != previous_position + 1:
                    if len(run) > 1:
                        candidates.append(run)
                    run = []
                if run and key != previous_key:
                    if len(run) > 1:
                        candidates.append(run)
                    run = []
                run.append(coordinate)
                previous_position = position
                previous_key = key
            if len(run) > 1:
                candidates.append(run)

        candidates.sort(key=len, reverse=True)
        for run in candidates:
            available = [coordinate for coordinate in run if coordinate not in consumed]
            if len(available) < 2:
                continue
            row, column = available[0]
            first = f"{get_column_letter(column)}{row}"
            row, column = available[-1]
            last = f"{get_column_letter(column)}{row}"
            expected_formula, actual_formula, _ = formula_changes[available[0]]
            direction = "rij" if row_first else "kolom"
            grouped.append(Difference(
                sheet,
                f"{first}:{last} inhoud/formule ({len(available)} cellen; doorgevoerde fout in {direction})",
                expected_formula,
                actual_formula,
            ))
            consumed.update(available)

    for coordinate, (expected_formula, actual_formula, _) in formula_changes.items():
        if coordinate in consumed:
            continue
        row, column = coordinate
        cell = f"{get_column_letter(column)}{row}"
        grouped.append(Difference(
            sheet,
            f"{cell} inhoud/formule",
            expected_formula,
            actual_formula,
        ))
    return grouped


def _short_value(value: Any, limit: int = 70) -> str:
    text = repr(value)
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _format_style_color(color: Any) -> str:
    if color is None:
        return "geen kleur"
    color_type, value, tint = color
    if color_type == "theme":
        description = f"themakleur {value}"
    elif color_type == "rgb" and isinstance(value, str):
        description = f"#{value[-6:]}"
    elif color_type == "indexed":
        description = f"geïndexeerde kleur {value}"
    else:
        description = str(value or color_type)
    if tint:
        description += f" (tint {tint:.2f})"
    return description


def _format_style_summary(changes: dict[str, tuple[Any, Any]], side: int) -> str:
    summaries = []
    for category, values in changes.items():
        value = values[side]
        if category == "opvulling":
            fill_type, foreground, background, _ = value
            description = "effen opvulling" if fill_type == "solid" else f"opvulling {fill_type}"
            color = foreground or background
            if color:
                description += f", {_format_style_color(color)}"
            summaries.append(description)
        else:
            summaries.append(f"{category}: {_short_value(value, 120)}")
    return "; ".join(summaries)


def _coalesce_cell_differences(
    sheet: str,
    changes: dict[tuple[int, int], tuple[Any, Any]],
    minimum_group_size: int = 20,
) -> list[Difference]:
    differences = []
    remaining = set(changes)
    while remaining:
        seed = min(remaining)
        remaining.remove(seed)
        component = {seed}
        pending = [seed]
        while pending:
            row, column = pending.pop()
            for neighbor in (
                (row - 1, column), (row + 1, column),
                (row, column - 1), (row, column + 1),
            ):
                if neighbor in remaining:
                    remaining.remove(neighbor)
                    component.add(neighbor)
                    pending.append(neighbor)

        if len(component) < minimum_group_size:
            for row, column in sorted(component):
                expected, actual = changes[(row, column)]
                coordinate = f"{get_column_letter(column)}{row}"
                _add_difference(
                    differences, sheet, f"{coordinate} inhoud/formule",
                    expected, actual,
                )
            continue

        min_row = min(row for row, _ in component)
        max_row = max(row for row, _ in component)
        min_column = min(column for _, column in component)
        max_column = max(column for _, column in component)
        first = f"{get_column_letter(min_column)}{min_row}"
        last = f"{get_column_letter(max_column)}{max_row}"
        ordered_component = sorted(component)
        sample_indexes = {
            round(index * (len(ordered_component) - 1) / 4)
            for index in range(min(5, len(ordered_component)))
        }
        samples = [ordered_component[index] for index in sorted(sample_indexes)]
        expected_samples = []
        actual_samples = []
        for row, column in samples:
            coordinate = f"{get_column_letter(column)}{row}"
            expected, actual = changes[(row, column)]
            expected_samples.append(f"{coordinate}: {_short_value(expected)}")
            actual_samples.append(f"{coordinate}: {_short_value(actual)}")
        differences.append(Difference(
            sheet,
            f"Inhoud/formules in bereik {first}:{last} "
            f"({len(component)} afwijkende cellen)",
            "Voorbeelden: " + "; ".join(expected_samples),
            "Voorbeelden: " + "; ".join(actual_samples),
        ))
    return differences


def _coalesce_style_differences(
    sheet: str,
    style_changes: dict[tuple[int, int], dict[str, tuple[Any, Any]]],
) -> list[Difference]:
    by_signature: dict[tuple[str, str, str], list[tuple[int, int]]] = {}
    signature_values: dict[tuple[str, str, str], tuple[Any, Any]] = {}
    for coordinate, changes in style_changes.items():
        for category, (expected, actual) in changes.items():
            signature = (category, repr(expected), repr(actual))
            by_signature.setdefault(signature, []).append(coordinate)
            signature_values[signature] = (expected, actual)

    differences = []
    for signature, coordinates in sorted(by_signature.items()):
        rows: dict[int, list[int]] = {}
        for row, column in coordinates:
            rows.setdefault(row, []).append(column)

        horizontal_ranges = []
        for row, columns in sorted(rows.items()):
            columns.sort()
            start = previous = columns[0]
            for column in columns[1:]:
                if column != previous + 1:
                    horizontal_ranges.append((row, start, row, previous))
                    start = column
                previous = column
            horizontal_ranges.append((row, start, row, previous))

        compact_ranges = []
        for row_start, column_start, row_end, column_end in horizontal_ranges:
            if (
                compact_ranges
                and compact_ranges[-1][1] == column_start
                and compact_ranges[-1][3] == column_end
                and compact_ranges[-1][2] + 1 == row_start
            ):
                compact_ranges[-1][2] = row_end
            else:
                compact_ranges.append([row_start, column_start, row_end, column_end])

        def address(row: int, column: int) -> str:
            return f"{get_column_letter(column)}{row}"

        range_labels = []
        for row_start, column_start, row_end, column_end in compact_ranges:
            first = address(row_start, column_start)
            last = address(row_end, column_end)
            range_labels.append(first if first == last else f"{first}:{last}")
        visible_ranges = range_labels[:8]
        if len(range_labels) > 8:
            visible_ranges.append(f"… en {len(range_labels) - 8} andere bereiken")

        category = signature[0]
        expected, actual = signature_values[signature]
        count_label = "cel" if len(coordinates) == 1 else "cellen"
        differences.append(Difference(
            sheet,
            f"Opmaak: {category} ({len(coordinates)} {count_label}; "
            f"bereiken: {', '.join(visible_ranges)})",
            _format_style_summary({category: (expected, actual)}, 0),
            _format_style_summary({category: (expected, actual)}, 1),
        ))
    return differences


def _cell_annotations(cell: Any) -> tuple[Any, Any]:
    hyperlink = cell.hyperlink
    comment = cell.comment
    return (
        None if hyperlink is None else (
            hyperlink.target, hyperlink.location, hyperlink.tooltip,
            hyperlink.display, hyperlink.ref,
        ),
        None if comment is None else (comment.text, comment.author),
    )


def _normalized_xml(node: Any, ignored_attributes: frozenset[str] = frozenset()) -> Any:
    return (
        node.tag,
        tuple(sorted(
            (key, value) for key, value in node.attrib.items()
            if key.rsplit("}", 1)[-1] not in ignored_attributes
        )),
        (node.text or "").strip(),
        tuple(_normalized_xml(child, ignored_attributes) for child in node),
    )


def _normalize_chart_formula(formula: str) -> str:
    return re.sub(
        r"'([A-Za-z_][A-Za-z0-9_.]*)'!",
        r"\1!",
        formula,
    )


def _title_signature(title: Any) -> Any:
    if title is None or title.tx is None:
        return None
    if title.tx.strRef is not None:
        return ("cell", _normalize_chart_formula(title.tx.strRef.f))
    return ("text", _chart_title_from_title(title))


def _chart_title_from_title(title: Any) -> str | None:
    if title is None or title.tx is None or title.tx.rich is None:
        return None
    return "".join(
        run.t or ""
        for paragraph in (title.tx.rich.p or [])
        for run in (paragraph.r or [])
    ) or None


def _series_title_signature(series: Any) -> Any:
    if series.tx is None:
        return None
    if series.tx.strRef is not None:
        return ("cell", _normalize_chart_formula(series.tx.strRef.f))
    return ("text", series.tx.v)


def _axis_signature(axis: Any) -> Any:
    if axis is None:
        return None
    scaling = getattr(axis, "scaling", None)
    number_format = getattr(axis, "numFmt", None)
    crosses = getattr(axis, "crosses", None)
    if crosses in (None, "autoZero"):
        crosses = "autoZero"
    tick_label_position = getattr(axis, "tickLblPos", None)
    if tick_label_position in (None, "nextTo"):
        tick_label_position = "nextTo"
    return (
        type(axis).__name__,
        getattr(axis, "axPos", None),
        _title_signature(getattr(axis, "title", None)),
        getattr(scaling, "orientation", None),
        getattr(scaling, "logBase", None),
        getattr(scaling, "min", None),
        getattr(scaling, "max", None),
        getattr(number_format, "formatCode", None),
        getattr(axis, "majorUnit", None),
        getattr(axis, "minorUnit", None),
        getattr(axis, "majorTickMark", None),
        getattr(axis, "minorTickMark", None),
        tick_label_position,
        getattr(axis, "delete", None),
        crosses,
        getattr(axis, "crossesAt", None),
        getattr(axis, "majorGridlines", None) is not None,
        getattr(axis, "minorGridlines", None) is not None,
    )


def _data_labels_signature(chart: Any) -> Any:
    labels = getattr(chart, "dLbls", None)
    if labels is None:
        return None
    flags = (
        "showLegendKey", "showVal", "showCatName", "showSerName",
        "showPercent", "showBubbleSize",
    )
    enabled = tuple(name for name in flags if getattr(labels, name, False))
    if not enabled:
        return None
    return (
        enabled,
        getattr(labels, "dLblPos", None),
        getattr(getattr(labels, "numFmt", None), "formatCode", None),
        getattr(labels, "separator", None),
    )


def _chart_signature(chart: Any) -> Any:
    axes = tuple(
        _axis_signature(getattr(chart, axis_name, None))
        for axis_name in ("x_axis", "y_axis", "z_axis")
    )
    series = tuple(
        (
            _series_title_signature(item),
            _series_reference(item, "val"),
            _series_reference(item, "cat"),
            _series_reference(item, "xVal"),
            _series_reference(item, "yVal"),
            _series_reference(item, "zVal"),
            _object_signature(item.graphicalProperties)
            if item.graphicalProperties is not None else None,
        )
        for item in chart.series
    )
    options = tuple(
        (name, getattr(chart, name))
        for name in (
            "barDir", "grouping", "overlap", "gapWidth", "scatterStyle",
            "firstSliceAng", "holeSize", "smooth", "wireframe",
            "varyColors", "style",
        )
        if hasattr(chart, name)
    )
    legend = None if chart.legend is None else (
        chart.legend.legendPos,
        chart.legend.overlay,
    )
    data_labels = _data_labels_signature(chart)
    return (
        type(chart).__name__,
        _title_signature(chart.title),
        legend,
        options,
        series,
        axes,
        data_labels,
    )


def _chart_title(chart: Any) -> str | None:
    signature = _title_signature(chart.title)
    if signature is None:
        return None
    if signature[0] == "cell":
        return f"celverwijzing {signature[1]}"
    return signature[1]


def _series_reference(series: Any, axis_name: str) -> str | None:
    source = getattr(series, axis_name, None)
    if source is None:
        return None
    for reference_name in ("numRef", "strRef", "multiLvlStrRef"):
        reference = getattr(source, reference_name, None)
        if reference is not None and reference.f:
            return _normalize_chart_formula(reference.f)
    return None


def _chart_description(chart: Any) -> str:
    chart_types = {
        "BarChart": "kolom-/staafgrafiek",
        "LineChart": "lijngrafiek",
        "PieChart": "cirkeldiagram",
        "DoughnutChart": "ringdiagram",
        "AreaChart": "vlakgrafiek",
        "ScatterChart": "spreidingsdiagram",
        "RadarChart": "radardiagram",
        "BubbleChart": "bellendiagram",
        "SurfaceChart": "oppervlaktegrafiek",
    }
    parts = [chart_types.get(type(chart).__name__, type(chart).__name__)]
    title = _chart_title(chart)
    if title:
        parts.append(f'titel "{title}"')
    series_count = len(chart.series)
    noun = "gegevensreeks" if series_count == 1 else "gegevensreeksen"
    parts.append(f"{series_count} {noun}")
    if hasattr(chart, "barDir"):
        directions = {"col": "kolommen", "bar": "balken"}
        parts.append(f"richting {directions.get(chart.barDir, chart.barDir)}")
    if hasattr(chart, "grouping"):
        groupings = {
            "clustered": "geclusterd",
            "stacked": "gestapeld",
            "percentStacked": "100% gestapeld",
            "standard": "standaard",
        }
        parts.append(f"groepering {groupings.get(chart.grouping, chart.grouping)}")
    if chart.style is not None:
        parts.append(f"stijl {chart.style}")
    for axis_name in ("x_axis", "y_axis", "z_axis"):
        axis = getattr(chart, axis_name, None)
        axis_title = _chart_title_from_title(getattr(axis, "title", None))
        if axis_title:
            parts.append(f"as {axis_name[0]}-titel \"{axis_title}\"")
    sources = []
    for series in chart.series:
        values = _series_reference(series, "val") or _series_reference(series, "yVal")
        categories = _series_reference(series, "cat") or _series_reference(series, "xVal")
        if values:
            sources.append(f"waarden {values}")
        if categories:
            sources.append(f"categorieën {categories}")
    if sources:
        parts.extend(sources)
    if chart.legend is not None and chart.legend.legendPos:
        legend_positions = {
            "b": "onderaan",
            "t": "bovenaan",
            "l": "links",
            "r": "rechts",
            "tr": "rechtsboven",
        }
        parts.append(f"legenda {legend_positions.get(chart.legend.legendPos, chart.legend.legendPos)}")
    return "; ".join(parts)


def _object_signature(obj: Any) -> Any:
    node = obj.to_tree()
    return _normalized_xml(ET.fromstring(ET.tostring(node)))


def _conditional_rule_signature(rule: Any) -> Any:
    node = ET.fromstring(ET.tostring(rule.to_tree()))
    dxf = _object_signature(rule.dxf) if rule.dxf is not None else None
    return (
        _normalized_xml(node, frozenset({"priority", "dxfId"})),
        dxf,
    )


def _format_cf_color(color: Any) -> str | None:
    if color is None:
        return None
    value = _color(color)
    if value and value[0] == "rgb" and isinstance(value[1], str):
        return f"#{value[1][-6:]}"
    if value and value[0] in ("theme", "indexed"):
        return f"{value[0]} {value[1]}"
    return None


def _format_cf_rule(cell_range: str, rule: Any) -> str:
    type_names = {
        "cellIs": "celwaarderegel",
        "expression": "formuleregel",
        "dataBar": "databalk",
        "colorScale": "kleurenschaal",
        "iconSet": "pictogrammenset",
    }
    parts = [f"{cell_range} — {type_names.get(rule.type, rule.type or 'regel')}"]
    if rule.formula:
        parts.append("voorwaarde: " + ", ".join(str(item) for item in rule.formula))
    elif rule.operator:
        operator_names = {
            "equal": "gelijk aan",
            "notEqual": "niet gelijk aan",
            "greaterThan": "groter dan",
            "greaterThanOrEqual": "groter dan of gelijk aan",
            "lessThan": "kleiner dan",
            "lessThanOrEqual": "kleiner dan of gelijk aan",
            "between": "tussen",
            "notBetween": "niet tussen",
        }
        parts.append(operator_names.get(rule.operator, rule.operator))

    if rule.dataBar is not None:
        color = _format_cf_color(rule.dataBar.color)
        if color:
            parts.append(f"balkkleur {color}")
    elif rule.colorScale is not None:
        colors = [_format_cf_color(color) for color in rule.colorScale.color]
        colors = [color for color in colors if color]
        if colors:
            parts.append("kleuren " + " → ".join(colors))
    elif rule.iconSet is not None:
        parts.append(f"set {rule.iconSet.iconSet}")

    dxf = rule.dxf
    if dxf is not None:
        formatting = []
        if dxf.font is not None:
            if dxf.font.b:
                formatting.append("vet")
            if dxf.font.i:
                formatting.append("cursief")
            color = _format_cf_color(dxf.font.color)
            if color:
                formatting.append(f"tekstkleur {color}")
        if dxf.fill is not None:
            color = _format_cf_color(dxf.fill.fgColor)
            if color and dxf.fill.fill_type:
                formatting.append(f"achtergrond {color}")
        if formatting:
            parts.append("opmaak: " + ", ".join(formatting))
    return "; ".join(parts)


def _conditional_formatting_rules(sheet: Any) -> list[tuple[Any, str]]:
    rules = []
    for cell_range, conditional_rules in sheet.conditional_formatting._cf_rules.items():
        range_name = str(getattr(cell_range, "sqref", cell_range))
        for rule in conditional_rules:
            rules.append((
                (range_name, _conditional_rule_signature(rule)),
                _format_cf_rule(range_name, rule),
            ))
    return rules


def _dimension_signature(dim: Any, fields: tuple[str, ...]) -> tuple[Any, ...]:
    return tuple(getattr(dim, name, None) for name in fields)


def _rounded_size(value: Any) -> Any:
    return round(value, 1) if isinstance(value, (int, float)) else value


def _rounded_row_height(value: Any) -> Any:
    return round(value * 2) / 2 if isinstance(value, (int, float)) else value


def _row_dimension_signatures(
    sheet: Any,
    ignored_hidden_rows: set[int] | None = None,
) -> dict[int, tuple[Any, ...]]:
    signatures = {}
    default_height = sheet.sheet_format.defaultRowHeight
    ignored_hidden_rows = ignored_hidden_rows or set()
    for index, dimension in sheet.row_dimensions.items():
        height = dimension.height
        if (
            height is not None
            and abs(height - (default_height or 15.0)) < 0.75
        ):
            height = None
        signature = (
            _rounded_row_height(height),
            bool(dimension.hidden) if index not in ignored_hidden_rows else False,
            dimension.outlineLevel or 0,
            bool(dimension.collapsed),
        )
        if signature != (None, False, 0, False):
            signatures[index] = signature
    return signatures


def _column_dimension_signatures(sheet: Any) -> dict[int, tuple[Any, ...]]:
    signatures: dict[int, tuple[Any, ...]] = {}
    default_width = sheet.sheet_format.defaultColWidth
    if default_width is None:
        default_width = 8.7109375
    for key, dimension in sheet.column_dimensions.items():
        start = dimension.min or column_index_from_string(key)
        end = dimension.max or start
        width = dimension.width
        if (
            width is not None
            and abs(width - default_width) < 0.1
            and not dimension.hidden
            and not (dimension.outlineLevel or 0)
            and not dimension.collapsed
            and not dimension.bestFit
        ):
            continue
        signature = (
            _rounded_size(width),
            bool(dimension.hidden),
            dimension.outlineLevel or 0,
            bool(dimension.collapsed),
            bool(dimension.bestFit),
        )
        for index in range(start, end + 1):
            signatures[index] = signature
    return signatures


def _format_dimension_differences(
    expected: dict[int, tuple[Any, ...]],
    actual: dict[int, tuple[Any, ...]],
    label: str,
    format_index: Any,
) -> tuple[str, str]:
    changes = []
    for index in sorted(expected.keys() | actual.keys()):
        expected_value = expected.get(index)
        actual_value = actual.get(index)
        if expected_value == actual_value:
            continue
        if changes:
            previous = changes[-1]
            if (
                previous[1] + 1 == index
                and previous[2] == expected_value
                and previous[3] == actual_value
            ):
                previous[1] = index
                continue
        changes.append([index, index, expected_value, actual_value])

    def format_side(value: tuple[Any, ...] | None) -> str:
        if value is None:
            return "standaard"
        size = value[0]
        parts = []
        if size is not None:
            parts.append(f"{label} {size}")
        if value[1]:
            parts.append("verborgen")
        if value[2]:
            parts.append(f"overzichtsniveau {value[2]}")
        if value[3]:
            parts.append("samengevouwen")
        if len(value) > 4 and value[4]:
            parts.append("best passend")
        return ", ".join(parts) or "standaard"

    expected_summary = []
    actual_summary = []
    for start, end, expected_value, actual_value in changes:
        extent = format_index(start, end)
        expected_summary.append(f"{extent}: {format_side(expected_value)}")
        actual_summary.append(f"{extent}: {format_side(actual_value)}")
    return "; ".join(expected_summary), "; ".join(actual_summary)


def _format_row_differences(
    expected: dict[int, tuple[Any, ...]],
    actual: dict[int, tuple[Any, ...]],
) -> tuple[str, str]:
    return _format_dimension_differences(
        expected,
        actual,
        "hoogte",
        lambda start, end: f"Rij {start}" if start == end else f"Rijen {start}–{end}",
    )


def _format_column_differences(
    expected: dict[int, tuple[Any, ...]],
    actual: dict[int, tuple[Any, ...]],
) -> tuple[str, str]:
    return _format_dimension_differences(
        expected,
        actual,
        "breedte",
        lambda start, end: (
            get_column_letter(start)
            if start == end
            else f"{get_column_letter(start)}–{get_column_letter(end)}"
        ),
    )


def _table_signatures(sheet: Any) -> dict[str, Any]:
    return {
        name: (
            table.ref,
            table.displayName,
            table.tableStyleInfo.name if table.tableStyleInfo else None,
            tuple((column.id, column.name) for column in table.tableColumns),
        )
        for name in sheet.tables
        for table in (sheet.tables[name],)
    }


def _filter_objects(sheet: Any) -> list[tuple[str, Any]]:
    filters = []
    auto_filter = sheet.auto_filter
    if auto_filter.ref or auto_filter.filterColumn or auto_filter.sortState:
        filters.append(("werkblad", auto_filter))
    for name in sheet.tables:
        table_filter = sheet.tables[name].autoFilter
        if table_filter and (
            table_filter.ref or table_filter.filterColumn or table_filter.sortState
        ):
            filters.append((name, table_filter))
    return filters


def _filter_criteria(sheet: Any, filter_object: Any) -> tuple[Any, ...]:
    if not filter_object.ref:
        return ()
    min_col, min_row, _, _ = range_boundaries(filter_object.ref)
    criteria = []
    for filter_column in filter_object.filterColumn:
        header = sheet.cell(min_row, min_col + filter_column.colId).value
        label = str(header) if header is not None else f"Kolom {filter_column.colId + 1}"
        details = []
        filters = filter_column.filters
        if filters is not None:
            if filters.filter:
                details.append("waarde in " + ", ".join(map(str, filters.filter)))
            if filters.blank:
                details.append("lege waarden")
            for item in filters.dateGroupItem:
                date_parts = [
                    str(getattr(item, name))
                    for name in ("year", "month", "day", "hour", "minute", "second")
                    if getattr(item, name, None) is not None
                ]
                details.append("datum " + "-".join(date_parts))
        custom = filter_column.customFilters
        if custom is not None:
            operators = {
                "equal": "=",
                "notEqual": "≠",
                "greaterThan": ">",
                "lessThan": "<",
                "greaterThanOrEqual": "≥",
                "lessThanOrEqual": "≤",
            }
            details.extend(
                f"{operators.get(rule.operator, rule.operator or '=')} {rule.val}"
                for rule in custom.customFilter
            )
        if filter_column.dynamicFilter is not None:
            details.append(
                "dynamisch: " + str(filter_column.dynamicFilter.type)
            )
        if filter_column.top10 is not None:
            top = filter_column.top10
            details.append(
                f"{'bovenste' if top.top else 'onderste'} {top.val}"
                f"{'%' if top.percent else ' items'}"
            )
        if filter_column.colorFilter is not None:
            details.append("kleurfilter")
        if filter_column.iconFilter is not None:
            details.append("pictogramfilter")
        criteria.append((label, tuple(details)))
    return tuple(sorted(criteria))


def _filter_signature(sheet: Any) -> Any:
    return tuple(
        (name, filter_object.ref, _filter_criteria(sheet, filter_object))
        for name, filter_object in _filter_objects(sheet)
    )


def _filter_description(sheet: Any) -> str:
    descriptions = []
    for name, filter_object in _filter_objects(sheet):
        criteria = _filter_criteria(sheet, filter_object)
        labels = [
            f"{column}: {', '.join(details) if details else 'geen criterium'}"
            for column, details in criteria
        ]
        prefix = f"{name}, " if name != "werkblad" else ""
        descriptions.append(
            f"{prefix}{filter_object.ref or 'bereik onbekend'}"
            + (f" — {'; '.join(labels)}" if labels else "")
        )
    return "; ".join(descriptions) if descriptions else "Geen filter ingesteld"


def _sort_conditions(sheet: Any) -> list[tuple[int, bool]]:
    conditions = []
    for _, filter_object in _filter_objects(sheet):
        sort_state = filter_object.sortState
        if sort_state is None:
            continue
        for condition in sort_state.sortCondition:
            if not condition.ref:
                continue
            min_col, _, max_col, _ = range_boundaries(condition.ref)
            if min_col != max_col:
                continue
            conditions.append((min_col, bool(condition.descending)))
    return conditions


def _sort_description(
    sheet: Any,
    conditions: list[tuple[int, bool]],
    header_row: int = 1,
) -> str:
    columns = []
    for column, descending in conditions:
        header = sheet.cell(header_row, column).value
        columns.append(f"{header or get_column_letter(column)} ({'aflopend' if descending else 'oplopend'})")
    return ", ".join(columns)


def _sort_value(value: Any) -> tuple[int, Any]:
    if value is None:
        return (0, "")
    if isinstance(value, (int, float)):
        return (1, float(value))
    if isinstance(value, str):
        return (2, value.casefold())
    if hasattr(value, "isoformat"):
        return (3, value.isoformat())
    return (4, str(value))


def _is_sorted_by(
    sheet: Any,
    rows: list[int],
    conditions: list[tuple[int, bool]],
) -> bool:
    if len(rows) < 2:
        return True
    previous = None
    for row in rows:
        current = tuple(
            _sort_value(sheet.cell(row, column).value)
            for column, _ in conditions
        )
        if previous is not None:
            for index, (_, descending) in enumerate(conditions):
                if current[index] == previous[index]:
                    continue
                if descending and current[index] > previous[index]:
                    return False
                if not descending and current[index] < previous[index]:
                    return False
                break
        previous = current
    return True


def _rows_matching_filter(
    sheet: Any,
    rows: range,
    filter_object: Any,
) -> list[int]:
    if not filter_object.ref:
        return list(rows)
    min_col, header_row, _, _ = range_boundaries(filter_object.ref)
    criteria = []
    for filter_column in filter_object.filterColumn:
        filters = filter_column.filters
        if (
            filters is None
            or filters.dateGroupItem
            or filter_column.customFilters is not None
            or filter_column.dynamicFilter is not None
            or filter_column.top10 is not None
            or filter_column.colorFilter is not None
            or filter_column.iconFilter is not None
        ):
            return list(rows)
        accepted = {str(value).casefold() for value in filters.filter}
        include_blank = bool(filters.blank)
        if not accepted and not include_blank:
            continue
        criteria.append((min_col + filter_column.colId, accepted, include_blank))

    return [
        row for row in rows
        if all(
            (sheet.cell(row, column).value is None and include_blank)
            or str(sheet.cell(row, column).value).casefold() in accepted
            for column, accepted, include_blank in criteria
        )
    ]


def _filter_data_bounds(sheet: Any) -> tuple[int, int, int, int] | None:
    for _, filter_object in _filter_objects(sheet):
        if filter_object.ref:
            min_col, min_row, max_col, max_row = range_boundaries(filter_object.ref)
            return min_col, min_row, max_col, max_row
    return None


def _same_rows_in_range(
    expected: Any,
    actual: Any,
    bounds: tuple[int, int, int, int],
) -> bool:
    min_col, min_row, max_col, max_row = bounds
    if (
        expected.max_row < max_row
        or actual.max_row < max_row
        or expected.max_column < max_col
        or actual.max_column < max_col
    ):
        return False
    expected_rows = Counter(
        tuple(expected.cell(row, column).value for column in range(min_col, max_col + 1))
        for row in range(min_row, max_row + 1)
    )
    actual_rows = Counter(
        tuple(actual.cell(row, column).value for column in range(min_col, max_col + 1))
        for row in range(min_row, max_row + 1)
    )
    return expected_rows == actual_rows


def _defined_name_signatures(workbook: Any) -> dict[str, Any]:
    names = workbook.defined_names
    items = names.items() if hasattr(names, "items") else (
        (item.name, item) for item in names.definedName
    )
    return {
        name: (item.attr_text, item.localSheetId, item.hidden, item.function)
        for name, item in items
    }


def _add_difference(
    differences: list[Difference],
    sheet: str,
    item: str,
    expected: Any,
    actual: Any,
) -> None:
    differences.append(Difference(sheet, item, repr(expected), repr(actual)))


def compare_workbooks(reference_path: Path, student_path: Path) -> ComparisonResult:
    differences: list[Difference] = []
    try:
        reference = load_workbook(reference_path, data_only=False)
        student = load_workbook(student_path, data_only=False)
    except Exception as exc:
        return ComparisonResult(student_path, differences, f"Bestand kon niet worden gelezen: {exc}")

    reference_names = reference.sheetnames
    student_names = student.sheetnames
    for name in reference_names:
        if name not in student_names:
            _add_difference(differences, name, "Werkblad", "aanwezig", "ontbreekt")
    for name in student_names:
        if name not in reference_names:
            _add_difference(differences, name, "Werkblad", "ontbreekt", "extra werkblad")
    if reference_names != student_names:
        _add_difference(differences, "Werkmap", "Volgorde werkbladen", reference_names, student_names)

    for name in (sheet for sheet in reference_names if sheet in student_names):
        expected = reference[name]
        actual = student[name]

        for property_name, expected_value, actual_value in (
            ("Zichtbaarheid", expected.sheet_state, actual.sheet_state),
            ("Samengevoegde cellen", sorted(str(x) for x in expected.merged_cells.ranges),
             sorted(str(x) for x in actual.merged_cells.ranges)),
            ("Vastgezette deelvensters", expected.freeze_panes, actual.freeze_panes),
            ("Tabkleur", _color(expected.sheet_properties.tabColor),
             _color(actual.sheet_properties.tabColor)),
            ("Tabellen", _table_signatures(expected), _table_signatures(actual)),
            ("Afdrukbereik", str(expected.print_area), str(actual.print_area)),
        ):
            if expected_value != actual_value:
                _add_difference(differences, name, property_name, expected_value, actual_value)

        if _filter_signature(expected) != _filter_signature(actual):
            differences.append(Difference(
                name,
                "Filter",
                _filter_description(expected),
                _filter_description(actual),
            ))

        filter_bounds = _filter_data_bounds(expected)
        ignored_filter_rows: set[int] = set()
        if filter_bounds is not None:
            _, header_row, _, last_row = filter_bounds
            ignored_filter_rows.update(range(header_row + 1, last_row + 1))

        expected_rows = _row_dimension_signatures(expected, ignored_filter_rows)
        actual_rows = _row_dimension_signatures(actual, ignored_filter_rows)
        if expected_rows != actual_rows:
            expected_row_changes, actual_row_changes = _format_row_differences(
                expected_rows, actual_rows
            )
            _add_difference(differences, name, "Rijhoogtes/-instellingen",
                            expected_row_changes, actual_row_changes)

        expected_columns = _column_dimension_signatures(expected)
        actual_columns = _column_dimension_signatures(actual)
        if expected_columns != actual_columns:
            expected_column_changes, actual_column_changes = _format_column_differences(
                expected_columns, actual_columns
            )
            _add_difference(differences, name, "Kolombreedtes/-instellingen",
                            expected_column_changes, actual_column_changes)

        cell_keys = sorted(expected._cells.keys() | actual._cells.keys())
        formula_changes: dict[
            tuple[int, int], tuple[str, str, tuple[str, str]]
        ] = {}
        cell_value_changes: dict[tuple[int, int], tuple[Any, Any]] = {}
        style_changes: dict[tuple[int, int], dict[str, tuple[Any, Any]]] = {}
        sort_conditions = _sort_conditions(expected)
        reordered_data_bounds = None
        if filter_bounds is not None and sort_conditions:
            _, header_row, _, last_row = filter_bounds
            data_rows = range(header_row + 1, last_row + 1)
            reference_filter = next((
                filter_object
                for _, filter_object in _filter_objects(expected)
                if filter_object.ref
                and range_boundaries(filter_object.ref) == filter_bounds
            ), None)
            rows_to_sort_check = (
                _rows_matching_filter(actual, data_rows, reference_filter)
                if reference_filter is not None
                else list(data_rows)
            )
            if not _is_sorted_by(actual, rows_to_sort_check, sort_conditions):
                _add_difference(
                    differences,
                    name,
                    "Sortering",
                    _sort_description(expected, sort_conditions, header_row),
                    "De gegevens staan niet in de verwachte sorteervolgorde",
                )
            if _same_rows_in_range(expected, actual, filter_bounds):
                reordered_data_bounds = filter_bounds
        for row, column in cell_keys:
            left = expected._cells.get((row, column)) or expected.cell(row, column)
            right = actual._cells.get((row, column)) or actual.cell(row, column)
            coordinate = left.coordinate
            in_reordered_data = (
                reordered_data_bounds is not None
                and reordered_data_bounds[1] < row <= reordered_data_bounds[3]
                and reordered_data_bounds[0] <= column <= reordered_data_bounds[2]
            )
            if _cell_value(left) != _cell_value(right) and not in_reordered_data:
                if (
                    left.data_type == "f"
                    and right.data_type == "f"
                    and isinstance(left.value, str)
                    and isinstance(right.value, str)
                ):
                    formula_changes[(row, column)] = (
                        left.value,
                        right.value,
                        _formula_difference_key(left.value, right.value, coordinate),
                    )
                else:
                    cell_value_changes[(row, column)] = (
                        _cell_value(left), _cell_value(right)
                    )
            if _cell_annotations(left) != _cell_annotations(right):
                _add_difference(differences, name, f"{coordinate} hyperlinks/opmerking",
                                _cell_annotations(left), _cell_annotations(right))
            left_style = _style(left)
            right_style = _style(right)
            for category, expected_style in left_style.items():
                if expected_style != right_style[category]:
                    style_changes.setdefault((row, column), {})[category] = (
                        expected_style, right_style[category]
                    )
        differences.extend(_coalesce_formula_differences(name, formula_changes))
        differences.extend(_coalesce_cell_differences(name, cell_value_changes))
        differences.extend(_coalesce_style_differences(name, style_changes))

        expected_charts = expected._charts
        actual_charts = actual._charts
        expected_chart_summary = []
        actual_chart_summary = []
        for index in range(max(len(expected_charts), len(actual_charts))):
            if index >= len(expected_charts):
                expected_chart_summary.append("Ontbreekt")
                actual_chart_summary.append(
                    f"Grafiek {index + 1}: {_chart_description(actual_charts[index])}"
                )
            elif index >= len(actual_charts):
                expected_chart_summary.append(
                    f"Grafiek {index + 1}: {_chart_description(expected_charts[index])}"
                )
                actual_chart_summary.append("Ontbreekt")
            elif _chart_signature(expected_charts[index]) != _chart_signature(actual_charts[index]):
                expected_chart_summary.append(
                    f"Grafiek {index + 1}: {_chart_description(expected_charts[index])}"
                )
                actual_chart_summary.append(
                    f"Grafiek {index + 1}: {_chart_description(actual_charts[index])}"
                )
        if expected_chart_summary:
            differences.append(Difference(
                name,
                "Grafieken",
                "\n".join(expected_chart_summary),
                "\n".join(actual_chart_summary),
            ))

        expected_validations = sorted(_object_signature(item) for item in expected.data_validations.dataValidation)
        actual_validations = sorted(_object_signature(item) for item in actual.data_validations.dataValidation)
        if expected_validations != actual_validations:
            _add_difference(differences, name, "Gegevensvalidatie", expected_validations, actual_validations)

        expected_cf = _conditional_formatting_rules(expected)
        actual_cf = _conditional_formatting_rules(actual)
        expected_descriptions = Counter()
        actual_descriptions = Counter()
        for signature, description in expected_cf:
            expected_descriptions[signature, description] += 1
        for signature, description in actual_cf:
            actual_descriptions[signature, description] += 1
        missing = expected_descriptions - actual_descriptions
        extra = actual_descriptions - expected_descriptions
        if missing or extra:
            expected_summary = [
                f"Ontbreekt bij leerling: {description}"
                for (_, description), count in missing.items()
                for _ in range(count)
            ]
            actual_summary = [
                f"Extra bij leerling: {description}"
                for (_, description), count in extra.items()
                for _ in range(count)
            ]
            differences.append(Difference(
                name,
                "Voorwaardelijke opmaak",
                "\n".join(expected_summary) or "Geen inhoudelijke verschillen",
                "\n".join(actual_summary) or "Geen inhoudelijke verschillen",
            ))

    reference_names_defined = _defined_name_signatures(reference)
    student_names_defined = _defined_name_signatures(student)
    if reference_names_defined != student_names_defined:
        _add_difference(differences, "Werkmap", "Benoemde bereiken",
                        reference_names_defined, student_names_defined)

    reference.close()
    student.close()
    return ComparisonResult(student_path, differences)


def _safe_filename(name: str) -> str:
    return re.sub(r"[^\w.-]+", "_", name, flags=re.UNICODE).strip("._") or "leerling"


def write_report(result: ComparisonResult, output_path: Path) -> None:
    title = html.escape(result.student_file.stem)
    if result.error:
        status = "Bestand niet beoordeeld"
        detail = f'<p class="error">{html.escape(result.error)}</p>'
    elif result.matches:
        status = "Geen verschillen gevonden"
        detail = '<p class="success">De gecontroleerde onderdelen komen overeen met het modelbestand.</p>'
    else:
        status = f"{len(result.differences)} verschil(len) gevonden"
        rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.sheet)}</td>"
            f"<td>{html.escape(item.item)}</td>"
            f"<td><code>{html.escape(item.expected)}</code></td>"
            f"<td><code>{html.escape(item.actual)}</code></td>"
            "</tr>"
            for item in result.differences
        )
        detail = (
            "<table><thead><tr><th>Werkblad</th><th>Onderdeel</th>"
            "<th>Modeloplossing</th><th>Leerlingbestand</th></tr></thead>"
            f"<tbody>{rows}</tbody></table>"
        )
    document = f"""<!doctype html>
<html lang="nl"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Excel Checker - {title}</title>
<style>
body{{font:16px/1.5 system-ui,sans-serif;max-width:1100px;margin:2rem auto;padding:0 1rem;color:#202124}}
h1{{font-size:1.6rem}}.success{{color:#137333}}.error{{color:#b3261e}}
table{{border-collapse:collapse;width:100%;margin-top:1rem}}th,td{{border:1px solid #dadce0;padding:.6rem;text-align:left;vertical-align:top}}
th{{background:#f1f3f4}}code{{white-space:pre-wrap;overflow-wrap:anywhere}}
</style>
<h1>Controlerapport: {title}</h1><p><strong>Resultaat:</strong> {status}</p>
{detail}
</html>"""
    output_path.write_text(document, encoding="utf-8")


def check_folder(reference_path: Path, student_folder: Path, output_folder: Path) -> list[Path]:
    output_folder.mkdir(parents=True, exist_ok=True)
    student_files = sorted(
        path for path in student_folder.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and not path.name.startswith("~$")
        and path.resolve() != reference_path.resolve()
    )
    report_paths: list[Path] = []
    used_names: dict[str, int] = {}
    for student_file in student_files:
        base = _safe_filename(student_file.stem)
        used_names[base] = used_names.get(base, 0) + 1
        suffix = f"_{used_names[base]}" if used_names[base] > 1 else ""
        report_path = output_folder / f"{base}{suffix}_rapport.html"
        write_report(compare_workbooks(reference_path, student_file), report_path)
        report_paths.append(report_path)
    return report_paths
