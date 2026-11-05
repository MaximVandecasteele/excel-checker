from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import (
    get_column_letter,
    column_index_from_string,
    range_boundaries,
)
from openpyxl.formula.translate import Translator
from openpyxl.styles.numbers import BUILTIN_FORMATS


@dataclass(frozen=True)
class Difference:
    sheet: str
    item: str
    expected: str
    actual: str
    category: str | None = None


@dataclass
class ComparisonResult:
    student_file: Path
    differences: list[Difference]
    error: str | None = None

    @property
    def matches(self) -> bool:
        return not self.differences and self.error is None


@lru_cache(maxsize=32)
def _theme_color_values(theme: bytes | None) -> tuple[str | None, ...]:
    if not theme:
        return ()
    root = ET.fromstring(theme)
    scheme = root.find(".//{*}clrScheme")
    if scheme is None:
        return ()

    color_names = (
        "lt1", "dk1", "lt2", "dk2",
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hlink", "folHlink",
    )
    entries = {
        entry.tag.rsplit("}", 1)[-1]: entry
        for entry in scheme
    }
    colors = []
    for name in color_names:
        entry = entries.get(name)
        if entry is None:
            colors.append(None)
            continue
        color = next(iter(entry), None)
        if color is None:
            colors.append(None)
        elif color.tag.rsplit("}", 1)[-1] == "srgbClr":
            colors.append(color.get("val"))
        else:
            colors.append(color.get("lastClr"))
    return tuple(colors)


def _tinted_rgb(rgb: str, tint: float) -> str:
    rgb = rgb[-6:]
    if tint == 0:
        return rgb.upper()

    hls_max = 240

    def rounded_division(numerator: int, denominator: int) -> int:
        if numerator < 0:
            return -((-numerator + denominator // 2) // denominator)
        return (numerator + denominator // 2) // denominator

    red, green, blue = (
        int(rgb[index:index + 2], 16)
        for index in (0, 2, 4)
    )
    maximum = max(red, green, blue)
    minimum = min(red, green, blue)
    delta = maximum - minimum
    total = maximum + minimum
    lightness = (total * hls_max + 255) // (2 * 255)
    if delta == 0:
        hue, saturation = 160, 0
    else:
        saturation_denominator = (
            total if lightness <= hls_max // 2 else 2 * 255 - total
        )
        saturation = (
            delta * hls_max + saturation_denominator // 2
        ) // saturation_denominator

        def hue_distance(channel: int) -> int:
            return (
                (maximum - channel) * (hls_max // 6) + delta // 2
            ) // delta

        if red == maximum:
            hue = hue_distance(blue) - hue_distance(green)
        elif green == maximum:
            hue = hls_max // 3 + hue_distance(red) - hue_distance(blue)
        else:
            hue = hls_max * 2 // 3 + hue_distance(green) - hue_distance(red)
        hue %= hls_max

    tint_millis = int(tint * 1000 + 0.5) if tint >= 0 else -int(-tint * 1000 + 0.5)
    tint_millis = max(-1000, min(1000, tint_millis))
    if tint_millis > 0:
        lightness += ((hls_max - lightness) * tint_millis) // 1000
    elif tint_millis < 0:
        lightness = (lightness * (1000 + tint_millis)) // 1000

    if saturation == 0:
        channels = (rounded_division(lightness * 255, hls_max),) * 3
    else:
        if lightness <= hls_max / 2:
            high = rounded_division(
                lightness * (hls_max + saturation), hls_max
            )
        else:
            high = lightness + saturation - rounded_division(
                lightness * saturation, hls_max
            )
        low = 2 * lightness - high

        def hue_channel(position: int) -> int:
            position %= hls_max
            if position < hls_max // 6:
                return low + rounded_division(
                    (high - low) * position, hls_max // 6
                )
            if position < hls_max // 2:
                return high
            if position < hls_max * 2 // 3:
                return low + rounded_division(
                    (high - low) * (hls_max * 2 // 3 - position),
                    hls_max // 6,
                )
            return low

        channels = (
            hue_channel(hue + hls_max // 3),
            hue_channel(hue),
            hue_channel(hue - hls_max // 3),
        )
        channels = tuple(
            rounded_division(channel * 255, hls_max)
            for channel in channels
        )
    return "".join(f"{channel:02X}" for channel in channels)


def _color(color: Any, theme: bytes | None = None) -> Any:
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
    signature = (color.type, value, color.tint)
    if color.type == "theme":
        theme_colors = _theme_color_values(theme)
        theme_index = color.theme
        if theme_index is not None and theme_index < len(theme_colors):
            rgb = theme_colors[theme_index]
            if rgb:
                return signature + (_tinted_rgb(rgb, color.tint),)
    return signature


def _side(side: Any, theme: bytes | None = None) -> Any:
    if side is None:
        return None
    color = _color(side.color, theme)
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


def _number_format_signature(cell: Any) -> tuple[str, str | None]:
    number_format = _normalize_number_format(cell.number_format)
    builtin_id = cell._style.numFmtId if cell.has_style else 0
    builtin_names = {
        0: "standaard",
        1: "getal",
        2: "getal",
        3: "getal",
        4: "getal",
        5: "valuta",
        6: "valuta",
        7: "valuta",
        8: "valuta",
        9: "percentage",
        10: "percentage",
        11: "wetenschappelijk",
        12: "breuk",
        13: "breuk",
        14: "datum",
        15: "datum",
        16: "datum",
        17: "datum",
        18: "tijd",
        19: "tijd",
        20: "tijd",
        21: "tijd",
        22: "datum en tijd",
        37: "getal",
        38: "getal",
        39: "getal",
        40: "getal",
        41: "financieel",
        42: "financieel",
        43: "financieel",
        44: "financieel",
        45: "tijd",
        46: "tijd",
        47: "tijd",
        48: "wetenschappelijk",
        49: "tekst",
    }
    if BUILTIN_FORMATS.get(builtin_id) == cell.number_format:
        return number_format, builtin_names.get(builtin_id)
    localized_defaults = {
        '_ "€" * #,##0.00_ ;_ "€" * -#,##0.00_ ;_ "€" * "-"??_ ;_ @_': "financieel",
        '"€" #,##0.00': "valuta",
    }
    localized_name = localized_defaults.get(number_format.strip())
    if localized_name:
        return number_format, localized_name
    return number_format, None


def _alignment_horizontal(cell: Any) -> str | None:
    horizontal = cell.alignment.horizontal
    if horizontal in (None, "general"):
        return "general"

    value = cell.value
    if (
        horizontal == "left"
        and isinstance(value, str)
        and cell.data_type != "f"
    ):
        return "general"
    if (
        horizontal == "right"
        and (
            cell.is_date
            or isinstance(value, (int, float))
            and not isinstance(value, bool)
        )
    ):
        return "general"
    return horizontal


def _style(cell: Any) -> dict[str, Any]:
    theme = cell.parent.parent.loaded_theme
    font = cell.font
    fill = cell.fill
    alignment = cell.alignment
    border = cell.border
    protection = cell.protection
    return {
        "lettertype": (
            font.name, font.sz, font.b, font.i, font.u, font.strike,
            _color(font.color, theme), font.vertAlign, font.outline, font.shadow,
            font.charset, font.family, font.scheme,
        ),
        "opvulling": (
            fill.fill_type, _color(fill.fgColor, theme), _color(fill.bgColor, theme),
            fill.patternType,
        ),
        "rand": (
            _side(border.left, theme), _side(border.right, theme),
            _side(border.top, theme), _side(border.bottom, theme),
            _side(border.diagonal, theme),
            border.diagonalUp, border.diagonalDown, border.outline,
        ),
        "uitlijning": (
            _alignment_horizontal(cell), alignment.vertical, alignment.textRotation,
            alignment.wrapText, alignment.shrinkToFit, alignment.indent,
            alignment.relativeIndent, alignment.justifyLastLine,
            alignment.readingOrder,
        ),
        "getalnotatie": _number_format_signature(cell),
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
                "Functies / berekeningen",
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
            "Functies / berekeningen",
        ))
    return grouped


def _short_value(value: Any, limit: int = 70) -> str:
    text = repr(value)
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _format_style_color(color: Any) -> str:
    if color is None:
        return "geen kleur"
    color_type, value, tint = color[:3]
    if color_type == "theme" and len(color) > 3:
        description = f"#{color[3]}"
    elif color_type == "theme":
        description = f"themakleur {value}"
    elif color_type == "rgb" and isinstance(value, str):
        description = f"#{value[-6:]}"
    elif color_type == "indexed":
        description = f"geïndexeerde kleur {value}"
    else:
        description = str(value or color_type)
    if tint and not (color_type == "theme" and len(color) > 3):
        description += f" (tint {tint:.2f})"
    return description


def _format_font_summary(value: tuple[Any, ...], other: tuple[Any, ...]) -> str:
    parts = []
    if value[0] != other[0]:
        parts.append(value[0] or "standaardlettertype")
    if value[1] != other[1]:
        size = value[1]
        formatted_size = f"{size:g}" if isinstance(size, (int, float)) else size
        parts.append(
            f"{formatted_size} pt" if size is not None else "standaardgrootte"
        )

    for index, positive, negative in (
        (2, "vet", "normaal"),
        (3, "cursief", "recht"),
        (5, "doorgestreept", "niet doorgestreept"),
        (8, "contour", "geen contour"),
        (9, "schaduw", "geen schaduw"),
    ):
        if value[index] != other[index]:
            parts.append(positive if value[index] else negative)

    if value[4] != other[4]:
        underline_names = {
            None: "niet onderstreept",
            "single": "enkel onderstreept",
            "double": "dubbel onderstreept",
            "singleAccounting": "enkel onderstreept (boekhouding)",
            "doubleAccounting": "dubbel onderstreept (boekhouding)",
        }
        parts.append(underline_names.get(value[4], f"onderstreept ({value[4]})"))

    if value[6] != other[6]:
        parts.append(
            "standaardkleur"
            if value[6] is None
            else _format_style_color(value[6])
        )
    if value[7] != other[7]:
        vertical_names = {
            None: "normale positie",
            "baseline": "normale positie",
            "superscript": "superscript",
            "subscript": "subscript",
        }
        parts.append(vertical_names.get(value[7], str(value[7])))

    if value[10] != other[10]:
        parts.append(f"tekenverzameling {value[10] or 'standaard'}")
    if value[11] != other[11]:
        family_names = {
            0: "standaard",
            1: "modern",
            2: "roman",
            3: "swiss",
            4: "script",
            5: "decoratief",
        }
        parts.append(f"lettertypefamilie {family_names.get(value[11], value[11])}")
    if value[12] != other[12]:
        scheme_names = {
            None: "geen themaschema",
            "major": "themakop",
            "minor": "themalichaam",
        }
        parts.append(scheme_names.get(value[12], f"themaschema {value[12]}"))
    return ", ".join(parts) or "lettertype-instelling gewijzigd"


def _format_border_summary(value: tuple[Any, ...], other: tuple[Any, ...]) -> str:
    side_names = ("linkerrand", "rechterrand", "bovenrand", "onderrand", "diagonale rand")
    style_names = {
        "hair": "haarlijn",
        "dotted": "gestippeld",
        "dashDot": "streep-punt",
        "dashDotDot": "streep-punt-punt",
        "dashed": "gestreept",
        "slantDashDot": "schuin streep-punt",
        "mediumDashDot": "middelmatig streep-punt",
        "mediumDashDotDot": "middelmatig streep-punt-punt",
        "mediumDashed": "middelmatig gestreept",
        "thin": "dun",
        "medium": "middelmatig",
        "thick": "dik",
        "double": "dubbel",
    }
    parts = []
    for index, name in enumerate(side_names):
        current_side = value[index]
        other_side = other[index]
        if current_side == other_side:
            continue
        if current_side is None or other_side is None:
            if current_side is None:
                description = "geen"
                parts.append(f"{name}: {description}")
            else:
                style, color = current_side
                if style:
                    parts.append(
                        f"{name} stijl: {style_names.get(style, style)}"
                    )
                if color:
                    parts.append(
                        f"{name} kleur: {_format_style_color(color)}"
                    )
                if not style and not color:
                    parts.append(f"{name}: geen")
            continue

        style, color = current_side
        other_style, other_color = other_side
        if style != other_style:
            parts.append(f"{name} stijl: {style_names.get(style, style) if style else 'geen'}")
        if color != other_color:
            parts.append(
                f"{name} kleur: "
                f"{_format_style_color(color) if color else 'geen kleur'}"
            )

    flag_names = (
        (5, "diagonale rand omhoog"),
        (6, "diagonale rand omlaag"),
        (7, "omlijning"),
    )
    for index, name in flag_names:
        if value[index] != other[index]:
            parts.append(name if value[index] else f"geen {name}")
    return "\n".join(parts) or "randinstelling gewijzigd"


def _format_alignment_summary(
    value: tuple[Any, ...],
    other: tuple[Any, ...],
) -> str:
    parts = []
    horizontal_names = {
        "general": "standaard",
        "left": "links",
        "center": "centreren",
        "right": "rechts",
        "fill": "opvullen",
        "justify": "uitvullen",
        "centerContinuous": "doorlopend centreren",
        "distributed": "gespreid",
    }
    vertical_names = {
        "bottom": "onderaan",
        "center": "midden",
        "top": "bovenaan",
        "justify": "uitvullen",
        "distributed": "gespreid",
    }
    if value[0] != other[0]:
        parts.append(
            f"horizontaal: "
            f"{horizontal_names.get(value[0], value[0]) if value[0] else 'standaard'}"
        )
    if value[1] != other[1]:
        parts.append(
            f"verticaal: {vertical_names.get(value[1], value[1]) if value[1] else 'standaard'}"
        )
    if value[2] != other[2]:
        parts.append(f"draaiing: {value[2]}°")
    if value[3] != other[3]:
        parts.append("tekstterugloop" if value[3] else "geen tekstterugloop")
    if value[4] != other[4]:
        parts.append("passend maken" if value[4] else "niet passend gemaakt")
    if value[5] != other[5]:
        parts.append(f"inspringing: {value[5]:g}" if value[5] else "geen inspringing")
    if value[6] != other[6]:
        parts.append(
            f"relatieve inspringing: {value[6]:g}"
            if value[6]
            else "geen relatieve inspringing"
        )
    if value[7] != other[7]:
        parts.append("laatste regel uitgevuld" if value[7] else "laatste regel niet uitgevuld")
    if value[8] != other[8]:
        reading_order_names = {
            0: "contextafhankelijk",
            1: "links naar rechts",
            2: "rechts naar links",
        }
        parts.append(
            f"leesrichting: {reading_order_names.get(value[8], value[8])}"
        )
    return "\n".join(parts) or "uitlijningsinstelling gewijzigd"


def _format_style_summary(
    changes: dict[str, tuple[Any, Any]],
    side: int,
    counterparts: dict[str, tuple[Any, Any]] | None = None,
) -> str:
    summaries = []
    for category, values in changes.items():
        value = values[side]
        if category == "opvulling":
            fill_type, foreground, background, _ = value
            if fill_type is None:
                description = "geen opvulling"
            else:
                description = (
                    "effen opvulling"
                    if fill_type == "solid"
                    else f"opvulling {fill_type}"
                )
                color = foreground or background
                if color:
                    description += f", {_format_style_color(color)}"
            summaries.append(description)
        elif category == "getalnotatie":
            number_format, builtin_name = value
            counterpart = (counterparts or {}).get(category)
            other_name = counterpart[1 - side][1] if counterpart else None
            other_format = counterpart[1 - side][0] if counterpart else None
            if builtin_name and builtin_name == other_name and number_format != other_format:
                description = f"{builtin_name} : {number_format}"
            elif builtin_name:
                description = builtin_name
            else:
                description = _short_value(number_format, 120)
            summaries.append(f"getalnotatie: {description}")
        elif category == "lettertype":
            counterpart = (counterparts or {}).get(category)
            other_value = counterpart[1 - side] if counterpart else value
            summaries.append(f"lettertype: {_format_font_summary(value, other_value)}")
        elif category == "rand":
            counterpart = (counterparts or {}).get(category)
            other_value = counterpart[1 - side] if counterpart else value
            summaries.append(_format_border_summary(value, other_value))
        elif category == "uitlijning":
            counterpart = (counterparts or {}).get(category)
            other_value = counterpart[1 - side] if counterpart else value
            summaries.append(_format_alignment_summary(value, other_value))
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
                    "Invoer data (inhoud)",
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
    border_changes: dict[tuple[int, int], tuple[Any, Any]] = {}
    alignment_changes: dict[tuple[int, int], tuple[Any, Any]] = {}
    for coordinate, changes in style_changes.items():
        for category, (expected, actual) in changes.items():
            if category == "rand":
                border_changes[coordinate] = expected, actual
                continue
            if category == "uitlijning":
                alignment_changes[coordinate] = expected, actual
                continue
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
            _format_style_summary(
                {category: (expected, actual)},
                0,
                {category: (expected, actual)},
            ),
            _format_style_summary(
                {category: (expected, actual)},
                1,
                {category: (expected, actual)},
            ),
            "Opmaak",
        ))

    remaining = set(alignment_changes)
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

        min_row = min(row for row, _ in component)
        max_row = max(row for row, _ in component)
        min_column = min(column for _, column in component)
        max_column = max(column for _, column in component)
        first = f"{get_column_letter(min_column)}{min_row}"
        last = f"{get_column_letter(max_column)}{max_row}"
        range_label = first if first == last else f"{first}:{last}"

        def alignment_summary(side: int) -> str:
            property_values: dict[str, list[str]] = {}
            standalone_summaries = []
            seen_standalone = set()
            for coordinate in sorted(component):
                expected, actual = alignment_changes[coordinate]
                value, other = (expected, actual) if side == 0 else (actual, expected)
                for summary in _format_alignment_summary(value, other).splitlines():
                    if ": " in summary:
                        property_name, property_value = summary.split(": ", 1)
                        values = property_values.setdefault(property_name, [])
                        if property_value not in values:
                            values.append(property_value)
                    elif summary not in seen_standalone:
                        standalone_summaries.append(summary)
                        seen_standalone.add(summary)
            summaries = [
                f"{property_name}: {', '.join(values)}"
                for property_name, values in property_values.items()
            ]
            return "\n".join(summaries + standalone_summaries)

        differences.append(Difference(
            sheet,
            f"Opmaak: uitlijning in bereik {range_label} "
            f"({len(component)} {'cel' if len(component) == 1 else 'cellen'})",
            alignment_summary(0),
            alignment_summary(1),
            "Opmaak",
        ))

    remaining = set(border_changes)
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

        min_row = min(row for row, _ in component)
        max_row = max(row for row, _ in component)
        min_column = min(column for _, column in component)
        max_column = max(column for _, column in component)
        first = f"{get_column_letter(min_column)}{min_row}"
        last = f"{get_column_letter(max_column)}{max_row}"
        range_label = first if first == last else f"{first}:{last}"

        expected_summary = _format_border_range_summary(
            [border_changes[coordinate] for coordinate in component],
            0,
        )
        actual_summary = _format_border_range_summary(
            [border_changes[coordinate] for coordinate in component],
            1,
        )
        differences.append(Difference(
            sheet,
            f"Opmaak: rand in bereik {range_label}",
            expected_summary,
            actual_summary,
            "Opmaak",
        ))
    return differences


def _format_border_range_summary(
    changes: list[tuple[tuple[Any, ...], tuple[Any, ...]]],
    side: int,
) -> str:
    side_names = (
        "linkerrand",
        "rechterrand",
        "bovenrand",
        "onderrand",
        "diagonale rand",
    )
    style_names = {
        "hair": "haarlijn",
        "dotted": "gestippeld",
        "dashDot": "streep-punt",
        "dashDotDot": "streep-punt-punt",
        "dashed": "gestreept",
        "slantDashDot": "schuin streep-punt",
        "mediumDashDot": "middelmatig streep-punt",
        "mediumDashDotDot": "middelmatig streep-punt-punt",
        "mediumDashed": "middelmatig gestreept",
        "thin": "dun",
        "medium": "middelmatig",
        "thick": "dik",
        "double": "dubbel",
    }

    parts = []
    for index, name in enumerate(side_names):
        edge_changes = [
            pair for pair in changes if pair[0][index] != pair[1][index]
        ]
        if not edge_changes:
            continue
        edge_values = [
            pair[side][index] for pair in edge_changes
        ]
        if all(value is None for value in edge_values):
            parts.append(f"{name}: geen")
            continue
        for value_index, label, formatter in (
            (0, "stijl", lambda value: style_names.get(value, value) if value else "geen"),
            (1, "kleur", lambda value: _format_style_color(value) if value else "geen kleur"),
        ):
            values = {
                pair[side][index][value_index]
                if pair[side][index] is not None
                else None
                for pair in edge_changes
                if (
                    pair[0][index][value_index]
                    if pair[0][index] is not None
                    else None
                ) != (
                    pair[1][index][value_index]
                    if pair[1][index] is not None
                    else None
                )
            }
            if values:
                parts.append(
                    f"{name} {label}: "
                    + ", ".join(formatter(value) for value in sorted(values, key=str))
                )

    for index, name in (
        (5, "diagonale rand omhoog"),
        (6, "diagonale rand omlaag"),
        (7, "omlijning"),
    ):
        values = {
            pair[side][index]
            for pair in changes
            if pair[0][index] != pair[1][index]
        }
        if values:
            labels = ["aan" if value else "uit" for value in sorted(values)]
            parts.append(f"{name}: {', '.join(labels)}")
    return "\n".join(parts) or "geen"


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


def _dimension_differences(
    expected: dict[int, tuple[Any, ...]],
    actual: dict[int, tuple[Any, ...]],
    expected_default_size: float,
    actual_default_size: float,
    tolerance: float = 0.2,
) -> tuple[dict[int, tuple[Any, ...]], dict[int, tuple[Any, ...]]]:
    indices = expected.keys() | actual.keys()

    def effective_size(
        signature: tuple[Any, ...] | None,
        default_size: float,
    ) -> float:
        if signature is None or signature[0] is None:
            return default_size
        return signature[0]

    expected_sizes = {
        index: effective_size(expected.get(index), expected_default_size)
        for index in indices
    }
    actual_sizes = {
        index: effective_size(actual.get(index), actual_default_size)
        for index in indices
    }
    equal_model_dimensions: dict[float, list[int]] = {}
    for index, size in expected_sizes.items():
        equal_model_dimensions.setdefault(size, []).append(index)

    inconsistent_groups = set()
    for indices_in_group in equal_model_dimensions.values():
        if len(indices_in_group) > 1 and len({
            actual_sizes[index] for index in indices_in_group
        }) > 1:
            inconsistent_groups.update(indices_in_group)

    expected_changes = {}
    actual_changes = {}
    for index in indices:
        expected_size = expected_sizes[index]
        actual_size = actual_sizes[index]
        size_changed = (
            abs(actual_size - expected_size) > abs(expected_size) * tolerance
            if expected_size
            else actual_size != expected_size
        )
        relation_changed = index in inconsistent_groups
        size_changed = size_changed or relation_changed
        expected_signature = expected.get(index)
        actual_signature = actual.get(index)
        expected_settings = expected_signature[1:] if expected_signature else ()
        actual_settings = actual_signature[1:] if actual_signature else ()
        signature_length = max(
            len(expected_signature or ()),
            len(actual_signature or ()),
        )
        default_settings = (False, 0, False, False) if signature_length > 4 else (
            False, 0, False
        )
        if (
            not size_changed
            and (expected_settings or default_settings)
            == (actual_settings or default_settings)
        ):
            continue

        if not size_changed:
            expected_size_value = (
                expected_signature[0] if expected_signature else None
            )
            actual_size_value = expected_size_value
        else:
            expected_size_value = expected_signature[0] if expected_signature else None
            actual_size_value = actual_signature[0] if actual_signature else None

        expected_value = (expected_size_value,) + (
            expected_settings or default_settings
        )
        actual_value = (actual_size_value,) + (
            actual_settings or default_settings
        )
        if relation_changed:
            if len(expected_value) <= 4:
                expected_value += (False,) * (5 - len(expected_value))
                actual_value += (False,) * (5 - len(actual_value))
            expected_value += (False,)
            actual_value += (True,)
        expected_changes[index] = expected_value
        actual_changes[index] = actual_value
    return expected_changes, actual_changes


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
        if len(value) > 5 and value[5]:
            parts.append("wijkt af van gelijke modelafmetingen")
        return ", ".join(parts) or "standaard"

    expected_summary = []
    actual_summary = []
    for start, end, expected_value, actual_value in changes:
        extent = format_index(start, end)
        expected_summary.append(f"{extent}: {format_side(expected_value)}")
        actual_summary.append(f"{extent}: {format_side(actual_value)}")
    return "\n".join(expected_summary), "\n".join(actual_summary)


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
    category: str | None = None,
) -> None:
    differences.append(
        Difference(sheet, item, repr(expected), repr(actual), category)
    )


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
        expected_row_changes, actual_row_changes = _dimension_differences(
            expected_rows,
            actual_rows,
            expected.sheet_format.defaultRowHeight or 15.0,
            actual.sheet_format.defaultRowHeight or 15.0,
        )
        if expected_row_changes:
            expected_row_changes, actual_row_changes = _format_row_differences(
                expected_row_changes, actual_row_changes
            )
            differences.append(Difference(
                name,
                "Rijhoogtes/-instellingen",
                expected_row_changes,
                actual_row_changes,
            ))

        expected_columns = _column_dimension_signatures(expected)
        actual_columns = _column_dimension_signatures(actual)
        expected_default_width = expected.sheet_format.defaultColWidth or 8.7109375
        actual_default_width = actual.sheet_format.defaultColWidth or 8.7109375
        expected_column_changes, actual_column_changes = _dimension_differences(
            expected_columns,
            actual_columns,
            expected_default_width,
            actual_default_width,
        )
        if expected_column_changes:
            expected_column_changes, actual_column_changes = _format_column_differences(
                expected_column_changes, actual_column_changes
            )
            differences.append(Difference(
                name,
                "Kolombreedtes/-instellingen",
                expected_column_changes,
                actual_column_changes,
            ))

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
                if (
                    category == "uitlijning"
                    and left.value is None
                    and right.value is None
                ):
                    continue
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


def _report_value(item: str, value: str) -> str:
    if not item.startswith(("Opmaak: lettertype", "Opmaak: opvulling")):
        return html.escape(value)

    parts = []
    start = 0
    for match in re.finditer(r"#[0-9A-Fa-f]{6}\b", value):
        parts.append(html.escape(value[start:match.start()]))
        color = match.group().upper()
        parts.append(
            f'{color} <span class="color-swatch" '
            f'style="background-color:{color}" title="{color}"></span>'
        )
        start = match.end()
    parts.append(html.escape(value[start:]))
    return "".join(parts)


_REPORT_CATEGORIES = (
    "Opmaak",
    "Functies / berekeningen",
    "Invoer data (inhoud)",
    "Sorteren/filteren",
    "Afdrukeigenschappen",
    "Grafieken",
)


def _report_category(difference: Difference) -> str:
    if difference.category:
        return difference.category
    item = difference.item
    if item.startswith(("Opmaak:", "Rijhoogtes", "Kolombreedtes")) or item in {
        "Tabkleur",
        "Voorwaardelijke opmaak",
    }:
        return "Opmaak"
    if item in {"Sortering", "Filter"}:
        return "Sorteren/filteren"
    if item == "Afdrukbereik":
        return "Afdrukeigenschappen"
    if item == "Grafieken":
        return "Grafieken"
    if item == "Benoemde bereiken":
        return "Functies / berekeningen"
    return "Invoer data (inhoud)"


def _report_format_subcategory(difference: Difference) -> str:
    item = difference.item
    if item.startswith("Opmaak: "):
        label = item[len("Opmaak: "):]
        if label.startswith("rand "):
            return "Rand"
        return label.split(" ", 1)[0].split(":", 1)[0].capitalize()
    return item


def _report_item_label(difference: Difference) -> str:
    prefix = "Opmaak: "
    if _report_category(difference) == "Opmaak" and difference.item.startswith(prefix):
        return difference.item[len(prefix):]
    return difference.item


def write_report(result: ComparisonResult, output_path: Path) -> None:
    title = html.escape(result.student_file.stem)
    scored_categories: set[str] = set()

    def score_fields(category: str) -> str:
        if category in scored_categories:
            return ""
        scored_categories.add(category)
        escaped_category = html.escape(category, quote=True)
        return (
            f'<span class="inline-score" data-score-category="{escaped_category}">'
            f'<label>Behaald <input type="number" min="0" step="any" '
            f'aria-label="{escaped_category} behaalde punten" '
            'data-score-field="earned"></label>'
            '<span aria-hidden="true">/</span>'
            f'<label>Maximum <input type="number" min="0" step="any" '
            f'aria-label="{escaped_category} maximum punten" '
            'data-score-field="maximum"></label>'
            "</span>"
        )

    if result.error:
        status = "Bestand niet beoordeeld"
        detail = f'<p class="error">{html.escape(result.error)}</p>'
    elif result.matches:
        status = "Geen verschillen gevonden"
        detail = '<p class="success">De gecontroleerde onderdelen komen overeen met het modelbestand.</p>'
    else:
        status = f"{len(result.differences)} verschil(len) gevonden"
        sheets: dict[str, dict[str, list[Difference]]] = {}
        for difference in result.differences:
            category = _report_category(difference)
            sheets.setdefault(difference.sheet, {}).setdefault(category, []).append(
                difference
            )

        sheet_sections = []
        for sheet, categories in sheets.items():
            category_sections = []
            for category in _REPORT_CATEGORIES:
                differences = categories.get(category)
                if not differences:
                    continue

                def render_table(items: list[Difference]) -> str:
                    rows = "".join(
                        "<tr>"
                        f"<td>{html.escape(_report_item_label(item))}</td>"
                        f"<td><code>{_report_value(item.item, item.expected)}</code></td>"
                        f"<td><code>{_report_value(item.item, item.actual)}</code></td>"
                        "</tr>"
                        for item in items
                    )
                    return (
                        "<table><thead><tr><th>Onderdeel</th>"
                        "<th>Modeloplossing</th><th>Leerlingbestand</th></tr></thead>"
                        f"<tbody>{rows}</tbody></table>"
                    )

                if category == "Opmaak":
                    subcategories: dict[str, list[Difference]] = {}
                    for difference in differences:
                        subcategory = _report_format_subcategory(difference)
                        subcategories.setdefault(subcategory, []).append(difference)
                    format_sections = "".join(
                        f"<h4>{html.escape(subcategory)}</h4>"
                        + render_table(subcategory_differences)
                        for subcategory, subcategory_differences in subcategories.items()
                    )
                    category_sections.append(
                        f'<div class="category-heading"><h3>{html.escape(category)}</h3>'
                        f"{score_fields(category)}</div>{format_sections}"
                    )
                    continue

                category_sections.append(
                    f'<div class="category-heading"><h3>{html.escape(category)}</h3>'
                    f"{score_fields(category)}</div>"
                    + render_table(differences)
                )
            sheet_sections.append(
                f'<section class="sheet-section"><h2>{html.escape(sheet)}</h2>'
                + "".join(category_sections)
                + "</section>"
            )
        detail = "<hr>".join(sheet_sections)
    if result.error is None:
        missing_score_categories = [
            category for category in _REPORT_CATEGORIES
            if category not in scored_categories
        ]
        if missing_score_categories:
            fallback_sections = "".join(
                f'<div class="category-heading"><h3>{html.escape(category)}</h3>'
                f'<p class="no-differences">Geen meldingen in deze categorie.</p>'
                f"{score_fields(category)}</div>"
                for category in missing_score_categories
            )
            detail += (
                '<section class="unreported-categories">'
                "<h2>Andere te beoordelen onderdelen</h2>"
                f"{fallback_sections}</section>"
            )

        score_key = html.escape(str(output_path.resolve()), quote=True)
        detail += (
            f'<section class="grading" data-score-key="{score_key}">'
            "<h2>Totaalscore</h2>"
            '<p class="score-total" aria-live="polite">— / —</p>'
            '<p class="score-status" aria-live="polite"></p>'
            "</section>"
        )
    document = f"""<!doctype html>
<html lang="nl"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Excel Checker - {title}</title>
<style>
body{{font:16px/1.5 system-ui,sans-serif;max-width:1100px;margin:2rem auto;padding:0 1rem;color:#202124}}
h1{{font-size:1.6rem}}h2{{font-size:1.35rem;margin-top:1.5rem}}h3{{font-size:1.1rem;margin-bottom:.3rem}}
h4{{font-size:1rem;margin:1rem 0 .2rem}}
.success{{color:#137333}}.error{{color:#b3261e}}hr{{border:0;border-top:2px solid #dadce0;margin:2rem 0}}
table{{border-collapse:collapse;width:100%;margin-top:1rem}}th,td{{border:1px solid #dadce0;padding:.6rem;text-align:left;vertical-align:top}}
th{{background:#f1f3f4}}code{{white-space:pre-wrap;overflow-wrap:anywhere}}
.color-swatch{{display:inline-block;width:50px;height:50px;border:1px solid #777;vertical-align:middle;margin:0 .25rem}}
.category-heading{{display:flex;align-items:center;gap:1rem;flex-wrap:wrap}}
.category-heading h3{{margin:.7rem 0}}
.inline-score{{display:inline-flex;align-items:center;gap:.45rem;font-size:.9rem}}
.inline-score label{{display:inline-flex;align-items:center;gap:.3rem}}
.inline-score input{{box-sizing:border-box;width:5.5rem;padding:.25rem;font:inherit}}
.no-differences{{margin:.4rem 0;color:#5f6368}}
.score-total{{font-weight:700}}.score-status{{min-height:1.5em;color:#b3261e}}
@media print{{.inline-score input{{border:0;padding:0}}}}
</style>
<h1>Controlerapport: {title}</h1><p><strong>Resultaat:</strong> {status}</p>
{detail}
<script>
(() => {{
  const grading = document.querySelector(".grading");
  if (!grading) return;

  const storageKey = "excel-checker-score:" + grading.dataset.scoreKey;
  const groups = [...document.querySelectorAll(".inline-score")];
  const fields = [...document.querySelectorAll(".inline-score input[data-score-field]")];
  const total = grading.querySelector(".score-total");
  const status = grading.querySelector(".score-status");

  const format = (value) => value.toLocaleString("nl-BE", {{
    maximumFractionDigits: 2
  }});

  const calculate = () => {{
    let earnedTotal = 0;
    let maximumTotal = 0;
    let valid = true;
    let complete = true;

    for (const group of groups) {{
      const earnedField = group.querySelector('[data-score-field="earned"]');
      const maximumField = group.querySelector('[data-score-field="maximum"]');
      const earnedEmpty = earnedField.value === "";
      const maximumEmpty = maximumField.value === "";
      const earned = earnedField.valueAsNumber;
      const maximum = maximumField.valueAsNumber;
      if (earnedEmpty || maximumEmpty) {{
        complete = false;
        earnedField.setCustomValidity("");
        maximumField.setCustomValidity("");
        continue;
      }}
      const rowValid =
        Number.isFinite(earned) &&
        Number.isFinite(maximum) &&
        earned >= 0 &&
        maximum >= 0 &&
        earned <= maximum;

      earnedField.setCustomValidity(
        rowValid ? "" : "Vul geldige punten in; behaald kan niet hoger zijn dan het maximum."
      );
      maximumField.setCustomValidity(
        rowValid ? "" : "Vul een geldig maximum in dat minstens gelijk is aan de behaalde punten."
      );
      valid = valid && rowValid;
      if (rowValid) {{
        earnedTotal += earned;
        maximumTotal += maximum;
      }}
    }}

    total.textContent = !valid
      ? "Controleer de punteninvoer"
      : complete
        ? `${{format(earnedTotal)}} / ${{format(maximumTotal)}}`
        : "— / —";
    status.textContent = !valid
      ? "Controleer de invoer: behaalde punten moeten tussen 0 en het maximum liggen."
      : complete
        ? "Punten worden automatisch in deze browser bewaard."
        : "Vul per onderdeel de behaalde punten en het maximum in. Gebruik 0 als het onderdeel niet meetelt.";

    try {{
      localStorage.setItem(storageKey, JSON.stringify(
        fields.map((field) => field.value)
      ));
    }} catch (error) {{
      status.textContent =
        "Opslaan in deze browser is niet beschikbaar: " + error.message;
    }}
  }};

  try {{
    const saved = localStorage.getItem(storageKey);
    if (saved !== null) {{
      const values = JSON.parse(saved);
      if (Array.isArray(values) && values.length === fields.length) {{
        fields.forEach((field, index) => {{
          if (typeof values[index] === "string") field.value = values[index];
        }});
      }} else {{
        status.textContent = "Opgeslagen scores konden niet worden hersteld.";
      }}
    }}
  }} catch (error) {{
    status.textContent =
      "Opgeslagen scores konden niet worden gelezen: " + error.message;
  }}

  fields.forEach((field) => field.addEventListener("input", calculate));
  calculate();
}})();
</script>
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
