#!/usr/bin/env python3
"""Verify workbook structure and values against simple-mode orders."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

import build_workbook
import core
import large_daily


WPS_CONDITIONAL_FORMULA_RE = re.compile(
    r'^=?LEFT\(\$[A-Z]{1,3}2,2\)="(?:少转|多转)"$',
    re.IGNORECASE,
)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("orders", type=Path)
    return parser.parse_args(argv)


def comparable(value: object) -> tuple[str, object]:
    if value is None:
        # Excel serializes an authored empty string as a blank cell. Treat both
        # representations as equivalent when checking generated workbooks.
        return ("text", "")
    if isinstance(value, datetime):
        return ("datetime", value.replace(microsecond=0))
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        try:
            return ("number", Decimal(str(value)).normalize())
        except InvalidOperation:
            pass
    return ("text", str(value))


def has_solid_fill(cell: Any, rgb: str) -> bool:
    color = cell.fill.fgColor
    return (
        cell.fill.fill_type == "solid"
        and color.type == "rgb"
        and str(color.rgb).upper().endswith(rgb)
    )


def has_summary_fill(cell: Any) -> bool:
    return has_solid_fill(cell, build_workbook.SUMMARY_FILL_RGB)


def font_rgb(cell: Any) -> str | None:
    color = cell.font.color
    if color is None or color.type != "rgb":
        return None
    return str(color.rgb).upper()


def review_conditional_colors(worksheet: Any, review_column_letter: str) -> dict[str, str]:
    colors: dict[str, str] = {}
    for conditional in worksheet.conditional_formatting:
        if review_column_letter not in str(conditional.sqref):
            continue
        for rule in conditional.rules:
            formula_text = " ".join(str(item) for item in (rule.formula or []))
            differential_font = getattr(getattr(rule, "dxf", None), "font", None)
            color = getattr(differential_font, "color", None)
            if color is None or color.type != "rgb":
                continue
            rgb = str(color.rgb).upper()
            for label in ("少转", "多转"):
                if label in formula_text:
                    colors[label] = rgb
    return colors


def wps_compatibility_errors(workbook: Any) -> list[str]:
    """Reject workbook features outside the deliberately small WPS-safe subset."""

    errors: list[str] = []
    if getattr(workbook, "_external_links", []):
        errors.append("workbook contains external links, which are forbidden for WPS output")
    if getattr(workbook, "vba_archive", None) is not None:
        errors.append("workbook contains macros, which are forbidden for WPS output")
    for worksheet in workbook.worksheets:
        for conditional in worksheet.conditional_formatting:
            for rule in conditional.rules:
                for formula in rule.formula or []:
                    formula_text = str(formula).strip()
                    if not WPS_CONDITIONAL_FORMULA_RE.fullmatch(formula_text):
                        errors.append(
                            f"{worksheet.title}: conditional formula is not WPS-compatible: "
                            f"{formula_text!r}"
                        )
    return errors


def canonical_number_format(value: object) -> str:
    """Normalize equivalent Excel literal escapes before semantic comparison."""
    return re.sub(r"\\(.)", r"\1", str(value or "")).strip().casefold()


def check_large_daily(workbook_path: Path, ledger: object) -> list[str]:
    errors: list[str] = []
    groups = build_workbook.validate_large_daily(ledger)
    used: set[str] = set()
    expected_names = [build_workbook.group_sheet_name(group, used) for group in groups]
    last_column = len(large_daily.HEADERS)
    last_column_letter = get_column_letter(last_column)
    time_column = large_daily.HEADERS.index("时间") + 1
    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    try:
        errors.extend(wps_compatibility_errors(workbook))
        if workbook.sheetnames != expected_names:
            errors.append(
                f"sheet names mismatch: expected {expected_names!r}, got {workbook.sheetnames!r}"
            )
        for group, sheet_name in zip(groups, expected_names):
            worksheet = workbook[sheet_name]
            if worksheet.sheet_state != "visible":
                errors.append(f"{sheet_name}: sheet is not visible")
            if str(worksheet.freeze_panes or "") != "A2":
                errors.append(f"{sheet_name}: first row must be frozen at A2")
            headers = [
                worksheet.cell(1, column).value
                for column in range(1, last_column + 1)
            ]
            if headers != large_daily.HEADERS:
                errors.append(f"{sheet_name}: header mismatch")
            expected_rows = build_workbook.large_group_rows(group)
            if max(0, worksheet.max_row - 1) != len(expected_rows):
                errors.append(
                    f"{sheet_name}: expected {len(expected_rows)} data rows, "
                    f"got {max(0, worksheet.max_row - 1)}"
                )
            for row_index, expected in enumerate(expected_rows, start=2):
                cells = [
                    worksheet.cell(row_index, column)
                    for column in range(1, last_column + 1)
                ]
                for column_index, (cell, value) in enumerate(zip(cells, expected), start=1):
                    if comparable(cell.value) != comparable(value):
                        errors.append(
                            f"{sheet_name}!{worksheet.cell(row_index, column_index).coordinate}: "
                            f"expected {value!r}, got {cell.value!r}"
                        )
                    if (
                        cell.alignment.horizontal != "center"
                        or cell.alignment.vertical != "center"
                    ):
                        errors.append(f"{sheet_name}!{cell.coordinate}: cell must be centered")
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        errors.append(f"{sheet_name}!{cell.coordinate}: formulas are forbidden")
                if expected[0] == "日终统计":
                    wrong_fill = [cell.coordinate for cell in cells if not has_summary_fill(cell)]
                    if wrong_fill:
                        errors.append(
                            f"{sheet_name}!{row_index}: daily summary row must be solid "
                            f"#{build_workbook.SUMMARY_FILL_RGB}"
                        )
                else:
                    filled = [cell.coordinate for cell in cells if cell.fill.fill_type is not None]
                    if filled:
                        errors.append(
                            f"{sheet_name}!{row_index}: exchange detail row must have no fill"
                        )
                time_cell = worksheet.cell(row_index, time_column)
                if (
                    time_cell.value is not None
                    and canonical_number_format(time_cell.number_format)
                    != canonical_number_format("yyyy-mm-dd hh:mm:ss")
                ):
                    errors.append(f"{sheet_name}!{time_cell.coordinate}: wrong datetime number format")
            if worksheet.max_column > last_column:
                for row in worksheet.iter_rows(
                    min_col=last_column + 1, max_col=worksheet.max_column
                ):
                    if any(cell.value is not None for cell in row):
                        errors.append(
                            f"{sheet_name}: values exist beyond column {last_column_letter}"
                        )
                        break
    finally:
        workbook.close()
    return errors


def check(workbook_path: Path, orders_path: Path) -> list[str]:
    errors: list[str] = []
    orders = json.loads(orders_path.read_text(encoding="utf-8-sig"))
    if (
        isinstance(orders, dict)
        and orders.get("contract_version") == large_daily.LEGACY_OUTPUT_CONTRACT
    ):
        return check_large_daily(workbook_path, orders)
    large_order_mode = (
        isinstance(orders, dict)
        and orders.get("contract_version") == large_daily.OUTPUT_CONTRACT
    )
    datetime_column = core.HEADERS.index("聊天消息时间") + 1
    completion_datetime_column = core.HEADERS.index("完成时间") + 1
    payee_column = core.HEADERS.index("收款方") + 1
    review_column = core.HEADERS.index("核对结果") + 1
    review_column_letter = get_column_letter(review_column)
    fund_row_types = {
        "客户付款",
        "付款退款",
        "内部回款",
        "内部回款分摊",
        "回款追回",
        "回款追回分摊",
        "未归单资金图片",
    }
    last_column = len(core.HEADERS)
    last_column_letter = get_column_letter(last_column)
    groups = build_workbook.validate_orders(orders)
    used: set[str] = set()
    expected_names = [build_workbook.group_sheet_name(group, used) for group in groups]
    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    try:
        errors.extend(wps_compatibility_errors(workbook))
        if workbook.sheetnames != expected_names:
            errors.append(f"sheet names mismatch: expected {expected_names!r}, got {workbook.sheetnames!r}")
        for worksheet in workbook.worksheets:
            if worksheet.sheet_state != "visible":
                errors.append(f"{worksheet.title}: sheet is not visible")
            if str(worksheet.freeze_panes or "") != "A2":
                errors.append(f"{worksheet.title}: first row must be frozen at A2")
            if "汇总" in worksheet.title or "总表" in worksheet.title:
                errors.append(f"{worksheet.title}: forbidden summary sheet name")
            headers = [worksheet.cell(1, column).value for column in range(1, len(core.HEADERS) + 1)]
            if headers != core.HEADERS:
                errors.append(f"{worksheet.title}: header mismatch")
            wrong_alignment: list[str] = []
            for row in worksheet.iter_rows(
                min_row=1,
                max_row=max(1, worksheet.max_row),
                min_col=1,
                max_col=last_column,
            ):
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue
                    if (
                        cell.alignment.horizontal != "center"
                        or cell.alignment.vertical != "center"
                    ):
                        wrong_alignment.append(cell.coordinate)
            if wrong_alignment:
                errors.append(
                    f"{worksheet.title}: cells must be horizontally and vertically centered; "
                    f"wrong cells: {wrong_alignment[:20]!r}"
                )
            conditional_colors = review_conditional_colors(
                worksheet,
                review_column_letter,
            )
            expected_conditional_colors = {
                "少转": build_workbook.REVIEW_SHORT_FONT_RGB,
                "多转": build_workbook.REVIEW_OVER_FONT_RGB,
            }
            for label, expected_color in expected_conditional_colors.items():
                actual_color = conditional_colors.get(label)
                if actual_color is None or not actual_color.endswith(expected_color):
                    errors.append(
                        f"{worksheet.title}: 核对结果 '{label}' conditional font must be "
                        f"#{expected_color}"
                    )
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        errors.append(f"{worksheet.title}!{cell.coordinate}: formulas are forbidden")
        for group, sheet_name in zip(groups, expected_names):
            worksheet = workbook[sheet_name]
            expected_rows: list[list[Any]] = []
            for order in group.get("orders", []):
                expected_rows.extend(build_workbook.order_rows(order))
            actual_row_count = max(0, worksheet.max_row - 1)
            expected_summary_rows = (
                build_workbook.large_summary_rows(group)
                if large_order_mode
                else []
            )
            expected_row_count = len(expected_rows) + (
                2 + len(expected_summary_rows) if large_order_mode else 0
            )
            if actual_row_count != expected_row_count:
                errors.append(
                    f"{sheet_name}: expected {expected_row_count} data/summary rows, "
                    f"got {actual_row_count}"
                )
            for row_index, expected in enumerate(expected_rows, start=2):
                row_cells = [worksheet.cell(row_index, column) for column in range(1, len(core.HEADERS) + 1)]
                actual = [cell.value for cell in row_cells]
                for column_index, (left, right) in enumerate(zip(actual, expected), start=1):
                    if comparable(left) != comparable(right):
                        errors.append(
                            f"{sheet_name}!{worksheet.cell(row_index, column_index).coordinate}: "
                            f"expected {right!r}, got {left!r}"
                        )
                if expected[0] in fund_row_types:
                    payee_cell = worksheet.cell(row_index, payee_column)
                    if not core.clean_text(payee_cell.value):
                        errors.append(
                            f"{sheet_name}!{payee_cell.coordinate}: fund-flow row requires payee"
                        )
                    elif payee_cell.data_type != "s":
                        errors.append(
                            f"{sheet_name}!{payee_cell.coordinate}: payee must be stored as text"
                        )
                    if canonical_number_format(payee_cell.number_format) != "@":
                        errors.append(
                            f"{sheet_name}!{payee_cell.coordinate}: payee must use Excel text format"
                        )
                if expected[0] == "订单汇总":
                    wrong_fill = [cell.coordinate for cell in row_cells if not has_summary_fill(cell)]
                    if wrong_fill:
                        errors.append(
                            f"{sheet_name}!{row_index}: order summary row must be solid "
                            f"#{build_workbook.SUMMARY_FILL_RGB}; wrong cells: {wrong_fill!r}"
                        )
                    review_value = core.clean_text(expected[review_column - 1])
                    expected_review_color = None
                    if review_value.startswith("少转"):
                        expected_review_color = build_workbook.REVIEW_SHORT_FONT_RGB
                    elif review_value.startswith("多转"):
                        expected_review_color = build_workbook.REVIEW_OVER_FONT_RGB
                    if expected_review_color:
                        review_cell = worksheet.cell(row_index, review_column)
                        actual_review_color = font_rgb(review_cell)
                        if (
                            actual_review_color is None
                            or not actual_review_color.endswith(expected_review_color)
                        ):
                            errors.append(
                                f"{sheet_name}!{review_cell.coordinate}: wrong review-result font color"
                            )
                else:
                    filled = [cell.coordinate for cell in row_cells if cell.fill.fill_type is not None]
                    if filled:
                        errors.append(
                            f"{sheet_name}!{row_index}: fund-flow detail row must have no fill; "
                            f"filled cells: {filled!r}"
                        )
                if (
                    worksheet.cell(row_index, datetime_column).value is not None
                    and canonical_number_format(
                        worksheet.cell(row_index, datetime_column).number_format
                    )
                    != canonical_number_format("yyyy-mm-dd hh:mm:ss")
                ):
                    coordinate = worksheet.cell(row_index, datetime_column).coordinate
                    errors.append(f"{sheet_name}!{coordinate}: wrong datetime number format")
                if (
                    worksheet.cell(row_index, completion_datetime_column).value is not None
                    and canonical_number_format(
                        worksheet.cell(row_index, completion_datetime_column).number_format
                    )
                    != canonical_number_format("yyyy-mm-dd hh:mm:ss")
                ):
                    coordinate = worksheet.cell(
                        row_index, completion_datetime_column
                    ).coordinate
                    errors.append(f"{sheet_name}!{coordinate}: wrong completion datetime number format")
            if large_order_mode:
                title_row = len(expected_rows) + 2
                summary_header_row = title_row + 1
                expected_title = "资金汇总"
                title_cell = worksheet.cell(title_row, 1)
                if title_cell.value != expected_title:
                    errors.append(
                        f"{sheet_name}!A{title_row}: daily summary title mismatch"
                    )
                summary_last_column = get_column_letter(len(large_daily.SUMMARY_HEADERS))
                expected_merge = f"A{title_row}:{summary_last_column}{title_row}"
                if expected_merge not in {
                    str(merged_range) for merged_range in worksheet.merged_cells.ranges
                }:
                    errors.append(
                        f"{sheet_name}!{title_row}: daily summary title must merge A:{summary_last_column}"
                    )
                if not has_solid_fill(
                    title_cell,
                    build_workbook.LARGE_SUMMARY_TITLE_FILL_RGB,
                ):
                    errors.append(
                        f"{sheet_name}!A{title_row}: daily summary title fill mismatch"
                    )
                title_font_rgb = font_rgb(title_cell)
                if (
                    not title_cell.font.bold
                    or title_font_rgb is None
                    or not title_font_rgb.endswith(
                        build_workbook.LARGE_SUMMARY_WHITE_FONT_RGB
                    )
                ):
                    errors.append(
                        f"{sheet_name}!A{title_row}: daily summary title font must be bold white"
                    )
                trailing_title_values = [
                    worksheet.cell(title_row, column).value
                    for column in range(len(large_daily.SUMMARY_HEADERS) + 1, last_column + 1)
                ]
                if any(value is not None for value in trailing_title_values):
                    errors.append(
                        f"{sheet_name}!{title_row}: values follow the daily summary title"
                    )
                summary_headers = [
                    worksheet.cell(summary_header_row, column).value
                    for column in range(1, len(large_daily.SUMMARY_HEADERS) + 1)
                ]
                if summary_headers != large_daily.SUMMARY_HEADERS:
                    errors.append(f"{sheet_name}!{summary_header_row}: daily summary header mismatch")
                for column in range(1, len(large_daily.SUMMARY_HEADERS) + 1):
                    header_cell = worksheet.cell(summary_header_row, column)
                    header_font_rgb = font_rgb(header_cell)
                    if not has_solid_fill(
                        header_cell,
                        build_workbook.LARGE_SUMMARY_HEADER_FILL_RGB,
                    ):
                        errors.append(
                            f"{sheet_name}!{header_cell.coordinate}: daily summary header fill mismatch"
                        )
                    if (
                        not header_cell.font.bold
                        or header_font_rgb is None
                        or not header_font_rgb.endswith(
                            build_workbook.LARGE_SUMMARY_WHITE_FONT_RGB
                        )
                    ):
                        errors.append(
                            f"{sheet_name}!{header_cell.coordinate}: daily summary header font must be bold white"
                        )
                trailing_headers = [
                    worksheet.cell(summary_header_row, column).value
                    for column in range(len(large_daily.SUMMARY_HEADERS) + 1, last_column + 1)
                ]
                if any(value is not None for value in trailing_headers):
                    errors.append(
                        f"{sheet_name}!{summary_header_row}: values follow the daily summary header"
                    )
                for offset, expected in enumerate(expected_summary_rows, start=1):
                    row_index = summary_header_row + offset
                    cells = [
                        worksheet.cell(row_index, column)
                        for column in range(1, last_column + 1)
                    ]
                    actual = [cell.value for cell in cells]
                    padded_expected = expected + [None] * (last_column - len(expected))
                    for column_index, (left, right) in enumerate(
                        zip(actual, padded_expected), start=1
                    ):
                        if comparable(left) != comparable(right):
                            errors.append(
                                f"{sheet_name}!{worksheet.cell(row_index, column_index).coordinate}: "
                                f"expected {right!r}, got {left!r}"
                            )
                    expected_fill_rgb = (
                        build_workbook.LARGE_SUMMARY_BODY_FILL_RGB
                        if offset % 2 == 1
                        else build_workbook.LARGE_SUMMARY_BODY_ALT_FILL_RGB
                    )
                    wrong_fill = [
                        cell.coordinate
                        for cell in cells[: len(large_daily.SUMMARY_HEADERS)]
                        if not has_solid_fill(cell, expected_fill_rgb)
                    ]
                    if wrong_fill:
                        errors.append(
                            f"{sheet_name}!{row_index}: daily summary row fill mismatch; "
                            f"wrong cells: {wrong_fill!r}"
                        )
                    for amount_column in (6, 7, 8):
                        amount_cell = worksheet.cell(row_index, amount_column)
                        if not amount_cell.font.bold:
                            errors.append(
                                f"{sheet_name}!{amount_cell.coordinate}: summary count/amount must be bold"
                            )
                minimum_widths = {
                    "A": 14,
                    "B": 18,
                    "C": 23,
                    "D": 16,
                    "E": 13,
                    "F": 11,
                    "G": 18,
                    "H": 18,
                    "I": 14,
                }
                for column_letter, minimum_width in minimum_widths.items():
                    actual_width = worksheet.column_dimensions[column_letter].width or 0
                    if actual_width < minimum_width:
                        errors.append(
                            f"{sheet_name}!{column_letter}: summary column width must be at least {minimum_width}"
                        )
            if worksheet.max_column > last_column:
                for row in worksheet.iter_rows(min_col=last_column + 1, max_col=worksheet.max_column):
                    if any(cell.value is not None for cell in row):
                        errors.append(f"{sheet_name}: values exist beyond column {last_column_letter}")
                        break
    finally:
        workbook.close()
    return errors


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        errors = check(
            args.workbook.resolve(),
            args.orders.resolve(),
        )
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(f"Workbook check failed: {exc}", file=sys.stderr)
        return 2
    if errors:
        print(json.dumps({"status": "ERROR", "errors": errors[:100], "total_errors": len(errors)}, ensure_ascii=False, indent=2))
        return 3
    print("OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
