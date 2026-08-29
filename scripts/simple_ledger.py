#!/usr/bin/env python3
"""Compile image-based customer payments and staff payouts with minimal rules."""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP
from pathlib import Path
from typing import Any, Mapping

import core


LEGACY_SIMPLE_PLAN_CONTRACT = "small-group-simple-plan/1.0"
INTERMEDIATE_SIMPLE_PLAN_CONTRACT = "small-group-simple-plan/1.1"
SIMPLE_PLAN_CONTRACT = "small-group-simple-plan/1.2"
SUPPORTED_SIMPLE_PLAN_CONTRACTS = {
    LEGACY_SIMPLE_PLAN_CONTRACT,
    INTERMEDIATE_SIMPLE_PLAN_CONTRACT,
    SIMPLE_PLAN_CONTRACT,
}
SIMPLE_MODE_VERSION = "1.3"
SUPPORTED_FUND_EVENT_TYPES = frozenset(core.FUND_EVENT_TYPES)
FLOW_LABEL = {
    "payment": "客户付款",
    "payment_refund": "付款退款",
    "payout": "内部回款",
    "recovery": "回款追回",
}
FLOW_SIDES = frozenset(FLOW_LABEL)
MODEL_FLOW_SIDES = FLOW_SIDES | {"unknown"}
RATE_OPERATORS = frozenset({"multiply", "divide"})
FEE_KINDS = frozenset({"delivery_fee", "service_fee", "network_fee"})
NETWORK_FEE_TREATMENTS = frozenset({"added_to_payment", "deducted_from_payout"})
FEE_TREATMENTS = frozenset(
    {
        "added_to_payment",
        "added_to_payout",
        "deducted_from_payout",
        "included_in_quote",
        "separate",
    }
)
ROUNDING_MODES = {
    "half_up": ROUND_HALF_UP,
    "down": ROUND_DOWN,
    "up": ROUND_UP,
}


def _simple_side(side_override: str | None) -> str:
    side = core.clean_text(side_override).casefold()
    if not side:
        return "unknown"
    core.require(side in MODEL_FLOW_SIDES, f"unsupported simple flow side: {side}")
    return side


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("normalized", type=Path)
    parser.add_argument("--events", required=True, type=Path)
    parser.add_argument("--plan", required=True, type=Path)
    parser.add_argument("-o", "--output", required=True, type=Path)
    parser.add_argument("--force", action="store_true")
    return parser.parse_args(argv)


def _cash_amount(event: Mapping[str, Any]) -> Decimal | None:
    ocr = event.get("ocr") if isinstance(event.get("ocr"), Mapping) else {}
    for raw in (ocr.get("amount_text"), ocr.get("amount")):
        if raw in (None, ""):
            continue
        try:
            parsed = core.parse_cash_amount(
                raw,
                field=f"{event.get('event_id') or 'cash_event'}.amount",
                allow_none=True,
            )
        except ValueError:
            continue
        return abs(parsed) if parsed is not None else None
    return None


def _visible_amount(event: Mapping[str, Any]) -> Decimal | None:
    if event.get("type") in {"cash_payment", "cash_payout"}:
        return _cash_amount(event)
    ocr = event.get("ocr") if isinstance(event.get("ocr"), Mapping) else {}
    try:
        return core.parse_decimal(
            ocr.get("amount"),
            field=f"{event.get('event_id') or 'fund_event'}.amount",
            allow_none=True,
        )
    except ValueError:
        return None


def _visible_currency(event: Mapping[str, Any]) -> str | None:
    ocr = event.get("ocr") if isinstance(event.get("ocr"), Mapping) else {}
    try:
        return core.normalize_currency(
            ocr.get("currency"),
            field=f"{event.get('event_id') or 'fund_event'}.currency",
            allow_none=True,
        )
    except ValueError:
        return None


def _visible_payee(event: Mapping[str, Any]) -> str:
    ocr = event.get("ocr") if isinstance(event.get("ocr"), Mapping) else {}
    return core.validate_payee(
        ocr.get("payee"),
        field=f"{event.get('event_id') or 'fund_event'}.payee",
        cash=event.get("type") in {"cash_payment", "cash_payout"},
    )


def _flow_from_event(
    event: Mapping[str, Any],
    message: Mapping[str, Any],
    *,
    duplicate_of: str | None,
    duplicate_basis: str | None = None,
    side_override: str | None = None,
) -> tuple[dict[str, Any], str | None]:
    event_id = str(event.get("event_id") or "")
    event_type = str(event.get("type") or "")
    side = _simple_side(side_override)
    ocr = event.get("ocr") if isinstance(event.get("ocr"), Mapping) else {}
    amount = _visible_amount(event)
    currency = _visible_currency(event)
    payee = _visible_payee(event)
    status = core.classify_fund_status(event)
    completeness = core.clean_text(ocr.get("amount_completeness")).casefold()
    if not completeness:
        completeness = "complete" if amount is not None else "unreadable"
    confidence = core.clean_text(ocr.get("confidence")).casefold()
    if not confidence:
        confidence = "high" if amount is not None and currency else "low"

    issue: str | None = None
    included = False
    display_in_workbook = not duplicate_of and status != "failed"
    # Payee is preserved as evidence data, including 未显示/无法辨认, but does
    # not control whether the order is complete.
    if not display_in_workbook:
        pass
    elif side == "unknown":
        issue = f"{event_id}: model left the fund-flow side unknown"
    elif status != "completed":
        issue = f"{event_id}: model marked fund result as {status}; not counted"
    elif amount is None or currency is None or completeness != "complete" or confidence != "high":
        issue = f"{event_id}: amount or currency is not clearly readable"
    else:
        included = True

    return (
        {
            "event_id": event_id,
            "side": side,
            "flow_type": FLOW_LABEL.get(side, side),
            "amount": core.decimal_text(amount),
            "currency": currency,
            "payee": payee,
            "message_time": message.get("timestamp"),
            "source_message_id": str(event.get("message_id") or ""),
            "status": status,
            "included": included,
            "duplicate_of": duplicate_of,
            "duplicate_basis": duplicate_basis if duplicate_of else None,
            "display_in_workbook": display_in_workbook,
            "leg_id": None,
            "leg_direction": None,
        },
        issue,
    )


def _unique_value(values: list[str]) -> str | None:
    unique = list(dict.fromkeys(value for value in values if value))
    return unique[0] if len(unique) == 1 else None


def _side_total(flows: list[dict[str, Any]], side: str) -> tuple[Decimal | None, bool]:
    included = [flow for flow in flows if flow["side"] == side and flow["included"]]
    unresolved = [
        flow
        for flow in flows
        if flow["side"] == side
        and not flow["included"]
        and flow.get("status") != "failed"
        and flow.get("duplicate_basis") is None
    ]
    if not included or unresolved:
        return None, False
    return (
        sum((Decimal(str(flow["amount"])) for flow in included), Decimal("0")),
        True,
    )


def _optional_side_total(
    flows: list[dict[str, Any]],
    side: str,
) -> tuple[Decimal | None, bool]:
    relevant = [flow for flow in flows if flow["side"] == side]
    unresolved = [
        flow
        for flow in relevant
        if not flow["included"]
        and flow.get("status") != "failed"
        and flow.get("duplicate_basis") is None
    ]
    if unresolved:
        return None, False
    return (
        sum(
            (
                Decimal(str(flow["amount"]))
                for flow in relevant
                if flow["included"]
            ),
            Decimal("0"),
        ),
        True,
    )


def _net_side_total(
    flows: list[dict[str, Any]],
    primary_side: str,
    subtract_side: str,
) -> tuple[Decimal | None, bool]:
    primary, primary_known = _side_total(flows, primary_side)
    subtraction, subtraction_known = _optional_side_total(flows, subtract_side)
    if not primary_known or not subtraction_known or primary is None or subtraction is None:
        return None, False
    value = primary - subtraction
    return (value, value >= 0)


def _parse_positive_decimal(value: object, *, field: str) -> Decimal:
    number = core.parse_decimal(value, field=field)
    assert number is not None
    core.require(number > 0, f"{field} must be positive")
    return number


def _append_note(existing: str, note: str) -> str:
    clean_note = core.clean_text(note)
    if not clean_note:
        return core.clean_text(existing)
    parts = [item for item in core.clean_text(existing).split("\n") if item]
    if clean_note not in parts:
        parts.append(clean_note)
    return "\n".join(parts)


def _rate_operator(value: object, *, field: str, required: bool = False) -> str:
    raw = core.clean_text(value).casefold()
    core.require(not required or bool(raw), f"{field} is required when rate is used to calculate expected payout")
    operator = raw or "multiply"
    core.require(operator in RATE_OPERATORS, f"{field} must be multiply or divide")
    return operator


def _compile_fees(
    value: object,
    *,
    payment_currency: str,
    payout_currency: str,
    field: str,
) -> tuple[list[dict[str, str]], Decimal, Decimal]:
    if value in (None, ""):
        return [], Decimal("0"), Decimal("0")
    core.require(isinstance(value, list), f"{field} must be a list")
    compiled: list[dict[str, str]] = []
    payment_deduction = Decimal("0")
    payout_adjustment = Decimal("0")
    for index, item in enumerate(value):
        item_field = f"{field}[{index}]"
        core.require(isinstance(item, Mapping), f"{item_field} must be an object")
        kind = core.clean_text(item.get("kind")).casefold()
        core.require(kind in FEE_KINDS, f"{item_field}.kind is unsupported")
        if kind == "network_fee" and item.get("customer_requested") is not True:
            continue
        treatment = core.clean_text(item.get("treatment")).casefold()
        core.require(
            treatment in FEE_TREATMENTS,
            f"{item_field}.treatment is unsupported",
        )
        if kind == "network_fee":
            core.require(
                treatment in NETWORK_FEE_TREATMENTS,
                f"{item_field}.treatment must deduct a customer-requested network fee "
                "from payment or payout",
            )
        amount = _parse_positive_decimal(item.get("amount"), field=f"{item_field}.amount")
        currency = core.normalize_currency(item.get("currency"), field=f"{item_field}.currency")
        assert currency is not None
        if treatment == "added_to_payment":
            core.require(
                currency == payment_currency,
                f"{item_field} added_to_payment must use {payment_currency}",
            )
            payment_deduction += amount
        elif treatment in {"added_to_payout", "deducted_from_payout"}:
            core.require(
                currency == payout_currency,
                f"{item_field} {treatment} must use {payout_currency}",
            )
            payout_adjustment += amount if treatment == "added_to_payout" else -amount
        compiled.append(
            {
                "kind": kind,
                "amount": core.decimal_text(amount),
                "currency": currency,
                "treatment": treatment,
            }
        )
    return compiled, payment_deduction, payout_adjustment


def _compile_rounding(
    value: object,
    *,
    payout_currency: str,
    field: str,
) -> tuple[dict[str, str], bool]:
    if value in (None, ""):
        return {"unit": "0.01", "mode": "half_up", "currency": payout_currency}, False
    core.require(isinstance(value, Mapping), f"{field} must be an object")
    unit = _parse_positive_decimal(value.get("unit"), field=f"{field}.unit")
    mode = core.clean_text(value.get("mode")).casefold() or "half_up"
    core.require(mode in ROUNDING_MODES, f"{field}.mode is unsupported")
    currency = core.normalize_currency(
        value.get("currency") or payout_currency,
        field=f"{field}.currency",
    )
    core.require(currency == payout_currency, f"{field}.currency must use {payout_currency}")
    return {
        "unit": core.decimal_text(unit),
        "mode": mode,
        "currency": currency,
    }, True


def _round_to_unit(value: Decimal, rounding: Mapping[str, str]) -> Decimal:
    unit = Decimal(str(rounding["unit"]))
    mode = ROUNDING_MODES[str(rounding["mode"])]
    return (value / unit).quantize(Decimal("1"), rounding=mode) * unit


def _pricing_result(
    *,
    payment_total: Decimal,
    payment_currency: str,
    payout_currency: str,
    rate_value: object,
    operator_value: object,
    explicit_expected_value: object,
    fees_value: object,
    rounding_value: object,
    field: str,
) -> tuple[
    Decimal | None,
    Decimal | None,
    str,
    list[dict[str, str]],
    dict[str, str] | None,
    str | None,
]:
    rate = _parse_optional_decimal(rate_value, field=f"{field}.rate")
    if rate is not None:
        core.require(rate > 0, f"{field}.rate must be positive")
    explicit_expected = _parse_optional_decimal(
        explicit_expected_value,
        field=f"{field}.expected_payout",
    )
    if explicit_expected is not None:
        core.require(explicit_expected >= 0, f"{field}.expected_payout cannot be negative")
    operator_supplied = bool(core.clean_text(operator_value))
    operator = _rate_operator(
        operator_value,
        field=f"{field}.rate_operator",
        required=False,
    )
    fees, payment_deduction, payout_adjustment = _compile_fees(
        fees_value,
        payment_currency=payment_currency,
        payout_currency=payout_currency,
        field=f"{field}.fees",
    )
    rounding, explicit_rounding = _compile_rounding(
        rounding_value,
        payout_currency=payout_currency,
        field=f"{field}.rounding",
    )
    payment_basis = payment_total - payment_deduction
    core.require(payment_basis >= 0, f"{field} fees exceed customer payment")
    expected: Decimal | None = None
    pricing_issue: str | None = None
    if explicit_expected is not None:
        expected = explicit_expected
    if rate is not None and (explicit_expected is None or operator_supplied):
        if operator == "multiply":
            base_expected = payment_basis * rate
        else:
            base_expected = payment_basis / rate
        adjusted = base_expected + payout_adjustment
        core.require(adjusted >= 0, f"{field} fees exceed expected payout")
        calculated_expected = _round_to_unit(adjusted, rounding)
        if explicit_expected is None:
            expected = calculated_expected
        elif calculated_expected != explicit_expected:
            pricing_issue = (
                f"明确应回 {core.decimal_text(explicit_expected)} 与汇率计算 "
                f"{core.decimal_text(calculated_expected)} 不一致"
            )
    return (
        expected,
        rate,
        operator,
        fees,
        rounding if explicit_rounding else None,
        pricing_issue,
    )


def _rate_display(
    direction: str,
    rate: Decimal | None,
    operator: str,
    *,
    labeled: bool,
) -> str | None:
    if rate is None:
        return None
    token = core.decimal_text(rate)
    if operator == "divide":
        token = f"÷{token}"
    return f"{direction}：{token}" if labeled else token


def _event_side_overrides(
    raw_order: Mapping[str, Any],
    event_ids: list[str],
    *,
    require_complete: bool,
) -> dict[str, str]:
    value = raw_order.get("event_sides", {})
    core.require(isinstance(value, Mapping), "simple_order.event_sides must be an object")
    overrides = {str(key): core.clean_text(side).casefold() for key, side in value.items()}
    unknown = sorted(set(overrides) - set(event_ids))
    core.require(not unknown, f"simple_order.event_sides cites events outside the order: {unknown}")
    for event_id, side in overrides.items():
        core.require(side in MODEL_FLOW_SIDES, f"simple_order.event_sides[{event_id}] is unsupported")
    if require_complete:
        missing = sorted(set(event_ids) - set(overrides))
        core.require(
            not missing,
            f"simple_order requires an explicit side for every event; missing: {missing}",
        )
    return overrides


def _same_transaction_map(
    raw_order: Mapping[str, Any],
    event_ids: list[str],
    event_index: Mapping[str, Mapping[str, Any]],
) -> dict[str, str]:
    value = raw_order.get("same_transactions", [])
    core.require(isinstance(value, list), "simple_order.same_transactions must be a list")
    mapping: dict[str, str] = {}
    event_set = set(event_ids)
    for index, item in enumerate(value):
        field = f"simple_order.same_transactions[{index}]"
        core.require(isinstance(item, Mapping), f"{field} must be an object")
        event_id = str(item.get("event_id") or "")
        same_as = str(item.get("same_as") or "")
        core.require(event_id in event_set and same_as in event_set, f"{field} must cite two order events")
        core.require(event_id != same_as, f"{field} cannot cite itself")
        core.require(event_id not in mapping, f"{field} repeats a duplicate event")
        mapping[event_id] = same_as
    core.require(not (set(mapping) & set(mapping.values())), "simple_order.same_transactions cannot contain chains")
    for event_id, same_as in mapping.items():
        left = event_index[event_id]
        right = event_index[same_as]
        left_amount = _visible_amount(left)
        right_amount = _visible_amount(right)
        left_currency = _visible_currency(left)
        right_currency = _visible_currency(right)
        core.require(
            left_amount is None or right_amount is None or left_amount == right_amount,
            f"same transaction amounts differ: {event_id} vs {same_as}",
        )
        core.require(
            left_currency is None or right_currency is None or left_currency == right_currency,
            f"same transaction currencies differ: {event_id} vs {same_as}",
        )
    return mapping


def _review_result(actual: Decimal, expected: Decimal | None, currency: str | None) -> str:
    if expected is None or not currency:
        return ""
    try:
        tolerance = core.currency_tolerance(currency)
    except ValueError:
        return ""
    difference = actual - expected
    if abs(difference) <= tolerance:
        return ""
    label = "多转" if difference > 0 else "少转"
    return f"{label} {core.decimal_text(abs(difference))} {currency}"


def _parse_optional_decimal(value: object, *, field: str) -> Decimal | None:
    return core.parse_decimal(value, field=field, allow_none=True)


def _assign_order_ids(orders: list[dict[str, Any]]) -> None:
    per_date: defaultdict[str, int] = defaultdict(int)
    orders.sort(key=lambda item: (str(item.get("start_time") or ""), str(item.get("case_id") or "")))
    for order in orders:
        date = str(order.get("start_time") or "0000-00-00")[:10].replace("-", "")
        per_date[date] += 1
        order["order_id"] = f"{date}-{per_date[date]:03d}"


def _compile_order(
    raw_order: Mapping[str, Any],
    *,
    group_key: str,
    position: int,
    messages: Mapping[str, Mapping[str, Any]],
    event_index: Mapping[str, Mapping[str, Any]],
    assigned_events: set[str],
    require_model_judgments: bool,
) -> tuple[dict[str, Any], list[str]]:
    event_ids_value = raw_order.get("event_ids")
    core.require(
        isinstance(event_ids_value, list) and event_ids_value,
        f"{group_key}: simple order {position} requires event_ids",
    )
    event_ids = [str(item) for item in event_ids_value]
    core.require(
        len(event_ids) == len(set(event_ids)),
        f"{group_key}: simple order {position} repeats an event_id",
    )
    if require_model_judgments:
        for field in ("customer_id", "customer_nickname", "direction"):
            core.require(field in raw_order, f"simple_order.{field} must be explicitly supplied by the model")
    side_overrides = _event_side_overrides(
        raw_order,
        event_ids,
        require_complete=require_model_judgments,
    )
    semantic_duplicates = _same_transaction_map(raw_order, event_ids, event_index)
    flows: list[dict[str, Any]] = []
    issues: list[str] = []
    for event_id in event_ids:
        core.require(event_id not in assigned_events, f"fund event assigned to multiple simple orders: {event_id}")
        event = event_index.get(event_id)
        core.require(event is not None, f"unknown simple order event_id: {event_id}")
        core.require(event.get("group_key") == group_key, f"simple order event belongs to another group: {event_id}")
        core.require(
            event.get("type") in SUPPORTED_FUND_EVENT_TYPES,
            f"simple order event is not supported fund evidence: {event_id}",
        )
        message_id = str(event.get("message_id") or "")
        message = messages.get(message_id)
        core.require(message is not None, f"simple order event has no source message: {event_id}")
        duplicate_of = semantic_duplicates.get(event_id)
        duplicate_basis = "declared_same_transaction" if duplicate_of else None
        flow, issue = _flow_from_event(
            event,
            message,
            duplicate_of=duplicate_of,
            duplicate_basis=duplicate_basis,
            side_override=side_overrides.get(event_id),
        )
        flows.append(flow)
        if issue:
            issues.append(issue)
        assigned_events.add(event_id)

    flows.sort(key=lambda item: (str(item.get("message_time") or ""), item["event_id"]))
    start_message_id = core.clean_text(raw_order.get("start_message_id"))
    if start_message_id:
        core.require(start_message_id in messages, f"unknown simple order start_message_id: {start_message_id}")
        core.require(messages[start_message_id].get("_group_key") == group_key, "simple order start message belongs to another group")
        start_message = messages[start_message_id]
    else:
        first_flow = min(flows, key=lambda item: (str(item.get("message_time") or ""), item["event_id"]))
        start_message_id = str(first_flow.get("source_message_id") or "")
        start_message = messages[start_message_id]

    flow_index = {str(flow["event_id"]): flow for flow in flows}
    for event_id, same_as in semantic_duplicates.items():
        core.require(
            flow_index[event_id]["side"] == flow_index[same_as]["side"],
            f"same transaction sides differ: {event_id} vs {same_as}",
        )

    customer_id = core.clean_text(raw_order.get("customer_id"))
    customer_nickname = core.clean_text(raw_order.get("customer_nickname"))
    identity_pending = not customer_id or not customer_nickname
    if identity_pending:
        issues.append(f"{group_key}: simple order {position} customer identity is not unique")

    relevant_flows = [
        flow
        for flow in flows
        if not flow.get("duplicate_of") and flow.get("status") != "failed"
    ]
    payment_currencies = [
        str(flow.get("currency") or "")
        for flow in relevant_flows
        if flow["side"] in {"payment", "payment_refund"}
    ]
    payout_currencies = [
        str(flow.get("currency") or "")
        for flow in relevant_flows
        if flow["side"] in {"payout", "recovery"}
    ]
    payment_currency_values = {value for value in payment_currencies if value}
    payout_currency_values = {value for value in payout_currencies if value}
    payment_currency = _unique_value(payment_currencies)
    payout_currency = _unique_value(payout_currencies)
    payment_total, payment_known = _net_side_total(
        flows,
        "payment",
        "payment_refund",
    )
    if payment_total is not None and payment_total < 0:
        payment_total, payment_known = None, False

    raw_legs = raw_order.get("legs")
    multi_leg = raw_legs not in (None, "")
    order_note = core.clean_text(raw_order.get("note"))
    compiled_legs: list[dict[str, Any]] = []
    fee_adjustments: list[dict[str, str]] = []
    expected: Decimal | None = None
    payout_total: Decimal | None = None
    rate: Decimal | None = None
    rate_operator = "multiply"
    rate_display: str | None = None
    rounding: dict[str, str] | None = None
    result = ""
    direction = ""
    pair_complete = False
    pricing_pending = False
    pricing_notes: list[str] = []

    if multi_leg:
        core.require(isinstance(raw_legs, list) and len(raw_legs) >= 2, "simple_order.legs requires at least two legs")
        leg_ids: set[str] = set()
        leg_directions: list[str] = []
        leg_payment_currencies: set[str] = set()
        referenced_settlement_ids: set[str] = set()
        allocation_total = Decimal("0")
        leg_specs: list[dict[str, Any]] = []
        for leg_position, raw_leg in enumerate(raw_legs, start=1):
            field = f"simple_order.legs[{leg_position - 1}]"
            core.require(isinstance(raw_leg, Mapping), f"{field} must be an object")
            leg_id = core.clean_text(raw_leg.get("leg_id")) or f"leg-{leg_position}"
            core.require(leg_id not in leg_ids, f"{field}.leg_id is duplicated")
            leg_ids.add(leg_id)
            leg_direction = core.canonical_direction(raw_leg.get("direction"), field=f"{field}.direction")
            leg_payment_currency, leg_payout_currency = core.direction_currencies(
                leg_direction,
                field=f"{field}.direction",
            )
            leg_directions.append(leg_direction)
            leg_payment_currencies.add(leg_payment_currency)
            allocation_amount = _parse_positive_decimal(
                raw_leg.get("allocation_amount"),
                field=f"{field}.allocation_amount",
            )
            allocation_total += allocation_amount
            payout_ids_value = raw_leg.get("payout_event_ids", [])
            recovery_ids_value = raw_leg.get("recovery_event_ids", [])
            core.require(isinstance(payout_ids_value, list), f"{field}.payout_event_ids must be a list")
            core.require(isinstance(recovery_ids_value, list), f"{field}.recovery_event_ids must be a list")
            payout_ids = [str(item) for item in payout_ids_value]
            recovery_ids = [str(item) for item in recovery_ids_value]
            core.require(len(payout_ids) == len(set(payout_ids)), f"{field}.payout_event_ids repeats an event")
            core.require(len(recovery_ids) == len(set(recovery_ids)), f"{field}.recovery_event_ids repeats an event")
            settlement_ids = payout_ids + recovery_ids
            core.require(
                set(settlement_ids) <= set(event_ids),
                f"{field} cites events outside the order",
            )
            core.require(
                not (referenced_settlement_ids & set(settlement_ids)),
                f"{field} repeats a settlement event from another leg",
            )
            referenced_settlement_ids.update(settlement_ids)
            for event_id in payout_ids:
                core.require(flow_index[event_id]["side"] == "payout", f"{field} payout event has the wrong side")
            for event_id in recovery_ids:
                core.require(flow_index[event_id]["side"] == "recovery", f"{field} recovery event has the wrong side")
            leg_specs.append(
                {
                    "raw": raw_leg,
                    "field": field,
                    "leg_id": leg_id,
                    "direction": leg_direction,
                    "payment_currency": leg_payment_currency,
                    "payout_currency": leg_payout_currency,
                    "allocation_amount": allocation_amount,
                    "payout_ids": payout_ids,
                    "recovery_ids": recovery_ids,
                }
            )
        core.require(len(leg_payment_currencies) == 1, "simple_order.legs must share one payment currency")
        payment_currency = next(iter(leg_payment_currencies))
        direction = "\n".join(leg_directions)
        payout_currency = "\n".join(spec["payout_currency"] for spec in leg_specs)
        if payment_currency_values and payment_currency_values != {payment_currency}:
            payment_known = False
            issues.append(f"{group_key}: simple order {position} payment currency conflicts with leg directions")
        allocation_closed = payment_known and payment_total is not None and allocation_total == payment_total
        if not allocation_closed:
            issues.append(f"{group_key}: simple order {position} allocation amounts do not close to net payment")
        order_settlement_ids = {
            str(flow["event_id"])
            for flow in relevant_flows
            if flow["side"] in {"payout", "recovery"}
        }
        if order_settlement_ids != referenced_settlement_ids:
            issues.append(f"{group_key}: simple order {position} leg settlement events are not completely assigned")
        all_legs_complete = allocation_closed and order_settlement_ids == referenced_settlement_ids
        rate_lines: list[str] = []
        payout_currency_lines: list[str] = []
        leg_result_lines: list[str] = []
        leg_note_lines: list[str] = []
        for spec in leg_specs:
            raw_leg = spec["raw"]
            selected = [flow_index[event_id] for event_id in spec["payout_ids"] + spec["recovery_ids"]]
            for flow in selected:
                flow["leg_id"] = spec["leg_id"]
                flow["leg_direction"] = spec["direction"]
            leg_payout, leg_payout_known = _net_side_total(selected, "payout", "recovery")
            if any(
                flow.get("currency") not in (None, spec["payout_currency"])
                for flow in selected
                if not flow.get("duplicate_of") and flow.get("status") != "failed"
            ):
                leg_payout, leg_payout_known = None, False
            leg_complete = allocation_closed and leg_payout_known and leg_payout is not None
            leg_rate = _parse_optional_decimal(raw_leg.get("rate"), field=f"{spec['field']}.rate")
            if leg_rate is not None:
                core.require(leg_rate > 0, f"{spec['field']}.rate must be positive")
            leg_operator = _rate_operator(raw_leg.get("rate_operator"), field=f"{spec['field']}.rate_operator")
            leg_explicit_expected = _parse_optional_decimal(
                raw_leg.get("expected_payout"),
                field=f"{spec['field']}.expected_payout",
            )
            if leg_explicit_expected is not None:
                core.require(
                    leg_explicit_expected >= 0,
                    f"{spec['field']}.expected_payout cannot be negative",
                )
            leg_expected: Decimal | None = None
            leg_fees: list[dict[str, str]] = []
            leg_rounding: dict[str, str] | None = None
            leg_pricing_issue: str | None = None
            if leg_complete:
                (
                    leg_expected,
                    leg_rate,
                    leg_operator,
                    leg_fees,
                    leg_rounding,
                    leg_pricing_issue,
                ) = _pricing_result(
                    payment_total=spec["allocation_amount"],
                    payment_currency=spec["payment_currency"],
                    payout_currency=spec["payout_currency"],
                    rate_value=raw_leg.get("rate"),
                    operator_value=raw_leg.get("rate_operator"),
                    explicit_expected_value=raw_leg.get("expected_payout"),
                    fees_value=raw_leg.get("fees"),
                    rounding_value=raw_leg.get("rounding"),
                    field=spec["field"],
                )
            else:
                (
                    leg_fees,
                    leg_payment_deduction,
                    _leg_payout_adjustment,
                ) = _compile_fees(
                    raw_leg.get("fees"),
                    payment_currency=spec["payment_currency"],
                    payout_currency=spec["payout_currency"],
                    field=f"{spec['field']}.fees",
                )
                core.require(
                    leg_payment_deduction <= spec["allocation_amount"],
                    f"{spec['field']} fees exceed allocated customer payment",
                )
                leg_rounding_rule, explicit_leg_rounding = _compile_rounding(
                    raw_leg.get("rounding"),
                    payout_currency=spec["payout_currency"],
                    field=f"{spec['field']}.rounding",
                )
                leg_rounding = leg_rounding_rule if explicit_leg_rounding else None
            leg_result = (
                _review_result(leg_payout, leg_expected, spec["payout_currency"])
                if leg_complete and leg_payout is not None
                else "待确认"
            )
            leg_issue_notes: list[str] = []
            if leg_pricing_issue:
                pricing_pending = True
                leg_result = "待确认"
                leg_issue_notes.append(leg_pricing_issue)
                pricing_notes.append(f"{spec['direction']}：{leg_pricing_issue}。")
                issues.append(f"{spec['field']}: {leg_pricing_issue}")
            if not leg_complete:
                leg_issue_notes.append("该换汇明细的付款分配或内部回款尚未完整确认。")
            leg_anomaly = "\n".join(leg_issue_notes)
            rate_line = _rate_display(
                spec["direction"],
                leg_rate,
                leg_operator,
                labeled=True,
            )
            if rate_line:
                rate_lines.append(rate_line)
            payout_currency_lines.append(spec["payout_currency"])
            if leg_result:
                leg_result_lines.append(f"{spec['direction']}：{leg_result}")
            if leg_anomaly:
                leg_note_lines.append(f"{spec['direction']}：{leg_anomaly}")
            for fee in leg_fees:
                fee["leg_id"] = spec["leg_id"]
                fee["direction"] = spec["direction"]
            fee_adjustments.extend(leg_fees)
            compiled_legs.append(
                {
                    "leg_id": spec["leg_id"],
                    "direction": spec["direction"],
                    "payment_currency": spec["payment_currency"],
                    "payment_total": core.decimal_text(spec["allocation_amount"]),
                    "actual_rate": core.decimal_text(leg_rate),
                    "rate_operator": leg_operator,
                    "actual_rate_display": _rate_display(
                        spec["direction"],
                        leg_rate,
                        leg_operator,
                        labeled=False,
                    ),
                    "payout_currency": spec["payout_currency"],
                    "expected_payout": core.decimal_text(leg_expected),
                    "actual_payout_total": core.decimal_text(leg_payout),
                    "review_result": leg_result,
                    "anomaly_note": leg_anomaly,
                    "payout_event_ids": spec["payout_ids"],
                    "recovery_event_ids": spec["recovery_ids"],
                    "fees": leg_fees,
                    "rounding": leg_rounding,
                }
            )
            all_legs_complete = all_legs_complete and leg_complete
        rate_display = "\n".join(rate_lines) or None
        payout_currency = "\n".join(payout_currency_lines)
        result = "\n".join(leg_result_lines)
        pair_complete = all_legs_complete
    else:
        requested_direction = core.clean_text(raw_order.get("direction"))
        if requested_direction:
            requested_payment, requested_payout = core.direction_currencies(
                requested_direction,
                field="simple_order.direction",
            )
            payment_currency, payout_currency = requested_payment, requested_payout
            direction = f"{requested_payment}->{requested_payout}"
        if payment_currency is None or len(payment_currency_values) > 1:
            payment_total, payment_known = None, False
            issues.append(f"{group_key}: simple order {position} payment currency is missing or mixed")
        elif any(
            flow.get("currency") not in (None, payment_currency)
            for flow in relevant_flows
            if flow["side"] in {"payment", "payment_refund"}
        ):
            payment_total, payment_known = None, False
            issues.append(f"{group_key}: simple order {position} payment flow currency conflicts with direction")
        payout_total, payout_known = _net_side_total(flows, "payout", "recovery")
        if payout_currency is None or len(payout_currency_values) > 1:
            payout_total, payout_known = None, False
            issues.append(f"{group_key}: simple order {position} payout currency is missing or mixed")
        elif any(
            flow.get("currency") not in (None, payout_currency)
            for flow in relevant_flows
            if flow["side"] in {"payout", "recovery"}
        ):
            payout_total, payout_known = None, False
            issues.append(f"{group_key}: simple order {position} payout flow currency conflicts with direction")
        pair_complete = bool(direction) and payment_known and payout_known
        rate = _parse_optional_decimal(raw_order.get("rate"), field="simple_order.rate")
        if rate is not None:
            core.require(rate > 0, "simple_order.rate must be positive")
        rate_operator = _rate_operator(raw_order.get("rate_operator"), field="simple_order.rate_operator")
        explicit_expected = _parse_optional_decimal(
            raw_order.get("expected_payout"),
            field="simple_order.expected_payout",
        )
        if explicit_expected is not None:
            core.require(explicit_expected >= 0, "simple_order.expected_payout cannot be negative")
        if pair_complete and payment_total is not None and payout_total is not None:
            assert payment_currency is not None and payout_currency is not None
            (
                expected,
                rate,
                rate_operator,
                fee_adjustments,
                rounding,
                pricing_issue,
            ) = _pricing_result(
                payment_total=payment_total,
                payment_currency=payment_currency,
                payout_currency=payout_currency,
                rate_value=raw_order.get("rate"),
                operator_value=raw_order.get("rate_operator"),
                explicit_expected_value=raw_order.get("expected_payout"),
                fees_value=raw_order.get("fees"),
                rounding_value=raw_order.get("rounding"),
                field="simple_order",
            )
            result = _review_result(payout_total, expected, payout_currency)
            if pricing_issue:
                pricing_pending = True
                result = "待确认"
                pricing_notes.append(pricing_issue + "。")
                issues.append(f"{group_key}: simple order {position}: {pricing_issue}")
        elif payment_currency is not None and payout_currency is not None:
            (
                fee_adjustments,
                payment_deduction,
                _payout_adjustment,
            ) = _compile_fees(
                raw_order.get("fees"),
                payment_currency=payment_currency,
                payout_currency=payout_currency,
                field="simple_order.fees",
            )
            if payment_total is not None:
                core.require(payment_deduction <= payment_total, "simple_order fees exceed customer payment")
            rounding_rule, explicit_rounding = _compile_rounding(
                raw_order.get("rounding"),
                payout_currency=payout_currency,
                field="simple_order.rounding",
            )
            rounding = rounding_rule if explicit_rounding else None
        rate_display = (
            _rate_display(direction, rate, rate_operator, labeled=False)
            if rate_operator == "divide"
            else None
        )

    notes: list[str] = []
    if not payment_known:
        notes.insert(0, "客户付款或付款退款图片的金额、币种尚未完整确认。")
    if not multi_leg and not pair_complete:
        notes.append("内部回款或回款追回图片的金额、币种尚未完整确认。")
    if multi_leg and any("allocation" in issue for issue in issues):
        notes.append("付款拆分金额尚未闭合。")
    if multi_leg:
        notes.extend(leg_note_lines)
    notes.extend(pricing_notes)
    if identity_pending:
        notes.append("该订单客户无法唯一确认。")
        result = "待确认"
    if not pair_complete:
        result = "待确认"
    anomaly_note = "\n".join(dict.fromkeys(note for note in notes if note))
    if pair_complete and not identity_pending and not pricing_pending:
        order_status = "completed"
    elif pricing_pending:
        order_status = "pending_pricing"
    elif identity_pending and pair_complete:
        order_status = "pending_identity"
    else:
        order_status = "pending_evidence"

    case_id = core.clean_text(raw_order.get("case_id")) or f"{group_key}:simple:{position:03d}"
    return (
        {
            "case_id": case_id,
            "order_id": None,
            "start_time": start_message.get("timestamp"),
            "start_message_id": start_message_id,
            "customer_nickname": customer_nickname,
            "customer_id": customer_id,
            "direction": direction,
            "payment_currency": payment_currency,
            "payment_total": core.decimal_text(payment_total),
            "actual_rate": core.decimal_text(rate),
            "actual_rate_display": rate_display,
            "rate_operator": rate_operator,
            "payout_currency": payout_currency,
            "expected_payout": core.decimal_text(expected),
            "balance_adjustment": "0",
            "actual_payout_total": core.decimal_text(payout_total),
            "review_result": result,
            "anomaly_note": anomaly_note,
            "note": order_note,
            "order_status": order_status,
            "fee_adjustments": fee_adjustments,
            "rounding": rounding,
            "legs": compiled_legs,
            "flows": flows,
        },
        issues,
    )


def _mark_relationship_pending(order: dict[str, Any], note: str) -> None:
    order["review_result"] = "待确认"
    order["anomaly_note"] = _append_note(str(order.get("anomaly_note") or ""), note)
    if order.get("order_status") == "completed":
        order["order_status"] = "pending_relationship"


def _apply_balance_links(
    value: object,
    orders: list[dict[str, Any]],
    *,
    group_key: str,
) -> tuple[list[dict[str, Any]], list[str]]:
    if value in (None, ""):
        return [], []
    core.require(isinstance(value, list), f"{group_key}: balance_links must be a list")
    order_index = {str(order.get("case_id") or ""): order for order in orders}
    core.require(len(order_index) == len(orders), f"{group_key}: order case_id values must be unique")
    compiled: list[dict[str, Any]] = []
    warnings: list[str] = []
    seen_links: set[tuple[str, str, str, str, str, bool]] = set()
    consumed_source_balances: defaultdict[tuple[str, str, str], Decimal] = defaultdict(Decimal)
    declared_source_balances: defaultdict[tuple[str, str, str], Decimal] = defaultdict(Decimal)
    for pre_index, pre_item in enumerate(value):
        if not isinstance(pre_item, Mapping):
            continue
        pre_source = core.clean_text(pre_item.get("source_case_id"))
        pre_kind = core.clean_text(pre_item.get("kind")).casefold()
        try:
            pre_currency = core.normalize_currency(
                pre_item.get("currency"),
                field=f"{group_key}: balance_links[{pre_index}].currency",
            )
            pre_amount = _parse_positive_decimal(
                pre_item.get("amount"),
                field=f"{group_key}: balance_links[{pre_index}].amount",
            )
        except ValueError:
            continue
        assert pre_currency is not None
        declared_source_balances[(pre_source, pre_currency, pre_kind)] += pre_amount
    for index, item in enumerate(value):
        field = f"{group_key}: balance_links[{index}]"
        core.require(isinstance(item, Mapping), f"{field} must be an object")
        source_case_id = core.clean_text(item.get("source_case_id"))
        target_case_id = core.clean_text(item.get("target_case_id"))
        kind = core.clean_text(item.get("kind")).casefold()
        core.require(kind in {"shortfall_carryover", "overpayment_carryover"}, f"{field}.kind is unsupported")
        core.require(source_case_id in order_index, f"{field}.source_case_id is unknown")
        core.require(target_case_id in order_index, f"{field}.target_case_id is unknown")
        core.require(source_case_id != target_case_id, f"{field} cannot link an order to itself")
        amount = _parse_positive_decimal(item.get("amount"), field=f"{field}.amount")
        currency = core.normalize_currency(item.get("currency"), field=f"{field}.currency")
        assert currency is not None
        already_value = item.get("already_in_expected", False)
        core.require(isinstance(already_value, bool), f"{field}.already_in_expected must be boolean")
        source = order_index[source_case_id]
        target = order_index[target_case_id]
        reasons: list[str] = []
        link_signature = (
            source_case_id,
            target_case_id,
            kind,
            core.decimal_text(amount) or "",
            currency,
            already_value,
        )
        if link_signature in seen_links:
            reasons.append("重复的跨单补款或抵扣关系")
        else:
            seen_links.add(link_signature)
        if source.get("legs") or target.get("legs"):
            reasons.append("多方向订单暂不应用跨单余额")
        if not source.get("customer_id") or source.get("customer_id") != target.get("customer_id"):
            reasons.append("前后订单客户不一致或未确认")
        if source.get("payout_currency") != currency or target.get("payout_currency") != currency:
            reasons.append("前后订单回款币种与余额币种不一致")
        if str(source.get("start_time") or "") >= str(target.get("start_time") or ""):
            reasons.append("余额必须从较早订单指向较晚订单")
        source_expected = _parse_optional_decimal(source.get("expected_payout"), field=f"{field}.source_expected")
        source_actual = _parse_optional_decimal(source.get("actual_payout_total"), field=f"{field}.source_actual")
        target_expected = _parse_optional_decimal(target.get("expected_payout"), field=f"{field}.target_expected")
        target_actual = _parse_optional_decimal(target.get("actual_payout_total"), field=f"{field}.target_actual")
        if source_expected is None or source_actual is None or target_expected is None:
            reasons.append("前单差额或后单应回金额尚未确认")
        else:
            source_difference = source_actual - source_expected
            direction_matches = (
                source_difference < 0
                if kind == "shortfall_carryover"
                else source_difference > 0
            )
            if not direction_matches:
                reasons.append("声明方向与前单实际差额不一致")
            source_key = (source_case_id, currency, kind)
            if declared_source_balances[source_key] != abs(source_difference):
                reasons.append("同一前单的全部补抵金额未与实际差额守恒")
            if consumed_source_balances[source_key] + amount > abs(source_difference):
                reasons.append("同一前单差额被重复或超额使用")
        status = "applied"
        signed_adjustment = amount if kind == "shortfall_carryover" else -amount
        if reasons:
            status = "pending"
            detail = "跨单补款或抵扣待确认：" + "；".join(reasons) + "。"
            _mark_relationship_pending(source, detail)
            _mark_relationship_pending(target, detail)
            warnings.append(f"{field}: {'; '.join(reasons)}")
        else:
            if not already_value:
                assert target_expected is not None
                adjusted_expected = target_expected + signed_adjustment
                if adjusted_expected < 0:
                    status = "pending"
                    detail = "跨单抵扣金额超过后单应回金额。"
                    _mark_relationship_pending(source, detail)
                    _mark_relationship_pending(target, detail)
                    warnings.append(f"{field}: target expected payout would be negative")
                else:
                    target["expected_payout"] = core.decimal_text(adjusted_expected)
                    target["balance_adjustment"] = core.decimal_text(
                        Decimal(str(target.get("balance_adjustment") or "0")) + signed_adjustment
                    )
                    if target.get("review_result") != "待确认" and target_actual is not None:
                        target["review_result"] = _review_result(target_actual, adjusted_expected, currency)
            if status == "applied":
                consumed_source_balances[(source_case_id, currency, kind)] += amount
                source_order_id = str(source.get("order_id") or source_case_id)
                target_order_id = str(target.get("order_id") or target_case_id)
                if kind == "shortfall_carryover":
                    source_note = f"少转 {core.decimal_text(amount)} {currency}，约定在订单 {target_order_id} 补回。"
                    target_note = f"跨单补款：承接订单 {source_order_id} 的 {core.decimal_text(amount)} {currency}。"
                else:
                    source_note = f"多转 {core.decimal_text(amount)} {currency}，约定在订单 {target_order_id} 抵扣。"
                    target_note = f"跨单抵扣：使用订单 {source_order_id} 的 {core.decimal_text(amount)} {currency}。"
                source["anomaly_note"] = _append_note(str(source.get("anomaly_note") or ""), source_note)
                target["anomaly_note"] = _append_note(str(target.get("anomaly_note") or ""), target_note)
        compiled.append(
            {
                "source_case_id": source_case_id,
                "target_case_id": target_case_id,
                "kind": kind,
                "amount": core.decimal_text(amount),
                "currency": currency,
                "already_in_expected": already_value,
                "status": status,
            }
        )
    return compiled, warnings


def compile_simple_ledger(
    normalized: Mapping[str, Any],
    events: list[dict[str, Any]],
    events_fingerprint: str,
    plan: Mapping[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    core.require(
        plan.get("contract_version") in SUPPORTED_SIMPLE_PLAN_CONTRACTS,
        "unsupported simple plan contract_version",
    )
    core.require(
        plan.get("normalized_source_fingerprint") in (None, normalized.get("source_fingerprint")),
        "simple plan normalized_source_fingerprint is stale",
    )
    core.require(
        plan.get("events_fingerprint") in (None, events_fingerprint),
        "simple plan events_fingerprint is stale",
    )
    messages, normalized_groups = core.message_and_group_indexes(normalized)
    event_index = {str(event.get("event_id") or ""): event for event in events}
    core.require(len(event_index) == len(events), "duplicate event IDs")
    plan_groups_value = plan.get("groups")
    core.require(isinstance(plan_groups_value, list), "simple plan groups must be a list")
    plan_groups = {
        str(group.get("group_key") or ""): group
        for group in plan_groups_value
        if isinstance(group, Mapping)
    }
    core.require(len(plan_groups) == len(plan_groups_value), "duplicate or invalid simple plan group_key")
    core.require(set(plan_groups) == set(normalized_groups), "simple plan must contain every selected group exactly once")

    assigned_events: set[str] = set()
    require_model_judgments = plan.get("contract_version") == SIMPLE_PLAN_CONTRACT
    result_groups: list[dict[str, Any]] = []
    warnings: list[str] = []
    for group_key, normalized_group in normalized_groups.items():
        plan_group = plan_groups[group_key]
        raw_orders = plan_group.get("orders")
        core.require(isinstance(raw_orders, list), f"{group_key}: simple plan orders must be a list")
        compiled_orders: list[dict[str, Any]] = []
        for position, raw_order in enumerate(raw_orders, start=1):
            core.require(isinstance(raw_order, Mapping), f"{group_key}: simple order must be an object")
            compiled, issues = _compile_order(
                raw_order,
                group_key=group_key,
                position=position,
                messages=messages,
                event_index=event_index,
                assigned_events=assigned_events,
                require_model_judgments=require_model_judgments,
            )
            compiled_orders.append(compiled)
            warnings.extend(issues)

        group_fund_events = [
            event
            for event in events
            if event.get("type") in SUPPORTED_FUND_EVENT_TYPES
            and (
                str(event.get("group_key") or "") == group_key
                or messages.get(str(event.get("message_id") or ""), {}).get("_group_key") == group_key
            )
        ]
        group_unassigned = [
            event for event in group_fund_events if str(event.get("event_id") or "") not in assigned_events
        ]
        group_unassigned_ids = sorted(str(event.get("event_id") or "") for event in group_unassigned)
        core.require(
            not group_unassigned_ids,
            f"{group_key}: every fund event must be explicitly assigned by the model; unassigned: {group_unassigned_ids}",
        )

        _assign_order_ids(compiled_orders)
        compiled_balance_links, balance_warnings = _apply_balance_links(
            plan_group.get("balance_links"),
            compiled_orders,
            group_key=group_key,
        )
        warnings.extend(balance_warnings)
        result_groups.append(
            {
                "group_key": group_key,
                "platform": normalized_group.get("platform"),
                "group_name": normalized_group.get("group_name"),
                "orders": compiled_orders,
                "balance_links": compiled_balance_links,
            }
        )

    fund_events = {
        str(event.get("event_id") or "")
        for event in events
        if event.get("type") in SUPPORTED_FUND_EVENT_TYPES
    }
    unassigned = sorted(fund_events - assigned_events)
    core.require(not unassigned, f"fund events remain unassigned: {unassigned}")
    document = {
        "contract_version": core.ORDERS_CONTRACT,
        "rule_version": core.RULE_VERSION,
        "accounting_mode": "simple",
        "simple_mode_version": SIMPLE_MODE_VERSION,
        "timezone": normalized.get("timezone"),
        "source_fingerprint": normalized.get("source_fingerprint"),
        "events_fingerprint": events_fingerprint,
        "plan_fingerprint": core.fingerprint_json(plan),
        "groups": result_groups,
    }
    statistics = {
        "warnings": warnings,
        "groups": len(result_groups),
        "orders": sum(len(group["orders"]) for group in result_groups),
        "fund_images": len(fund_events),
        "assigned_fund_images": len(assigned_events & fund_events),
        "unassigned_fund_images": len(unassigned),
        "originally_unassigned_fund_images": 0,
        "visible_pending_fund_images": 0,
        "pending_orders": sum(
            order["order_status"] != "completed"
            for group in result_groups
            for order in group["orders"]
        ),
    }
    return document, statistics


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        normalized = core.load_normalized(args.normalized.resolve())
        events, events_fingerprint = core.load_events(args.events.resolve(), normalized)
        plan = json.loads(args.plan.resolve().read_text(encoding="utf-8-sig"))
        output = args.output.resolve()
        if output.exists() and not args.force:
            raise ValueError(f"output already exists (use --force): {output}")
        orders, statistics = compile_simple_ledger(
            normalized,
            events,
            events_fingerprint,
            plan,
        )
        core.atomic_json(output, orders)
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(f"Simple ledger compilation failed: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(statistics, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
