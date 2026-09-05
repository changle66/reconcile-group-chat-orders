#!/usr/bin/env python3
"""Compile full-order large-group ledgers and legacy exchange-only ledgers."""

from __future__ import annotations

import copy
from decimal import Decimal
from typing import Any, Mapping

import core


LEGACY_DECISION_CONTRACT = "group-chat-large-daily-decision/1.0"
DECISION_CONTRACT = "group-chat-large-daily-decision/2.0"
LEGACY_OUTPUT_CONTRACT = "large-group-daily-ledger/1.0"
OUTPUT_CONTRACT = "large-group-daily-orders/2.0"
FUND_TYPES = frozenset({"wechat", "alipay", "bank_card", "usdt", "cash"})
RATE_OPERATORS = frozenset({"multiply", "divide"})
LEGACY_HEADERS = [
    "记录类型",
    "时间",
    "资金类型",
    "换汇方向",
    "换出金额",
    "汇率",
    "换入金额",
    "备注",
]
HEADERS = LEGACY_HEADERS
SUMMARY_HEADERS = [
    "资金类型",
    "换汇方向",
    "计算方式",
    "汇率",
    "笔数",
    "换出合计",
    "换入合计",
    "状态",
]
FUND_TYPE_LABELS = {
    "wechat": "微信",
    "alipay": "支付宝",
    "bank_card": "银行卡",
    "usdt": "USDT",
    "cash": "现金",
}
FUND_TYPE_ORDER = tuple(FUND_TYPE_LABELS)


def normalize_fund_type(value: object, *, field: str) -> str:
    fund_type = core.clean_text(value).casefold()
    core.require(
        fund_type in FUND_TYPES,
        f"{field} must be wechat, alipay, bank_card, usdt, or cash",
    )
    return fund_type


def validate_exchanges(
    value: object,
    *,
    group_key: str,
    message_labels: set[str],
    fund_entry_ids: set[str],
) -> dict[str, Any]:
    core.require(isinstance(value, list), "exchanges must be a list")
    exchange_ids: set[str] = set()
    assigned_entries: set[str] = set()
    allowed_fields = {
        "id",
        "source_messages",
        "entry_ids",
        "fund_type",
        "direction",
        "source_amount",
        "rate",
        "operator",
        "target_amount",
        "note",
    }
    for position, exchange in enumerate(value):
        field = f"{group_key}.exchanges[{position}]"
        core.require(isinstance(exchange, dict), f"{field} must be an object")
        unknown = sorted(set(exchange) - allowed_fields)
        core.require(not unknown, f"{field}: unsupported fields: {', '.join(unknown)}")

        exchange_id = core.clean_text(exchange.get("id"))
        core.require(bool(exchange_id), f"{field}.id is required")
        core.require(exchange_id not in exchange_ids, f"{field}.id is duplicated")
        exchange_ids.add(exchange_id)
        exchange["id"] = exchange_id

        source_messages = exchange.get("source_messages")
        core.require(
            isinstance(source_messages, list) and source_messages,
            f"{field}.source_messages is required",
        )
        source_labels = [str(item) for item in source_messages]
        core.require(
            len(source_labels) == len(set(source_labels)),
            f"{field}.source_messages repeats a label",
        )
        core.require(
            set(source_labels) <= message_labels,
            f"{field}.source_messages contains an unknown label",
        )
        exchange["source_messages"] = source_labels

        raw_entries = exchange.get("entry_ids", [])
        core.require(isinstance(raw_entries, list), f"{field}.entry_ids must be a list")
        entry_ids = [str(item) for item in raw_entries]
        core.require(
            len(entry_ids) == len(set(entry_ids)),
            f"{field}.entry_ids repeats an entry",
        )
        for entry_id in entry_ids:
            core.require(entry_id in fund_entry_ids, f"{field}: unknown fund entry {entry_id}")
            core.require(
                entry_id not in assigned_entries,
                f"fund entry assigned to multiple exchanges: {entry_id}",
            )
            assigned_entries.add(entry_id)
        if "entry_ids" in exchange:
            exchange["entry_ids"] = entry_ids

        fund_type = normalize_fund_type(
            exchange.get("fund_type"), field=f"{field}.fund_type"
        )
        exchange["fund_type"] = fund_type
        exchange["direction"] = core.canonical_direction(
            exchange.get("direction"), field=f"{field}.direction"
        )

        for amount_field in ("source_amount", "target_amount"):
            amount = core.parse_decimal(
                exchange.get(amount_field), field=f"{field}.{amount_field}"
            )
            core.require(amount is not None and amount > 0, f"{field}.{amount_field} must be positive")
            exchange[amount_field] = core.decimal_text(amount)

        rate = core.parse_decimal(exchange.get("rate"), field=f"{field}.rate")
        core.require(rate is not None and rate > 0, f"{field}.rate must be positive")
        exchange["rate"] = core.decimal_text(rate)
        operator = core.clean_text(exchange.get("operator")).casefold()
        core.require(
            operator in RATE_OPERATORS,
            f"{field}.operator must be multiply or divide",
        )
        exchange["operator"] = operator
        note = core.clean_text(exchange.get("note"))
        if note:
            exchange["note"] = note
        else:
            exchange.pop("note", None)

    return {
        "exchanges": len(value),
        "assigned_exchange_entries": len(assigned_entries),
        "unassigned_exchange_entries": len(fund_entry_ids - assigned_entries),
    }


def rate_display(rate: object, operator: object) -> str:
    parsed = core.parse_decimal(rate, field="large daily rate")
    core.require(parsed is not None and parsed > 0, "large daily rate must be positive")
    normalized_operator = core.clean_text(operator).casefold()
    core.require(normalized_operator in RATE_OPERATORS, "large daily rate operator is unsupported")
    symbol = "×" if normalized_operator == "multiply" else "÷"
    return symbol + core.decimal_text(parsed)


def compile_daily_ledger(
    normalized: Mapping[str, Any],
    decisions: Mapping[str, Mapping[str, Any]],
    *,
    accounting_date: str,
) -> tuple[dict[str, Any], dict[str, int]]:
    result_groups: list[dict[str, Any]] = []
    exchange_count = 0
    summary_count = 0
    for group in normalized.get("groups", []):
        group_key = str(group.get("group_key") or "")
        decision = decisions[group_key]
        raw_exchanges = decision.get("exchanges", [])
        core.require(isinstance(raw_exchanges, list), f"{group_key}: exchanges must be a list")
        if not raw_exchanges:
            continue
        message_by_label = {
            f"S{position:05d}": message
            for position, message in enumerate(group.get("messages", []), start=1)
        }
        compiled_exchanges: list[dict[str, Any]] = []
        for raw in raw_exchanges:
            core.require(isinstance(raw, Mapping), f"{group_key}: exchange must be an object")
            sources = [str(label) for label in raw.get("source_messages", [])]
            timestamps = [
                str(message_by_label[label].get("timestamp") or "")
                for label in sources
                if label in message_by_label
            ]
            core.require(timestamps and all(timestamps), f"{group_key}: exchange has no timestamped source")
            fund_type = core.clean_text(raw.get("fund_type")).casefold()
            compiled = {
                "exchange_id": core.clean_text(raw.get("id")),
                "exchange_time": min(timestamps),
                "fund_type": fund_type,
                "fund_type_label": FUND_TYPE_LABELS[fund_type],
                "direction": core.clean_text(raw.get("direction")),
                "source_amount": core.decimal_text(
                    core.parse_decimal(raw.get("source_amount"), field="source_amount")
                ),
                "rate": core.decimal_text(
                    core.parse_decimal(raw.get("rate"), field="rate")
                ),
                "operator": core.clean_text(raw.get("operator")).casefold(),
                "target_amount": core.decimal_text(
                    core.parse_decimal(raw.get("target_amount"), field="target_amount")
                ),
                "source_messages": sources,
                "entry_ids": [str(item) for item in raw.get("entry_ids", [])],
            }
            compiled["rate_display"] = rate_display(compiled["rate"], compiled["operator"])
            note = core.clean_text(raw.get("note"))
            if note:
                compiled["note"] = note
            compiled_exchanges.append(compiled)
        compiled_exchanges.sort(
            key=lambda item: (item["exchange_time"], item["exchange_id"])
        )

        buckets: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        for exchange in compiled_exchanges:
            key = (
                exchange["fund_type"],
                exchange["direction"],
                exchange["operator"],
                exchange["rate"],
            )
            if key not in buckets:
                buckets[key] = {
                    "fund_type": exchange["fund_type"],
                    "fund_type_label": exchange["fund_type_label"],
                    "direction": exchange["direction"],
                    "operator": exchange["operator"],
                    "rate": exchange["rate"],
                    "rate_display": exchange["rate_display"],
                    "count": 0,
                    "source_total": Decimal("0"),
                    "target_total": Decimal("0"),
                }
            bucket = buckets[key]
            bucket["count"] += 1
            bucket["source_total"] += core.parse_decimal(
                exchange["source_amount"], field="summary source_amount"
            )
            bucket["target_total"] += core.parse_decimal(
                exchange["target_amount"], field="summary target_amount"
            )
        fund_order = {value: position for position, value in enumerate(FUND_TYPE_ORDER)}
        summary_items = list(enumerate(buckets.items()))
        summary_items.sort(
            key=lambda item: (fund_order[item[1][0][0]], item[0])
        )
        summaries = [bucket for _, (_, bucket) in summary_items]
        for summary in summaries:
            summary["source_total"] = core.decimal_text(summary["source_total"])
            summary["target_total"] = core.decimal_text(summary["target_total"])

        result_groups.append(
            {
                "group_key": group_key,
                "platform": group.get("platform"),
                "group_name": group.get("group_name"),
                "exchanges": compiled_exchanges,
                "summaries": summaries,
            }
        )
        exchange_count += len(compiled_exchanges)
        summary_count += len(summaries)

    core.require(bool(result_groups), "no completed large-group exchanges were recorded for this day")
    document = {
        "contract_version": LEGACY_OUTPUT_CONTRACT,
        "accounting_mode": "large_daily",
        "accounting_date": accounting_date,
        "timezone": normalized.get("timezone"),
        "source_fingerprint": normalized.get("source_fingerprint"),
        "groups": result_groups,
    }
    return document, {
        "groups": len(result_groups),
        "exchanges": exchange_count,
        "summary_rows": summary_count,
    }


def _reconciliation_pending(value: object) -> bool:
    if not isinstance(value, Mapping):
        return True
    status = core.clean_text(value.get("status")).casefold()
    if status == "pending":
        return True
    if status != "composite":
        return False
    items = value.get("items")
    return not isinstance(items, list) or any(
        not isinstance(item, Mapping)
        or _reconciliation_pending(item.get("reconciliation"))
        for item in items
    )


def _summary_scopes(order: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    legs = order.get("legs")
    if isinstance(legs, list) and legs:
        return [item for item in legs if isinstance(item, Mapping)]
    return [order]


def compile_order_summaries(
    orders: list[Mapping[str, Any]],
    *,
    group_key: str,
) -> list[dict[str, Any]]:
    """Group all order scopes while keeping pending amounts explicitly unknown."""

    buckets: dict[tuple[str, str, str, str, str], dict[str, Any]] = {}
    first_seen: dict[tuple[str, str, str, str, str], int] = {}
    sequence = 0
    for order_position, order in enumerate(orders):
        field = f"{group_key}.orders[{order_position}]"
        fund_type = normalize_fund_type(
            order.get("fund_type"), field=f"{field}.fund_type"
        )
        for scope_position, scope in enumerate(_summary_scopes(order)):
            scope_field = f"{field}.summary_scopes[{scope_position}]"
            pending = _reconciliation_pending(scope.get("reconciliation"))
            status_label = "待确认" if pending else "已确认"
            raw_direction = core.clean_text(scope.get("direction"))
            if raw_direction:
                direction = core.canonical_direction(
                    raw_direction, field=f"{scope_field}.direction"
                )
            else:
                core.require(
                    pending,
                    f"{scope_field}.direction is required for a confirmed daily summary",
                )
                direction = "待确认"
            source_amount = core.parse_decimal(
                scope.get("payment_total"),
                field=f"{scope_field}.payment_total",
                allow_none=True,
            )
            target_amount = core.parse_decimal(
                scope.get("actual_payout_total"),
                field=f"{scope_field}.actual_payout_total",
                allow_none=True,
            )
            rate = core.parse_decimal(
                scope.get("actual_rate"),
                field=f"{scope_field}.actual_rate",
                allow_none=True,
            )
            operator = core.clean_text(scope.get("rate_operator")).casefold()
            core.require(
                source_amount is None or source_amount > 0,
                f"{scope_field}.payment_total must be positive when known",
            )
            core.require(
                target_amount is None or target_amount >= 0,
                f"{scope_field}.actual_payout_total cannot be negative",
            )
            core.require(
                rate is None or rate > 0,
                f"{scope_field}.actual_rate must be positive when known",
            )
            core.require(
                not operator or operator in RATE_OPERATORS,
                f"{scope_field}.rate_operator must be multiply or divide when known",
            )
            if not pending:
                core.require(
                    source_amount is not None,
                    f"{scope_field}.payment_total is required for a confirmed daily summary",
                )
                core.require(
                    target_amount is not None,
                    f"{scope_field}.actual_payout_total is required for a confirmed daily summary",
                )
                core.require(
                    rate is not None and operator in RATE_OPERATORS,
                    f"{scope_field}.rate and operator are required for a confirmed daily summary",
                )
            rate_text = core.decimal_text(rate) if rate is not None else ""
            rate_display_text = (
                rate_display(rate_text, operator)
                if rate is not None and operator in RATE_OPERATORS
                else "待确认"
            )
            operator_label = (
                "乘" if operator == "multiply" else "除" if operator == "divide" else ""
            )
            key = (fund_type, direction, operator, rate_text, status_label)
            if key not in buckets:
                first_seen[key] = sequence
                sequence += 1
                buckets[key] = {
                    "fund_type": fund_type,
                    "fund_type_label": FUND_TYPE_LABELS[fund_type],
                    "direction": direction,
                    "operator": operator,
                    "operator_label": operator_label,
                    "rate": rate_text,
                    "rate_display": rate_display_text,
                    "status_label": status_label,
                    "count": 0,
                    "source_total": Decimal("0"),
                    "target_total": Decimal("0"),
                    "source_total_complete": True,
                    "target_total_complete": True,
                }
            bucket = buckets[key]
            bucket["count"] += 1
            if source_amount is None:
                bucket["source_total_complete"] = False
            else:
                bucket["source_total"] += source_amount
            if target_amount is None:
                bucket["target_total_complete"] = False
            else:
                bucket["target_total"] += target_amount

    fund_order = {value: position for position, value in enumerate(FUND_TYPE_ORDER)}
    keys = sorted(
        buckets,
        key=lambda key: (fund_order[key[0]], first_seen[key]),
    )
    result = [buckets[key] for key in keys]
    for summary in result:
        summary["source_total"] = (
            core.decimal_text(summary["source_total"])
            if summary.pop("source_total_complete")
            else None
        )
        summary["target_total"] = (
            core.decimal_text(summary["target_total"])
            if summary.pop("target_total_complete")
            else None
        )
    return result


def compile_order_daily_ledger(
    orders_document: Mapping[str, Any],
    *,
    accounting_date: str,
) -> tuple[dict[str, Any], dict[str, int]]:
    """Wrap the shared small-format order ledger with large-group daily summaries."""

    raw_groups = orders_document.get("groups")
    core.require(isinstance(raw_groups, list), "large daily order groups must be a list")
    result_groups: list[dict[str, Any]] = []
    order_count = 0
    summary_count = 0
    for raw_group in raw_groups:
        core.require(isinstance(raw_group, Mapping), "large daily group must be an object")
        raw_orders = raw_group.get("orders")
        core.require(isinstance(raw_orders, list), "large daily group orders must be a list")
        if not raw_orders:
            continue
        typed_orders = [item for item in raw_orders if isinstance(item, Mapping)]
        core.require(
            len(typed_orders) == len(raw_orders),
            "large daily orders must contain only objects",
        )
        group = copy.deepcopy(dict(raw_group))
        group["daily_summaries"] = compile_order_summaries(
            typed_orders,
            group_key=core.clean_text(group.get("group_key")),
        )
        result_groups.append(group)
        order_count += len(typed_orders)
        summary_count += len(group["daily_summaries"])

    core.require(
        bool(result_groups),
        "no large-group orders were recorded for this day",
    )
    document = copy.deepcopy(dict(orders_document))
    document.update(
        {
            "contract_version": OUTPUT_CONTRACT,
            "accounting_mode": "large_daily",
            "accounting_date": accounting_date,
            "groups": result_groups,
        }
    )
    return document, {
        "groups": len(result_groups),
        "orders": order_count,
        "summary_rows": summary_count,
    }
