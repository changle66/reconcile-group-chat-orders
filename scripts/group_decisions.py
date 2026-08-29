#!/usr/bin/env python3
"""Read complete group chats and compile one semantic decision document per group."""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import tempfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Mapping

import core


GROUP_PAGE_CONTRACT = "group-chat-page/1.0"
GROUP_DECISION_CONTRACT = "group-chat-decision/1.0"
PLAN_CONTRACT = "small-group-simple-plan/1.1"
DEFAULT_PAGE_SIZE = 100
MAX_PAGE_SIZE = 500
DISPOSITIONS = frozenset({"order_evidence", "reference", "uncertain"})
FLOW_SIDES = frozenset({"payment", "payment_refund", "payout", "recovery"})
ORDER_FIELDS = frozenset(
    {
        "case_id",
        "event_ids",
        "source_message_ids",
        "start_message_id",
        "customer_id",
        "customer_nickname",
        "direction",
        "rate",
        "rate_operator",
        "expected_payout",
        "note",
        "same_transactions",
        "legs",
        "fees",
        "rounding",
    }
)
DECISION_FIELDS = frozenset(
    {"disposition", "event_type", "flow_side", "ocr", "note"}
)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)

    prepare = commands.add_parser("prepare", help="create one decision template per group")
    prepare.add_argument("normalized", type=Path)
    prepare.add_argument("--decisions", required=True, type=Path)
    prepare.add_argument("--force", action="store_true")

    read = commands.add_parser("read", help="read one chronological page from a group")
    read.add_argument("normalized", type=Path)
    read.add_argument("--group-key", required=True)
    read.add_argument("--cursor")
    read.add_argument("--limit", type=int, default=DEFAULT_PAGE_SIZE)

    compile_command = commands.add_parser(
        "compile", help="validate group decisions and compile events plus a complete plan"
    )
    compile_command.add_argument("normalized", type=Path)
    compile_command.add_argument("--decisions", required=True, type=Path)
    compile_command.add_argument("--events", required=True, type=Path)
    compile_command.add_argument("--plan", required=True, type=Path)
    compile_command.add_argument("--force", action="store_true")
    return parser.parse_args(argv)


def group_fingerprint(group: Mapping[str, Any]) -> str:
    """Fingerprint all normalized content for exactly one group."""
    return core.fingerprint_json(group)


def _group_index(normalized: Mapping[str, Any]) -> dict[str, Mapping[str, Any]]:
    groups = normalized.get("groups")
    core.require(isinstance(groups, list), "normalized groups must be a list")
    index: dict[str, Mapping[str, Any]] = {}
    for position, group in enumerate(groups):
        core.require(isinstance(group, Mapping), f"groups[{position}] must be an object")
        group_key = str(group.get("group_key") or "")
        core.require(group_key, f"groups[{position}].group_key is required")
        core.require(group_key not in index, f"duplicate group_key: {group_key}")
        index[group_key] = group
    return index


def _cursor_for(group_key: str, fingerprint: str, offset: int) -> str:
    token = core.stable_token(
        "group-page-cursor", group_key, fingerprint, offset, length=16
    )
    return f"v1.{offset}.{token}"


def _cursor_offset(
    cursor: str | None,
    *,
    group_key: str,
    fingerprint: str,
    message_count: int,
) -> int:
    if cursor in (None, ""):
        return 0
    matched = re.fullmatch(r"v1\.(\d+)\.([0-9a-f]{16})", str(cursor))
    core.require(matched is not None, "invalid group cursor")
    offset = int(matched.group(1))
    core.require(offset <= message_count, "group cursor is past the end of the group")
    core.require(
        matched.group(2) == _cursor_for(group_key, fingerprint, offset).rsplit(".", 1)[1],
        "group cursor is stale or belongs to another group",
    )
    return offset


def _message_snapshot(message: Mapping[str, Any]) -> dict[str, Any]:
    """Return the complete normalized message without recursively expanding replies."""
    return copy.deepcopy(dict(message))


def read_group(
    normalized: Mapping[str, Any],
    group_key: str,
    *,
    cursor: str | None = None,
    limit: int = DEFAULT_PAGE_SIZE,
) -> dict[str, Any]:
    """Read a chronological page while resolving direct replies against the whole group."""
    core.require(isinstance(limit, int) and 1 <= limit <= MAX_PAGE_SIZE, f"limit must be 1..{MAX_PAGE_SIZE}")
    groups = _group_index(normalized)
    core.require(group_key in groups, f"unknown group_key: {group_key}")
    group = groups[group_key]
    messages_value = group.get("messages")
    core.require(isinstance(messages_value, list), f"{group_key}: messages must be a list")
    messages = [message for message in messages_value if isinstance(message, Mapping)]
    core.require(len(messages) == len(messages_value), f"{group_key}: every message must be an object")
    fingerprint = group_fingerprint(group)
    offset = _cursor_offset(
        cursor,
        group_key=group_key,
        fingerprint=fingerprint,
        message_count=len(messages),
    )
    end = min(offset + limit, len(messages))
    message_by_id = {
        str(message.get("message_id") or ""): message for message in messages
    }
    page_messages: list[dict[str, Any]] = []
    for message in messages[offset:end]:
        snapshot = _message_snapshot(message)
        reply_id = str(message.get("reply_to_message_id") or "")
        snapshot["reply_message"] = (
            _message_snapshot(message_by_id[reply_id])
            if reply_id in message_by_id
            else None
        )
        page_messages.append(snapshot)
    done = end == len(messages)
    return {
        "contract_version": GROUP_PAGE_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "group_fingerprint": fingerprint,
        "group_key": group_key,
        "group_name": group.get("group_name"),
        "platform": group.get("platform"),
        "message_count": len(messages),
        "page_start": offset,
        "page_end": end,
        "messages": page_messages,
        "next_cursor": None if done else _cursor_for(group_key, fingerprint, end),
        "done": done,
    }


def _event_id(media_id: str) -> str:
    return f"{media_id}#event"


def _decision_template() -> dict[str, Any]:
    return {
        "disposition": None,
        "event_type": None,
        "flow_side": None,
        "ocr": None,
        "note": None,
    }


def fund_evidence_decision(
    *,
    event_type: str,
    payee: object,
    ocr: Mapping[str, Any],
    disposition: str = "order_evidence",
    flow_side: str | None = None,
    note: object = None,
) -> dict[str, Any]:
    """Build a fund decision while keeping the observed payee explicit and unchanged."""
    core.require(
        disposition in {"order_evidence", "uncertain"},
        "fund evidence disposition must be order_evidence or uncertain",
    )
    core.require(
        event_type in core.FUND_EVENT_TYPES,
        f"unsupported fund event_type {event_type!r}",
    )
    core.require(isinstance(ocr, Mapping), "fund evidence ocr must be an object")

    payee_text = core.validate_payee(
        payee,
        field="payee",
        cash=event_type in {"cash_payment", "cash_payout"},
    )
    normalized_ocr = copy.deepcopy(dict(ocr))
    existing_payee = core.clean_text(normalized_ocr.get("payee"))
    core.require(
        not existing_payee or existing_payee == payee_text,
        "ocr.payee conflicts with the explicit payee",
    )
    normalized_ocr["payee"] = payee_text

    normalized_side = core.clean_text(flow_side) or None
    if disposition == "uncertain":
        core.require(normalized_side is None, "uncertain fund evidence cannot contain flow_side")
    elif normalized_side is not None:
        core.require(normalized_side in FLOW_SIDES, f"unsupported flow_side {flow_side!r}")

    return {
        "disposition": disposition,
        "event_type": event_type,
        "flow_side": normalized_side,
        "ocr": normalized_ocr,
        "note": copy.deepcopy(note),
    }


def _evidence_media(group: Mapping[str, Any]) -> list[tuple[Mapping[str, Any], Mapping[str, Any]]]:
    result: list[tuple[Mapping[str, Any], Mapping[str, Any]]] = []
    for message in group.get("messages", []):
        if not isinstance(message, Mapping):
            continue
        for media in message.get("media", []):
            if isinstance(media, Mapping) and core.media_is_evidence(media):
                result.append((message, media))
    return result


def decision_template(
    normalized: Mapping[str, Any], group: Mapping[str, Any]
) -> dict[str, Any]:
    evidence = _evidence_media(group)
    return {
        "contract_version": GROUP_DECISION_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "group_fingerprint": group_fingerprint(group),
        "group_key": group.get("group_key"),
        "group_name": group.get("group_name"),
        "platform": group.get("platform"),
        "message_count": len(group.get("messages", [])),
        "evidence_media_count": len(evidence),
        "media_decisions": [
            {
                "media_id": media.get("media_id"),
                "event_id": _event_id(str(media.get("media_id") or "")),
                "message_id": message.get("message_id"),
                "availability": media.get("availability"),
                "decision": _decision_template(),
            }
            for message, media in evidence
        ],
        "orders": [],
        "balance_links": [],
    }


def _decision_filename(group_key: str) -> str:
    return f"decision_{group_key.replace(':', '-')}.json"


def prepare_decision_files(
    normalized: Mapping[str, Any],
    decisions_dir: Path,
    *,
    force: bool = False,
) -> dict[str, Any]:
    decisions_dir.mkdir(parents=True, exist_ok=True)
    reports: list[dict[str, Any]] = []
    for group in normalized.get("groups", []):
        group_key = str(group.get("group_key") or "")
        path = decisions_dir / _decision_filename(group_key)
        created = not path.exists()
        written = created or force
        template = decision_template(normalized, group)
        if written:
            core.atomic_json(path, template)
        reports.append(
            {
                "group_key": group_key,
                "group_name": group.get("group_name"),
                "messages": template["message_count"],
                "evidence_media": template["evidence_media_count"],
                "decision_file": str(path.resolve()),
                "created": created,
                "written": written,
            }
        )
    return {
        "decision_contract": GROUP_DECISION_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "groups": reports,
        "total_groups": len(reports),
        "total_evidence_media": sum(item["evidence_media"] for item in reports),
    }


def _read_decision_files(decisions_dir: Path) -> list[tuple[Path, dict[str, Any]]]:
    paths = sorted(decisions_dir.glob("*.json"), key=lambda item: item.name.casefold())
    core.require(bool(paths), f"no group decision JSON files found in {decisions_dir}")
    result: list[tuple[Path, dict[str, Any]]] = []
    for path in paths:
        try:
            document = json.loads(path.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError as exc:
            raise ValueError(f"{path}: invalid JSON: {exc.msg}") from exc
        core.require(isinstance(document, dict), f"{path}: group decision must be an object")
        result.append((path, document))
    return result


def _present(value: object) -> bool:
    return value not in (None, "", [], {})


def _validate_media_decisions(
    document: Mapping[str, Any],
    expected: Mapping[str, Any],
) -> tuple[dict[str, dict[str, Any]], dict[str, int]]:
    group_key = str(expected.get("group_key") or "")
    values = document.get("media_decisions")
    core.require(isinstance(values, list), f"{group_key}: media_decisions must be a list")
    expected_values = expected["media_decisions"]
    expected_by_media = {str(item["media_id"]): item for item in expected_values}
    actual_by_media: dict[str, dict[str, Any]] = {}
    counts: defaultdict[str, int] = defaultdict(int)
    for position, item in enumerate(values):
        where = f"{group_key}: media_decisions[{position}]"
        core.require(isinstance(item, dict), f"{where} must be an object")
        media_id = str(item.get("media_id") or "")
        core.require(media_id in expected_by_media, f"{where}: unknown media_id {media_id!r}")
        core.require(media_id not in actual_by_media, f"{where}: duplicate media_id {media_id}")
        source = expected_by_media[media_id]
        for field in ("event_id", "message_id", "availability"):
            core.require(item.get(field) == source.get(field), f"{where}: generated {field} was edited")
        decision = item.get("decision")
        core.require(isinstance(decision, dict), f"{where}: decision must be an object")
        unknown_fields = sorted(set(decision) - DECISION_FIELDS)
        core.require(not unknown_fields, f"{where}: unsupported decision fields: {', '.join(unknown_fields)}")
        if item.get("availability") == "missing":
            core.require(
                not any(_present(decision.get(field)) for field in DECISION_FIELDS),
                f"{where}: missing media cannot contain a decision",
            )
            counts["missing"] += 1
            actual_by_media[media_id] = item
            continue

        disposition = decision.get("disposition")
        core.require(
            disposition in DISPOSITIONS,
            f"{where}: disposition is required and must be order_evidence, reference, or uncertain",
        )
        event_type = decision.get("event_type")
        flow_side = decision.get("flow_side")
        ocr = decision.get("ocr")
        if disposition == "reference":
            core.require(
                not any(_present(value) for value in (event_type, flow_side, ocr)),
                f"{where}: reference must not contain fund event_type, flow_side, or ocr",
            )
        else:
            core.require(
                event_type in core.FUND_EVENT_TYPES,
                f"{where}: {disposition} requires a supported fund event_type",
            )
            core.require(isinstance(ocr, dict), f"{where}: {disposition} requires ocr")
            if disposition == "uncertain":
                core.require(not _present(flow_side), f"{where}: uncertain must not contain flow_side")
            elif _present(flow_side):
                core.require(flow_side in FLOW_SIDES, f"{where}: unsupported flow_side {flow_side!r}")
        counts[str(disposition)] += 1
        actual_by_media[media_id] = item

    missing = sorted(set(expected_by_media) - set(actual_by_media))
    core.require(not missing, f"{group_key}: missing media decisions: {', '.join(missing[:20])}")
    core.require(
        len(actual_by_media) == len(expected_by_media),
        f"{group_key}: media decision coverage does not match normalized evidence media",
    )
    return actual_by_media, dict(counts)


def validate_group_decisions(
    normalized: Mapping[str, Any], decisions_dir: Path
) -> list[tuple[Mapping[str, Any], dict[str, Any], dict[str, dict[str, Any]], dict[str, int]]]:
    groups = _group_index(normalized)
    documents: dict[str, tuple[Path, dict[str, Any]]] = {}
    for path, document in _read_decision_files(decisions_dir):
        group_key = str(document.get("group_key") or "")
        core.require(group_key in groups, f"{path}: unknown group_key {group_key!r}")
        core.require(group_key not in documents, f"{path}: duplicate decision for {group_key}")
        documents[group_key] = (path, document)
    missing_groups = sorted(set(groups) - set(documents))
    core.require(not missing_groups, "missing group decisions: " + ", ".join(missing_groups))
    core.require(len(documents) == len(groups), "group decisions must cover every selected group exactly once")

    validated: list[tuple[Mapping[str, Any], dict[str, Any], dict[str, dict[str, Any]], dict[str, int]]] = []
    for group_key, group in groups.items():
        path, document = documents[group_key]
        expected = decision_template(normalized, group)
        core.require(
            document.get("contract_version") == GROUP_DECISION_CONTRACT,
            f"{path}: unsupported contract_version",
        )
        core.require(
            document.get("normalized_source_fingerprint") == normalized.get("source_fingerprint"),
            f"{path}: normalized_source_fingerprint is stale",
        )
        core.require(
            document.get("group_fingerprint") == expected["group_fingerprint"],
            f"{path}: group_fingerprint is stale",
        )
        for field in ("group_name", "platform", "message_count", "evidence_media_count"):
            core.require(document.get(field) == expected.get(field), f"{path}: generated {field} was edited")
        core.require(isinstance(document.get("orders"), list), f"{path}: orders must be a list")
        core.require(isinstance(document.get("balance_links"), list), f"{path}: balance_links must be a list")
        media, counts = _validate_media_decisions(document, expected)
        validated.append((group, document, media, counts))
    return validated


def _source_ids(
    value: object,
    *,
    message_ids: set[str],
    field: str,
    required: bool,
) -> list[str]:
    if value is None and not required:
        return []
    core.require(isinstance(value, list), f"{field} must be a list")
    ids = [str(item) for item in value]
    core.require(bool(ids) or not required, f"{field} must cite at least one source message")
    core.require(len(ids) == len(set(ids)), f"{field} repeats a message ID")
    unknown = [message_id for message_id in ids if message_id not in message_ids]
    core.require(
        not unknown,
        f"{field} cites a missing or cross-group message: {unknown[0] if unknown else ''}",
    )
    return ids


def _compile_orders(
    group: Mapping[str, Any],
    document: Mapping[str, Any],
    event_by_id: Mapping[str, dict[str, Any]],
    order_event_ids: set[str],
) -> tuple[list[dict[str, Any]], set[str]]:
    group_key = str(group.get("group_key") or "")
    messages = {
        str(message.get("message_id") or ""): message
        for message in group.get("messages", [])
        if isinstance(message, Mapping)
    }
    sequences = {
        message_id: int(message.get("source_sequence") or position + 1)
        for position, (message_id, message) in enumerate(messages.items())
    }
    assigned: set[str] = set()
    case_ids: set[str] = set()
    compiled: list[dict[str, Any]] = []
    for position, raw_order in enumerate(document.get("orders", []), start=1):
        field = f"{group_key}: orders[{position - 1}]"
        core.require(isinstance(raw_order, Mapping), f"{field} must be an object")
        unknown_fields = sorted(set(raw_order) - ORDER_FIELDS)
        core.require(not unknown_fields, f"{field}: unsupported fields: {', '.join(unknown_fields)}")
        if raw_order.get("legs") not in (None, "", []):
            core.require(
                bool(core.clean_text(raw_order.get("note"))),
                f"{field}.note is required for a multi-leg order because legs are shown in the order remark",
            )
        event_ids_value = raw_order.get("event_ids")
        core.require(isinstance(event_ids_value, list) and event_ids_value, f"{field}.event_ids is required")
        event_ids = [str(item) for item in event_ids_value]
        core.require(len(event_ids) == len(set(event_ids)), f"{field}.event_ids repeats an event")
        for event_id in event_ids:
            core.require(event_id in order_event_ids, f"{field}: event is not order_evidence: {event_id}")
            core.require(event_id in event_by_id, f"{field}: unknown event_id: {event_id}")
            core.require(event_id not in assigned, f"fund event assigned to multiple orders: {event_id}")

        source_message_ids = _source_ids(
            raw_order.get("source_message_ids"),
            message_ids=set(messages),
            field=f"{field}.source_message_ids",
            required=True,
        )
        explicit_start = core.clean_text(raw_order.get("start_message_id"))
        if explicit_start:
            core.require(explicit_start in messages, f"{field}: start_message_id is missing or cross-group")
            start_message_id = explicit_start
        else:
            event_message_ids = [str(event_by_id[event_id]["message_id"]) for event_id in event_ids]
            candidates = list(dict.fromkeys([*source_message_ids, *event_message_ids]))
            start_message_id = min(candidates, key=lambda message_id: (sequences[message_id], message_id))

        case_id = core.clean_text(raw_order.get("case_id")) or (
            "case-" + core.stable_token(group_key, *event_ids, *source_message_ids)
        )
        core.require(case_id not in case_ids, f"{field}: duplicate case_id {case_id}")
        case_ids.add(case_id)

        order = {
            key: copy.deepcopy(value)
            for key, value in raw_order.items()
            if key not in {"case_id", "event_ids", "source_message_ids", "start_message_id"}
        }
        order.update(
            {
                "case_id": case_id,
                "event_ids": event_ids,
                "source_message_ids": source_message_ids,
                "start_message_id": start_message_id,
            }
        )
        event_sides = {
            event_id: event_by_id[event_id]["flow_side"]
            for event_id in event_ids
            if event_by_id[event_id].get("flow_side")
        }
        if event_sides:
            order["event_sides"] = event_sides
        compiled.append(order)
        assigned.update(event_ids)

    missing = sorted(order_event_ids - assigned)
    core.require(
        not missing,
        f"{group_key}: order_evidence events not assigned to an order: {', '.join(missing[:20])}",
    )
    return compiled, assigned


def compile_decision_documents(
    normalized: Mapping[str, Any],
    validated: list[tuple[Mapping[str, Any], dict[str, Any], dict[str, dict[str, Any]], dict[str, int]]],
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, Any], dict[str, int]]:
    events_by_group: dict[str, list[dict[str, Any]]] = {}
    plan_groups: list[dict[str, Any]] = []
    total_counts: defaultdict[str, int] = defaultdict(int)
    for group, document, media_by_id, counts in validated:
        group_key = str(group.get("group_key") or "")
        group_events: list[dict[str, Any]] = []
        event_by_id: dict[str, dict[str, Any]] = {}
        order_event_ids: set[str] = set()
        for message, media in _evidence_media(group):
            media_id = str(media.get("media_id") or "")
            item = media_by_id[media_id]
            decision = item["decision"]
            availability = media.get("availability")
            if availability == "missing":
                disposition = "missing"
                event_type = "missing_evidence_media"
            else:
                disposition = str(decision["disposition"])
                event_type = "irrelevant_media" if disposition == "reference" else str(decision["event_type"])
            event: dict[str, Any] = {
                "event_id": item["event_id"],
                "group_key": group_key,
                "message_id": message.get("message_id"),
                "media_id": media_id,
                "type": event_type,
                "decision_disposition": disposition,
                "source_group_fingerprint": document["group_fingerprint"],
            }
            if disposition in {"order_evidence", "uncertain"}:
                event["ocr"] = copy.deepcopy(decision["ocr"])
                media_path = core.clean_text(media.get("path"))
                existing_hash = core.clean_text(media.get("blob_sha256"))
                if existing_hash:
                    event["evidence_sha256"] = existing_hash
                elif availability == "available" and media_path and Path(media_path).is_file():
                    event["evidence_sha256"] = core.sha256_file(Path(media_path))
            if _present(decision.get("note")):
                event["note"] = decision["note"]
            if disposition == "order_evidence":
                order_event_ids.add(str(event["event_id"]))
                if _present(decision.get("flow_side")):
                    event["flow_side"] = str(decision["flow_side"])
            group_events.append(event)
            event_by_id[str(event["event_id"])] = event

        orders, _ = _compile_orders(group, document, event_by_id, order_event_ids)
        message_ids = set(
            str(message.get("message_id") or "")
            for message in group.get("messages", [])
            if isinstance(message, Mapping)
        )
        balance_links = copy.deepcopy(document.get("balance_links", []))
        for position, link in enumerate(balance_links):
            core.require(isinstance(link, Mapping), f"{group_key}: balance_links[{position}] must be an object")
            if "source_message_ids" in link:
                _source_ids(
                    link.get("source_message_ids"),
                    message_ids=message_ids,
                    field=f"{group_key}: balance_links[{position}].source_message_ids",
                    required=False,
                )
        events_by_group[group_key] = group_events
        plan_groups.append(
            {
                "group_key": group_key,
                "orders": orders,
                "balance_links": balance_links,
            }
        )
        for key, value in counts.items():
            total_counts[key] += value

    plan = {
        "contract_version": PLAN_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "groups": plan_groups,
    }
    return events_by_group, plan, dict(total_counts)


def _event_filename(group_key: str) -> str:
    return f"events_{group_key.replace(':', '-')}.jsonl"


def _event_jsonl(events: list[Mapping[str, Any]]) -> str:
    return "".join(json.dumps(event, ensure_ascii=False) + "\n" for event in events)


def _atomic_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.parent / f".{path.name}.tmp"
    temporary.write_text(content, encoding="utf-8")
    temporary.replace(path)


def _validate_event_payloads(
    normalized: Mapping[str, Any], payloads: Mapping[str, str]
) -> str:
    with tempfile.TemporaryDirectory(prefix="group-decision-events-") as temporary:
        root = Path(temporary)
        for filename, content in payloads.items():
            (root / filename).write_text(content, encoding="utf-8")
        _, fingerprint = core.load_events(root, normalized)
        return fingerprint


def compile_group_decisions(
    normalized: Mapping[str, Any],
    decisions_dir: Path,
    events_dir: Path,
    plan_path: Path,
    *,
    force: bool = False,
) -> dict[str, Any]:
    validated = validate_group_decisions(normalized, decisions_dir)
    events_by_group, plan, counts = compile_decision_documents(normalized, validated)
    payloads = {
        _event_filename(group_key): _event_jsonl(events)
        for group_key, events in events_by_group.items()
    }
    events_fingerprint = _validate_event_payloads(normalized, payloads)
    plan["events_fingerprint"] = events_fingerprint

    expected_paths = {events_dir / filename for filename in payloads}
    existing_paths = set(events_dir.glob("*.jsonl")) if events_dir.exists() else set()
    extras = sorted(existing_paths - expected_paths, key=lambda item: item.name.casefold())
    core.require(
        not extras,
        "events directory contains unexpected JSONL files: " + ", ".join(path.name for path in extras),
    )
    if not force:
        occupied = sorted(existing_paths & expected_paths, key=lambda item: item.name.casefold())
        core.require(not occupied, f"event output already exists (use --force): {occupied[0] if occupied else ''}")
        core.require(not plan_path.exists(), f"plan output already exists (use --force): {plan_path}")

    events_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in payloads.items():
        _atomic_text(events_dir / filename, content)
    _, written_fingerprint = core.load_events(events_dir, normalized)
    core.require(written_fingerprint == events_fingerprint, "written events fingerprint changed unexpectedly")
    core.atomic_json(plan_path, plan)
    return {
        "decision_contract": GROUP_DECISION_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "events_fingerprint": events_fingerprint,
        "groups": len(validated),
        "events": sum(len(events) for events in events_by_group.values()),
        "orders": sum(len(group["orders"]) for group in plan["groups"]),
        "dispositions": counts,
        "events_dir": str(events_dir.resolve()),
        "plan": str(plan_path.resolve()),
    }


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        normalized = core.load_normalized(args.normalized.resolve())
        if args.command == "prepare":
            report = prepare_decision_files(
                normalized,
                args.decisions.resolve(),
                force=args.force,
            )
        elif args.command == "read":
            report = read_group(
                normalized,
                args.group_key,
                cursor=args.cursor,
                limit=args.limit,
            )
        else:
            report = compile_group_decisions(
                normalized,
                args.decisions.resolve(),
                args.events.resolve(),
                args.plan.resolve(),
                force=args.force,
            )
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(f"Group decision workflow failed: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
