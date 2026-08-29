#!/usr/bin/env python3
"""Run a fresh group-chat reconciliation through start, review, and finish."""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Iterable, Mapping
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import build_workbook
import check_workbook
import core
import extract_line_ios
import group_decisions
import normalize_exports
import simple_ledger


RUN_CONTRACT = "group-chat-reconcile-run/1.0"
DECISION_CONTRACT = "group-chat-decision/2.2"
REVIEW_PAGE_CONTRACT = "group-chat-review-page/1.0"
NORMALIZER_VERSION = "reconcile-start/1.0"
DEFAULT_PAGE_SIZE = 200
MAX_PAGE_SIZE = 500
FLOW_SIDES = frozenset({"payment", "payment_refund", "payout", "recovery", "unknown"})
ENTRY_KINDS = frozenset({"transfer", "cash"})
ENTRY_RESULTS = frozenset({"completed", "failed", "pending", "not_shown", "unknown"})
AMOUNT_STATES = frozenset({"clear", "partial", "unreadable"})
RATE_STATES = frozenset({"adopted", "not_stated", "uncertain"})
EXPECTED_PAYOUT_STATES = frozenset(
    {"explicit", "calculated_from_rate", "not_stated", "uncertain"}
)
STATUS_CLASS = {
    "completed": "completed",
    "failed": "failed",
    "pending": "pending",
    "not_shown": "blank",
    "unknown": "unknown",
}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)

    start = commands.add_parser("start", help="create a new run from raw exports")
    start.add_argument("inputs", nargs="+", type=Path, help="raw export files or roots")
    start.add_argument("--work", required=True, type=Path, help="new, non-existing run directory")
    start.add_argument("--contains", default="小额", help="required text in the parsed group name")
    start.add_argument("--timezone", default="Asia/Bangkok")
    start.add_argument(
        "--roster",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "config" / "roster.yaml",
    )
    start.add_argument("--line-backup", action="append", default=[], type=Path)
    start.add_argument("--line-self-name", default="LINE_SELF")

    review = commands.add_parser("review", help="read, validate, or seal one group review")
    review.add_argument("work", type=Path)
    review_actions = review.add_subparsers(dest="review_action", required=True)
    next_page = review_actions.add_parser("next", help="return the next compact chronological page")
    next_page.add_argument("--group", help="group key, run label, or unique exact group name")
    next_page.add_argument("--limit", type=int, default=DEFAULT_PAGE_SIZE)
    status = review_actions.add_parser("status", help="show review progress")
    status.add_argument("--group", help="optional group selector")
    check = review_actions.add_parser("check", help="validate the current decision data")
    check.add_argument("--group", required=True)
    seal = review_actions.add_parser("seal", help="validate and seal a completely reviewed group")
    seal.add_argument("--group", required=True)

    finish = commands.add_parser("finish", help="compile, verify, and publish the final workbook")
    finish.add_argument("work", type=Path)
    finish.add_argument("-o", "--output", required=True, type=Path)
    finish.add_argument(
        "--template",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "assets" / "模版.xlsx",
    )
    return parser.parse_args(argv)


def _run_path(work: Path) -> Path:
    return work / "run.json"


def _snapshot_path(work: Path) -> Path:
    return work / "snapshot" / "normalized.json"


def _load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    core.require(isinstance(value, dict), f"{path}: top level must be an object")
    return value


def _load_run(work: Path) -> dict[str, Any]:
    path = _run_path(work)
    core.require(path.is_file(), f"run manifest does not exist: {path}")
    run = _load_json(path)
    core.require(run.get("contract_version") == RUN_CONTRACT, "unsupported run contract")
    core.require(isinstance(run.get("groups"), list), "run groups must be a list")
    return run


def _load_snapshot(work: Path, run: Mapping[str, Any]) -> dict[str, Any]:
    path = _snapshot_path(work)
    core.require(path.is_file(), f"normalized snapshot does not exist: {path}")
    expected = core.clean_text(run.get("snapshot_sha256"))
    actual = core.sha256_file(path)
    core.require(actual == expected, "normalized snapshot changed after start; create a new run")
    normalized = core.load_normalized(path)
    core.require(
        normalized.get("source_fingerprint") == run.get("normalized_source_fingerprint"),
        "run manifest and normalized snapshot disagree",
    )
    return normalized


def _source_document(
    inputs: Iterable[Path],
    *,
    zone: ZoneInfo,
    roster: Mapping[str, Any],
) -> dict[str, Any] | None:
    try:
        telegram_sources, whatsapp_sources = normalize_exports.discover_sources(inputs)
    except ValueError as exc:
        if str(exc) == "no Telegram result.json or WhatsApp .txt exports found":
            return None
        raise
    groups: list[dict[str, Any]] = []
    warnings: list[str] = []
    fingerprint_paths: list[Path] = []
    for source in telegram_sources:
        group, local_warnings, local_paths = normalize_exports.normalize_telegram(
            source,
            zone=zone,
            roster=dict(roster),
        )
        groups.append(group)
        warnings.extend(local_warnings)
        fingerprint_paths.extend(local_paths)
    for source in whatsapp_sources:
        group, local_warnings, local_paths = normalize_exports.normalize_whatsapp(
            source,
            zone=zone,
            roster=dict(roster),
        )
        groups.append(group)
        warnings.extend(local_warnings)
        fingerprint_paths.extend(local_paths)
    group_keys = [str(group["group_key"]) for group in groups]
    core.require(len(group_keys) == len(set(group_keys)), "duplicate group export; provide one snapshot per group")
    groups.sort(key=lambda item: (str(item.get("platform")), str(item.get("group_name")), str(item.get("group_key"))))
    source_fingerprint = core.fingerprint_files(
        fingerprint_paths,
        context={
            "normalizer": NORMALIZER_VERSION,
            "timezone": str(zone.key),
            "roster": roster,
            "groups": group_keys,
        },
    )
    messages = [message for group in groups for message in group.get("messages", [])]
    return {
        "contract_version": core.NORMALIZED_CONTRACT,
        "timezone": str(zone.key),
        "source_fingerprint": source_fingerprint,
        "groups": groups,
        "statistics": {
            "groups": len(groups),
            "messages": len(messages),
            "excluded_messages": sum(bool(message.get("excluded_from_accounting")) for message in messages),
            "media": sum(len(message.get("media", [])) for message in messages),
            "source_files": len(telegram_sources) + len(whatsapp_sources),
            "fingerprinted_files": len({path.resolve() for path in fingerprint_paths}),
        },
        "warnings": warnings,
        "skipped_files": [],
    }


def _generated_path(path: Path) -> bool:
    return any(
        marker.casefold() in part.casefold()
        for part in path.parts
        for marker in normalize_exports.GENERATED_PATH_MARKERS
    )


def _discover_line_backups(inputs: Iterable[Path], explicit: Iterable[Path]) -> list[Path]:
    candidates: set[Path] = set()
    for raw in explicit:
        candidates.add(extract_line_ios.find_device_dir(raw).resolve())
    for raw in inputs:
        path = raw.resolve()
        if path.is_file() and path.name.casefold() == "manifest.db":
            candidates.add(path.parent)
        elif path.is_dir():
            direct = path / "Manifest.db"
            if direct.is_file() and not _generated_path(direct):
                candidates.add(path)
            for manifest in path.rglob("Manifest.db"):
                if not _generated_path(manifest):
                    candidates.add(manifest.parent.resolve())
    return sorted(candidates, key=lambda item: str(item).casefold())


def _line_has_matching_group(device_dir: Path, contains: str) -> bool:
    manifest = extract_line_ios.sqlite_ro(device_dir / "Manifest.db")
    line_db = None
    group_db = None
    try:
        line_row = extract_line_ios.manifest_row_for_suffix(manifest, "/Messages/Line.sqlite")
        group_row = extract_line_ios.manifest_row_for_suffix(manifest, "/Messages/UnifiedGroup.sqlite")
        line_path = extract_line_ios.physical_backup_file(device_dir, str(line_row["fileID"]))
        group_path = extract_line_ios.physical_backup_file(device_dir, str(group_row["fileID"]))
        line_db = extract_line_ios.sqlite_ro(line_path)
        group_db = extract_line_ios.sqlite_ro(group_path)
        groups = extract_line_ios.available_groups(line_db, group_db)
        needle = contains.casefold()
        return any(needle in str(group.get("group_name") or "").casefold() for group in groups)
    finally:
        if line_db is not None:
            line_db.close()
        if group_db is not None:
            group_db.close()
        manifest.close()


def _line_documents(
    work: Path,
    backups: Iterable[Path],
    *,
    contains: str,
    timezone_name: str,
    roster_path: Path,
    self_name: str,
) -> list[dict[str, Any]]:
    documents: list[dict[str, Any]] = []
    for index, backup in enumerate(backups, start=1):
        if not _line_has_matching_group(backup, contains):
            continue
        output = work / "snapshot" / "sources" / f"line_{index:02d}.json"
        media_dir = work / "snapshot" / "line_media" / f"backup_{index:02d}"
        exit_code = extract_line_ios.main(
            [
                str(backup),
                "-o",
                str(output),
                "--timezone",
                timezone_name,
                "--group-pattern",
                re.escape(contains),
                "--self-name",
                self_name,
                "--roster",
                str(roster_path),
                "--media-dir",
                str(media_dir),
            ]
        )
        core.require(exit_code == 0, f"LINE extraction failed for {backup}")
        documents.append(core.load_normalized(output))
    return documents


def _merge_documents(documents: list[dict[str, Any]], *, contains: str, timezone_name: str) -> dict[str, Any]:
    core.require(bool(documents), "no supported raw chat exports were found")
    groups: list[dict[str, Any]] = []
    warnings: list[str] = []
    source_fingerprints: list[str] = []
    seen: set[str] = set()
    needle = contains.casefold()
    for document in documents:
        core.require(document.get("timezone") == timezone_name, "normalized source timezone mismatch")
        source_fingerprints.append(str(document.get("source_fingerprint") or ""))
        warnings.extend(str(item) for item in document.get("warnings", []))
        for group in document.get("groups", []):
            if needle not in str(group.get("group_name") or "").casefold():
                continue
            key = str(group.get("group_key") or "")
            core.require(key and key not in seen, f"duplicate group snapshot: {key}")
            seen.add(key)
            groups.append(copy.deepcopy(group))
    core.require(bool(groups), f"no parsed group name contains {contains!r}")
    groups.sort(key=lambda item: (str(item.get("platform")), str(item.get("group_name")), str(item.get("group_key"))))
    messages = [message for group in groups for message in group.get("messages", [])]
    media = [item for message in messages for item in message.get("media", [])]
    source_fingerprint = core.fingerprint_json(
        {
            "normalizer": NORMALIZER_VERSION,
            "contains": contains,
            "sources": sorted(source_fingerprints),
            "groups": [group["group_key"] for group in groups],
        }
    )
    return {
        "contract_version": core.NORMALIZED_CONTRACT,
        "timezone": timezone_name,
        "source_fingerprint": source_fingerprint,
        "groups": groups,
        "statistics": {
            "groups": len(groups),
            "messages": len(messages),
            "excluded_messages": sum(bool(message.get("excluded_from_accounting")) for message in messages),
            "media": len(media),
            "available_media": sum(item.get("availability") == "available" for item in media),
            "missing_media": sum(item.get("availability") == "missing" for item in media),
            "normalized_inputs": len(documents),
        },
        "warnings": warnings,
        "skipped_files": [],
    }


def _decision_filename(group_key: str) -> str:
    return f"decision_{group_key.replace(':', '-')}.json"


def _decision_template(normalized: Mapping[str, Any], group: Mapping[str, Any]) -> dict[str, Any]:
    evidence = _media_inventory(group)
    return {
        "contract_version": DECISION_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "group_fingerprint": group_decisions.group_fingerprint(group),
        "group_key": group.get("group_key"),
        "group_name": group.get("group_name"),
        "platform": group.get("platform"),
        "message_count": len(group.get("messages", [])),
        "evidence_media_count": len(evidence),
        "reviewed_through": 0,
        "read_complete": False,
        "sealed": False,
        "sealed_decision_fingerprint": None,
        "media_decisions": {},
        "orders": [],
        "balance_links": [],
    }


def _platform_prefix(platform: object) -> str:
    return {"Telegram": "TG", "WhatsApp": "WA", "LINE": "LINE"}.get(str(platform), "CHAT")


def start_run(args: argparse.Namespace) -> dict[str, Any]:
    work = args.work.resolve()
    core.require(not work.exists(), f"work directory already exists; start requires a new path: {work}")
    inputs = [path.resolve() for path in args.inputs]
    for path in inputs:
        core.require(path.exists(), f"input does not exist: {path}")
    core.require(bool(core.clean_text(args.contains)), "--contains cannot be empty")
    zone = ZoneInfo(args.timezone)
    core.require(args.timezone == "Asia/Bangkok", "accounting timezone must be Asia/Bangkok")
    roster_path = args.roster.resolve()
    roster = core.load_roster(roster_path)
    work.mkdir(parents=True)
    (work / "snapshot" / "sources").mkdir(parents=True)
    (work / "decisions").mkdir()

    documents: list[dict[str, Any]] = []
    chat_document = _source_document(inputs, zone=zone, roster=roster)
    if chat_document is not None:
        documents.append(chat_document)
    line_backups = _discover_line_backups(inputs, args.line_backup)
    documents.extend(
        _line_documents(
            work,
            line_backups,
            contains=args.contains,
            timezone_name=args.timezone,
            roster_path=roster_path,
            self_name=args.line_self_name,
        )
    )
    normalized = _merge_documents(documents, contains=args.contains, timezone_name=args.timezone)
    snapshot_path = _snapshot_path(work)
    core.atomic_json(snapshot_path, normalized)
    core.load_normalized(snapshot_path)

    group_reports: list[dict[str, Any]] = []
    used_labels: set[str] = set()
    for group in normalized["groups"]:
        decision = _decision_template(normalized, group)
        filename = _decision_filename(str(group["group_key"]))
        decision_path = work / "decisions" / filename
        core.atomic_json(decision_path, decision)
        base_label = f"[{_platform_prefix(group.get('platform'))}] {group.get('group_name')}"
        run_label = base_label
        if run_label in used_labels:
            run_label = f"{base_label} ({core.stable_token(group.get('group_key'), length=6)})"
        used_labels.add(run_label)
        media_count = len(_media_inventory(group))
        group_reports.append(
            {
                "group_key": group["group_key"],
                "group_name": group.get("group_name"),
                "platform": group.get("platform"),
                "run_label": run_label,
                "messages": len(group.get("messages", [])),
                "evidence_media": media_count,
                "decision_file": f"decisions/{filename}",
            }
        )
    run = {
        "contract_version": RUN_CONTRACT,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "accounting_timezone": args.timezone,
        "group_name_contains": args.contains,
        "status": "reviewing",
        "normalized_source_fingerprint": normalized["source_fingerprint"],
        "snapshot_sha256": core.sha256_file(snapshot_path),
        "input_roots": [str(path) for path in inputs],
        "line_backups": [str(path) for path in line_backups],
        "groups": group_reports,
    }
    core.atomic_json(_run_path(work), run)
    return {
        "work": str(work),
        "groups": len(group_reports),
        "messages": normalized["statistics"]["messages"],
        "evidence_media": sum(item["evidence_media"] for item in group_reports),
        "selected_groups": group_reports,
    }


def _group_index(normalized: Mapping[str, Any]) -> dict[str, Mapping[str, Any]]:
    return {str(group["group_key"]): group for group in normalized.get("groups", [])}


def _message_labels(group: Mapping[str, Any]) -> tuple[dict[str, str], dict[str, Mapping[str, Any]]]:
    label_by_id: dict[str, str] = {}
    message_by_label: dict[str, Mapping[str, Any]] = {}
    for index, message in enumerate(group.get("messages", []), start=1):
        label = f"S{index:05d}"
        message_id = str(message.get("message_id") or "")
        label_by_id[message_id] = label
        message_by_label[label] = message
    return label_by_id, message_by_label


def _media_inventory(group: Mapping[str, Any]) -> dict[str, tuple[Mapping[str, Any], Mapping[str, Any]]]:
    inventory: dict[str, tuple[Mapping[str, Any], Mapping[str, Any]]] = {}
    index = 0
    for message in group.get("messages", []):
        for media in message.get("media", []):
            if isinstance(media, Mapping) and core.media_is_evidence(media):
                index += 1
                inventory[f"M{index:04d}"] = (message, media)
    return inventory


def _select_run_group(run: Mapping[str, Any], selector: str | None, *, prefer_incomplete: bool = False) -> Mapping[str, Any]:
    groups = [item for item in run.get("groups", []) if isinstance(item, Mapping)]
    if selector:
        exact = [
            item
            for item in groups
            if selector in {str(item.get("group_key")), str(item.get("run_label"))}
            or selector == str(item.get("group_name"))
        ]
        core.require(len(exact) == 1, f"group selector must match exactly one group: {selector!r}")
        return exact[0]
    core.require(prefer_incomplete, "--group is required")
    for item in groups:
        if not bool(item.get("sealed")):
            return item
    raise ValueError("all selected groups are already sealed")


def _decision_path(work: Path, run_group: Mapping[str, Any]) -> Path:
    relative = Path(str(run_group.get("decision_file") or ""))
    core.require(not relative.is_absolute() and ".." not in relative.parts, "invalid decision path in run manifest")
    return work / relative


def _semantic_fingerprint(decision: Mapping[str, Any]) -> str:
    return core.fingerprint_json(
        {
            "media_decisions": decision.get("media_decisions"),
            "orders": decision.get("orders"),
            "balance_links": decision.get("balance_links"),
        }
    )


def _validate_entry(entry: dict[str, Any], *, field: str) -> None:
    allowed = {
        "amount",
        "amount_text",
        "currency",
        "payee",
        "kind",
        "side",
        "result",
        "status_text",
        "amount_state",
        "note",
    }
    unknown = sorted(set(entry) - allowed)
    core.require(not unknown, f"{field}: unsupported fields: {', '.join(unknown)}")
    kind = core.clean_text(entry.get("kind")).casefold() or "transfer"
    core.require(kind in ENTRY_KINDS, f"{field}.kind must be transfer or cash")
    entry["kind"] = kind
    core.require("side" in entry, f"{field}.side is required")
    side = core.clean_text(entry.get("side")).casefold()
    core.require(side in FLOW_SIDES, f"{field}.side is unsupported")
    entry["side"] = side
    result = core.clean_text(entry.get("result")).casefold()
    core.require(result in ENTRY_RESULTS, f"{field}.result is required")
    entry["result"] = result
    amount_state = core.clean_text(entry.get("amount_state")).casefold()
    core.require(amount_state in AMOUNT_STATES, f"{field}.amount_state is required")
    entry["amount_state"] = amount_state
    if entry.get("amount") not in (None, ""):
        amount = core.parse_decimal(entry.get("amount"), field=f"{field}.amount")
        core.require(amount is not None and amount >= 0, f"{field}.amount cannot be negative")
        entry["amount"] = core.decimal_text(amount)
    if entry.get("currency") not in (None, ""):
        entry["currency"] = core.normalize_currency(entry.get("currency"), field=f"{field}.currency")
    if amount_state == "clear":
        core.require(entry.get("amount") not in (None, ""), f"{field}.amount is required when amount_state is clear")
        core.require(entry.get("currency") not in (None, ""), f"{field}.currency is required when amount_state is clear")
    entry["payee"] = core.validate_payee(
        entry.get("payee"),
        field=f"{field}.payee",
        cash=kind == "cash",
    )


def _validate_network_fees(value: object, *, field: str) -> None:
    if value in (None, "", []):
        return
    core.require(isinstance(value, list), f"{field} must be a list")
    for position, fee in enumerate(value):
        core.require(isinstance(fee, Mapping), f"{field}[{position}] must be an object")
        if core.clean_text(fee.get("kind")).casefold() == "network_fee":
            core.require(
                fee.get("customer_requested") is True,
                f"{field}[{position}].customer_requested must be true for a network fee",
            )


def _validate_pricing_scope(value: dict[str, Any], *, field: str) -> list[str]:
    """Validate one order or leg pricing judgment and return missing state fields."""

    rate = core.parse_decimal(value.get("rate"), field=f"{field}.rate", allow_none=True)
    if rate is not None:
        core.require(rate > 0, f"{field}.rate must be positive")
        value["rate"] = core.decimal_text(rate)
    expected_payout = core.parse_decimal(
        value.get("expected_payout"),
        field=f"{field}.expected_payout",
        allow_none=True,
    )
    if expected_payout is not None:
        core.require(expected_payout >= 0, f"{field}.expected_payout cannot be negative")
        value["expected_payout"] = core.decimal_text(expected_payout)
    operator = core.clean_text(value.get("rate_operator")).casefold()
    if operator:
        core.require(operator in simple_ledger.RATE_OPERATORS, f"{field}.rate_operator is unsupported")
        value["rate_operator"] = operator

    missing = [
        f"{field}.{state_field}"
        for state_field in ("rate_state", "expected_payout_state")
        if not core.clean_text(value.get(state_field))
    ]
    if missing:
        return missing

    rate_state = core.clean_text(value.get("rate_state")).casefold()
    expected_state = core.clean_text(value.get("expected_payout_state")).casefold()
    core.require(rate_state in RATE_STATES, f"{field}.rate_state is unsupported")
    core.require(
        expected_state in EXPECTED_PAYOUT_STATES,
        f"{field}.expected_payout_state is unsupported",
    )
    value["rate_state"] = rate_state
    value["expected_payout_state"] = expected_state

    if rate_state == "adopted":
        core.require(rate is not None, f"{field}.rate is required when rate_state is adopted")
    else:
        core.require(
            rate is None,
            f"{field}.rate must be empty when rate_state is {rate_state}",
        )
        core.require(
            not operator,
            f"{field}.rate_operator must be empty when rate_state is {rate_state}",
        )

    if expected_state == "explicit":
        core.require(
            expected_payout is not None,
            f"{field}.expected_payout is required when expected_payout_state is explicit",
        )
    elif expected_state == "calculated_from_rate":
        core.require(
            expected_payout is None,
            f"{field}.expected_payout must be empty when it is calculated from rate",
        )
        core.require(
            rate_state == "adopted" and rate is not None,
            f"{field}.rate_state must be adopted when expected payout is calculated from rate",
        )
        core.require(
            bool(operator),
            f"{field}.rate_operator is required when rate must calculate expected payout",
        )
    else:
        core.require(
            expected_payout is None,
            f"{field}.expected_payout must be empty when expected_payout_state is {expected_state}",
        )
        core.require(
            rate_state != "adopted",
            f"{field}.expected_payout_state must be calculated_from_rate when an adopted rate supplies the result",
        )

    core.require(not operator or rate is not None, f"{field}.rate_operator requires a rate")
    return []


def _validate_decision(
    normalized: Mapping[str, Any],
    group: Mapping[str, Any],
    decision: dict[str, Any],
    *,
    require_complete: bool,
    capture_hashes: bool,
) -> dict[str, int]:
    expected = _decision_template(normalized, group)
    core.require(decision.get("contract_version") == DECISION_CONTRACT, "unsupported decision contract")
    for field in (
        "normalized_source_fingerprint",
        "group_fingerprint",
        "group_key",
        "group_name",
        "platform",
        "message_count",
        "evidence_media_count",
    ):
        core.require(decision.get(field) == expected.get(field), f"{group.get('group_key')}: generated {field} was edited")
    messages = list(group.get("messages", []))
    reviewed_through = decision.get("reviewed_through")
    core.require(isinstance(reviewed_through, int) and 0 <= reviewed_through <= len(messages), "invalid reviewed_through")
    core.require(isinstance(decision.get("read_complete"), bool), "read_complete must be boolean")
    if require_complete:
        core.require(decision.get("read_complete") is True and reviewed_through == len(messages), "group chronology has not been read completely")

    inventory = _media_inventory(group)
    available = {label for label, (_, media) in inventory.items() if media.get("availability") == "available"}
    missing = set(inventory) - available
    media_decisions = decision.get("media_decisions")
    core.require(isinstance(media_decisions, dict), "media_decisions must be an object keyed by M labels")
    unknown_labels = sorted(set(media_decisions) - available)
    core.require(not unknown_labels, f"media decisions contain missing or unknown labels: {unknown_labels[:10]}")
    if require_complete:
        unclassified = sorted(available - set(media_decisions))
        core.require(not unclassified, f"unclassified available media: {unclassified[:20]}")

    entry_ids: set[str] = set()
    reference_count = 0
    fund_media_count = 0
    for label, raw in media_decisions.items():
        field = f"{group.get('group_key')}.media_decisions.{label}"
        core.require(isinstance(raw, dict), f"{field} must be an object")
        classification = core.clean_text(raw.get("classification")).casefold()
        core.require(classification in {"reference", "fund"}, f"{field}.classification must be reference or fund")
        if classification == "reference":
            unknown = sorted(set(raw) - {"classification", "note"})
            core.require(not unknown, f"{field}: reference has unsupported fields: {', '.join(unknown)}")
            reference_count += 1
            continue
        unknown = sorted(set(raw) - {"classification", "viewed_original", "evidence_sha256", "entries", "note"})
        core.require(not unknown, f"{field}: fund decision has unsupported fields: {', '.join(unknown)}")
        core.require(raw.get("viewed_original") is True, f"{field}: open the original image before recording fund facts")
        entries = raw.get("entries")
        core.require(isinstance(entries, list) and entries, f"{field}.entries must contain at least one fund entry")
        _, media = inventory[label]
        media_path = Path(str(media.get("path") or ""))
        core.require(media_path.is_file(), f"{field}: original media file is unavailable: {media_path}")
        actual_hash = core.sha256_file(media_path)
        recorded_hash = core.clean_text(raw.get("evidence_sha256"))
        if recorded_hash:
            core.require(recorded_hash == actual_hash, f"{field}: original fund evidence changed after review")
        elif capture_hashes:
            raw["evidence_sha256"] = actual_hash
        elif require_complete:
            raise ValueError(f"{field}: evidence hash has not been captured; run review seal")
        for position, entry in enumerate(entries, start=1):
            core.require(isinstance(entry, dict), f"{field}.entries[{position - 1}] must be an object")
            _validate_entry(entry, field=f"{field}.entries[{position - 1}]")
            entry_id = f"{label}.{position}"
            core.require(entry_id not in entry_ids, f"duplicate entry id: {entry_id}")
            entry_ids.add(entry_id)
        fund_media_count += 1

    label_by_message_id, message_by_label = _message_labels(group)
    del label_by_message_id
    orders = decision.get("orders")
    core.require(isinstance(orders, list), "orders must be a list")
    order_ids: set[str] = set()
    assigned_entries: set[str] = set()
    pricing_scopes = 0
    missing_pricing_states: list[str] = []
    order_fields = {
        "id",
        "entry_ids",
        "source_messages",
        "customer_id",
        "customer_nickname",
        "direction",
        "rate_state",
        "rate",
        "rate_operator",
        "expected_payout_state",
        "expected_payout",
        "note",
        "same_transactions",
        "legs",
        "fees",
        "rounding",
    }
    for position, order in enumerate(orders):
        field = f"{group.get('group_key')}.orders[{position}]"
        core.require(isinstance(order, dict), f"{field} must be an object")
        unknown = sorted(set(order) - order_fields)
        core.require(not unknown, f"{field}: unsupported fields: {', '.join(unknown)}")
        for required_field in ("customer_id", "customer_nickname", "direction"):
            core.require(
                required_field in order,
                f"{field}.{required_field} must be explicitly supplied by the model",
            )
        order_note = core.clean_text(order.get("note"))
        if order_note:
            order["note"] = order_note
        order_id = core.clean_text(order.get("id"))
        core.require(bool(order_id), f"{field}.id is required")
        core.require(order_id not in order_ids, f"{field}.id is duplicated")
        order_ids.add(order_id)
        raw_entries = order.get("entry_ids")
        core.require(isinstance(raw_entries, list) and raw_entries, f"{field}.entry_ids is required")
        refs = [str(item) for item in raw_entries]
        core.require(len(refs) == len(set(refs)), f"{field}.entry_ids repeats an entry")
        for entry_id in refs:
            core.require(entry_id in entry_ids, f"{field}: unknown fund entry {entry_id}")
            core.require(entry_id not in assigned_entries, f"fund entry assigned to multiple orders: {entry_id}")
            assigned_entries.add(entry_id)
        raw_messages = order.get("source_messages")
        core.require(isinstance(raw_messages, list) and raw_messages, f"{field}.source_messages is required")
        source_labels = [str(item) for item in raw_messages]
        core.require(len(source_labels) == len(set(source_labels)), f"{field}.source_messages repeats a label")
        core.require(set(source_labels) <= set(message_by_label), f"{field}.source_messages contains an unknown label")
        if order.get("direction") not in (None, ""):
            order["direction"] = core.canonical_direction(order.get("direction"), field=f"{field}.direction")
        _validate_network_fees(order.get("fees"), field=f"{field}.fees")
        same_transactions = order.get("same_transactions")
        if same_transactions not in (None, "", []):
            core.require(isinstance(same_transactions, list), f"{field}.same_transactions must be a list")
            duplicate_entries: set[str] = set()
            canonical_entries: set[str] = set()
            for pair_position, pair in enumerate(same_transactions):
                pair_field = f"{field}.same_transactions[{pair_position}]"
                core.require(isinstance(pair, Mapping), f"{pair_field} must be an object")
                unknown_pair_fields = sorted(set(pair) - {"entry_id", "same_as"})
                core.require(not unknown_pair_fields, f"{pair_field}: unsupported fields: {', '.join(unknown_pair_fields)}")
                entry_id = core.clean_text(pair.get("entry_id"))
                same_as = core.clean_text(pair.get("same_as"))
                core.require(entry_id in refs and same_as in refs, f"{pair_field} must cite two entries in this order")
                core.require(entry_id != same_as, f"{pair_field} cannot cite itself")
                core.require(entry_id not in duplicate_entries, f"{pair_field} repeats a duplicate entry")
                duplicate_entries.add(entry_id)
                canonical_entries.add(same_as)
            core.require(
                not (duplicate_entries & canonical_entries),
                f"{field}.same_transactions cannot contain chains",
            )

        legs = order.get("legs")
        if legs not in (None, "", []):
            core.require(isinstance(legs, list) and len(legs) >= 2, f"{field}.legs requires at least two legs")
            for pricing_field in (
                "rate_state",
                "rate",
                "rate_operator",
                "expected_payout_state",
                "expected_payout",
            ):
                core.require(
                    order.get(pricing_field) in (None, ""),
                    f"{field}.{pricing_field} must be empty when pricing is defined per leg",
                )
            referenced_settlements: set[str] = set()
            for leg_position, leg in enumerate(legs):
                leg_field = f"{field}.legs[{leg_position}]"
                core.require(isinstance(leg, dict), f"{leg_field} must be an object")
                allowed_leg_fields = {
                    "leg_id",
                    "direction",
                    "allocation_amount",
                    "rate_state",
                    "rate",
                    "rate_operator",
                    "expected_payout_state",
                    "expected_payout",
                    "payout_entry_ids",
                    "recovery_entry_ids",
                    "fees",
                    "rounding",
                }
                unknown_leg_fields = sorted(set(leg) - allowed_leg_fields)
                core.require(not unknown_leg_fields, f"{leg_field}: unsupported fields: {', '.join(unknown_leg_fields)}")
                leg_direction = core.canonical_direction(leg.get("direction"), field=f"{leg_field}.direction")
                if isinstance(leg, dict):
                    leg["direction"] = leg_direction
                allocation = core.parse_decimal(leg.get("allocation_amount"), field=f"{leg_field}.allocation_amount")
                core.require(allocation is not None and allocation > 0, f"{leg_field}.allocation_amount must be positive")
                if isinstance(leg, dict):
                    leg["allocation_amount"] = core.decimal_text(allocation)
                for reference_field in ("payout_entry_ids", "recovery_entry_ids"):
                    raw_references = leg.get(reference_field, [])
                    core.require(isinstance(raw_references, list), f"{leg_field}.{reference_field} must be a list")
                    entry_references = [str(item) for item in raw_references]
                    core.require(
                        len(entry_references) == len(set(entry_references)),
                        f"{leg_field}.{reference_field} repeats an entry",
                    )
                    core.require(
                        set(entry_references) <= set(refs),
                        f"{leg_field}.{reference_field} cites an entry outside this order",
                    )
                    core.require(
                        not (referenced_settlements & set(entry_references)),
                        f"{leg_field}.{reference_field} repeats an entry used by another leg",
                    )
                    referenced_settlements.update(entry_references)
                pricing_scopes += 1
                missing_pricing_states.extend(
                    _validate_pricing_scope(leg, field=leg_field)
                )
                _validate_network_fees(
                    leg.get("fees"),
                    field=f"{leg_field}.fees",
                )
        else:
            pricing_scopes += 1
            missing_pricing_states.extend(
                _validate_pricing_scope(order, field=field)
            )

    if require_complete:
        core.require(
            not missing_pricing_states,
            "pricing states are required before seal: "
            + ", ".join(missing_pricing_states[:40]),
        )

    if require_complete:
        unassigned_entries = sorted(entry_ids - assigned_entries)
        core.require(
            not unassigned_entries,
            f"fund entries must be explicitly assigned to an order: {unassigned_entries}",
        )

    links = decision.get("balance_links")
    core.require(isinstance(links, list), "balance_links must be a list")
    for position, link in enumerate(links):
        field = f"{group.get('group_key')}.balance_links[{position}]"
        core.require(isinstance(link, dict), f"{field} must be an object")
        allowed_link_fields = {
            "source_order_id",
            "target_order_id",
            "kind",
            "amount",
            "currency",
            "source_messages",
            "already_in_expected",
        }
        unknown_link_fields = sorted(set(link) - allowed_link_fields)
        core.require(not unknown_link_fields, f"{field}: unsupported fields: {', '.join(unknown_link_fields)}")
        source_order_id = core.clean_text(link.get("source_order_id"))
        target_order_id = core.clean_text(link.get("target_order_id"))
        core.require(source_order_id in order_ids, f"{field}.source_order_id is unknown")
        core.require(target_order_id in order_ids, f"{field}.target_order_id is unknown")
        core.require(source_order_id != target_order_id, f"{field} cannot link an order to itself")
        kind = core.clean_text(link.get("kind")).casefold()
        core.require(kind in {"shortfall_carryover", "overpayment_carryover"}, f"{field}.kind is unsupported")
        link["kind"] = kind
        amount = core.parse_decimal(link.get("amount"), field=f"{field}.amount")
        core.require(amount is not None and amount > 0, f"{field}.amount must be positive")
        link["amount"] = core.decimal_text(amount)
        link["currency"] = core.normalize_currency(link.get("currency"), field=f"{field}.currency")
        already_in_expected = link.get("already_in_expected", False)
        core.require(isinstance(already_in_expected, bool), f"{field}.already_in_expected must be boolean")
        if "already_in_expected" in link:
            link["already_in_expected"] = already_in_expected
        source_messages = link.get("source_messages", [])
        core.require(isinstance(source_messages, list), f"{field}.source_messages must be a list")
        source_labels = [str(item) for item in source_messages]
        core.require(len(source_labels) == len(set(source_labels)), f"{field}.source_messages repeats a label")
        core.require(set(source_labels) <= set(message_by_label), f"{field}.source_messages contains an unknown label")

    return {
        "available_media": len(available),
        "missing_media": len(missing),
        "classified_media": len(media_decisions),
        "reference_media": reference_count,
        "fund_media": fund_media_count,
        "fund_entries": len(entry_ids),
        "orders": len(orders),
        "pricing_scopes": pricing_scopes,
        "missing_pricing_states": len(missing_pricing_states),
        "unassigned_entries": len(entry_ids - assigned_entries),
    }


def _compact_page(group: Mapping[str, Any], start: int, end: int) -> tuple[list[dict[str, Any]], int]:
    messages = list(group.get("messages", []))
    label_by_id, _ = _message_labels(group)
    message_by_id = {
        str(message.get("message_id") or ""): message
        for message in messages
    }
    media_labels = {
        str(media.get("media_id") or ""): label
        for label, (_, media) in _media_inventory(group).items()
    }
    result: list[dict[str, Any]] = []
    collapsed = 0
    for index in range(start, end):
        message = messages[index]
        media_items = []
        for media in message.get("media", []):
            media_id = str(media.get("media_id") or "")
            if media_id not in media_labels:
                continue
            media_items.append(
                {
                    "label": media_labels[media_id],
                    "kind": media.get("kind"),
                    "availability": media.get("availability"),
                    "path": media.get("path"),
                }
            )
        if message.get("excluded_from_accounting") and not media_items:
            collapsed += 1
            continue
        reply_id = str(message.get("reply_to_message_id") or "")
        reply = None
        if reply_id in label_by_id:
            target = message_by_id.get(reply_id)
            reply = {
                "label": label_by_id[reply_id],
                "sender": target.get("sender_name") if target else None,
                "text": target.get("text") if target else None,
            }
        result.append(
            {
                "label": f"S{index + 1:05d}",
                "timestamp": message.get("timestamp"),
                "sender": message.get("sender_name") or message.get("sender_id"),
                "role": message.get("role"),
                "text": message.get("text"),
                "reply": reply,
                "media": media_items,
            }
        )
    return result, collapsed


def review_command(args: argparse.Namespace) -> dict[str, Any]:
    work = args.work.resolve()
    run = _load_run(work)
    normalized = _load_snapshot(work, run)
    groups = _group_index(normalized)
    if args.review_action == "status":
        selected = (
            [_select_run_group(run, args.group)]
            if args.group
            else [item for item in run["groups"] if isinstance(item, Mapping)]
        )
        reports = []
        for run_group in selected:
            decision = _load_json(_decision_path(work, run_group))
            reports.append(
                {
                    "group_key": run_group.get("group_key"),
                    "run_label": run_group.get("run_label"),
                    "reviewed_through": decision.get("reviewed_through"),
                    "message_count": decision.get("message_count"),
                    "classified_media": len(decision.get("media_decisions", {})),
                    "evidence_media_count": decision.get("evidence_media_count"),
                    "sealed": bool(decision.get("sealed")),
                }
            )
        return {"groups": reports}

    run_group = _select_run_group(
        run,
        args.group,
        prefer_incomplete=args.review_action == "next",
    )
    group_key = str(run_group["group_key"])
    group = groups[group_key]
    decision_path = _decision_path(work, run_group)
    decision = _load_json(decision_path)

    if args.review_action == "next":
        core.require(1 <= args.limit <= MAX_PAGE_SIZE, f"--limit must be 1..{MAX_PAGE_SIZE}")
        start = int(decision.get("reviewed_through") or 0)
        messages = list(group.get("messages", []))
        end = min(start + args.limit, len(messages))
        compact, collapsed = _compact_page(group, start, end)
        decision["reviewed_through"] = end
        decision["read_complete"] = end == len(messages)
        decision["sealed"] = False
        decision["sealed_decision_fingerprint"] = None
        core.atomic_json(decision_path, decision)
        run_group["sealed"] = False
        core.atomic_json(_run_path(work), run)
        return {
            "contract_version": REVIEW_PAGE_CONTRACT,
            "group_key": group_key,
            "run_label": run_group.get("run_label"),
            "page_start": start,
            "page_end": end,
            "message_count": len(messages),
            "messages": compact,
            "collapsed_system_or_empty_messages": collapsed,
            "done": end == len(messages),
        }

    prior_fingerprint = _semantic_fingerprint(decision)
    statistics = _validate_decision(
        normalized,
        group,
        decision,
        require_complete=args.review_action == "seal",
        capture_hashes=True,
    )
    current_fingerprint = _semantic_fingerprint(decision)
    if args.review_action == "seal":
        decision["sealed"] = True
        decision["sealed_decision_fingerprint"] = current_fingerprint
    elif prior_fingerprint != current_fingerprint or decision.get("sealed_decision_fingerprint") != current_fingerprint:
        decision["sealed"] = False
        decision["sealed_decision_fingerprint"] = None
    core.atomic_json(decision_path, decision)
    run_group["sealed"] = bool(decision.get("sealed"))
    core.atomic_json(_run_path(work), run)
    return {
        "group_key": group_key,
        "run_label": run_group.get("run_label"),
        "sealed": bool(decision.get("sealed")),
        **statistics,
    }


def _entry_event_type(entry: Mapping[str, Any]) -> str:
    if core.clean_text(entry.get("side")).casefold() == "recovery":
        return "payout_recovery"
    if core.clean_text(entry.get("kind")).casefold() == "cash":
        return "cash_payment"
    return "payment_screenshot"


def _translated_advanced_order(
    raw: Mapping[str, Any],
    *,
    event_by_entry: Mapping[str, str],
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    same_transactions = raw.get("same_transactions")
    if same_transactions not in (None, "", []):
        core.require(isinstance(same_transactions, list), "same_transactions must be a list")
        result["same_transactions"] = [
            {
                "event_id": event_by_entry[str(item.get("entry_id") or "")],
                "same_as": event_by_entry[str(item.get("same_as") or "")],
            }
            for item in same_transactions
        ]
    legs = raw.get("legs")
    if legs not in (None, "", []):
        core.require(isinstance(legs, list), "legs must be a list")
        compiled_legs = []
        for item in legs:
            core.require(isinstance(item, Mapping), "each leg must be an object")
            leg = {
                key: copy.deepcopy(value)
                for key, value in item.items()
                if key
                not in {
                    "payout_entry_ids",
                    "recovery_entry_ids",
                    "rate_state",
                    "expected_payout_state",
                }
            }
            leg["payout_event_ids"] = [
                event_by_entry[str(value)] for value in item.get("payout_entry_ids", [])
            ]
            if item.get("recovery_entry_ids") not in (None, []):
                leg["recovery_event_ids"] = [
                    event_by_entry[str(value)] for value in item.get("recovery_entry_ids", [])
                ]
            compiled_legs.append(leg)
        result["legs"] = compiled_legs
    for field in ("fees", "rounding"):
        if raw.get(field) not in (None, "", [], {}):
            result[field] = copy.deepcopy(raw[field])
    return result


def _compile_decisions_v2(
    normalized: Mapping[str, Any],
    decisions: Mapping[str, Mapping[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    events: list[dict[str, Any]] = []
    plan_groups: list[dict[str, Any]] = []
    for group in normalized.get("groups", []):
        group_key = str(group["group_key"])
        decision = decisions[group_key]
        inventory = _media_inventory(group)
        media_decisions = decision["media_decisions"]
        event_by_entry: dict[str, str] = {}
        side_by_entry: dict[str, str] = {}
        for label, (message, media) in inventory.items():
            media_id = str(media.get("media_id") or "")
            message_id = str(message.get("message_id") or "")
            if media.get("availability") == "missing":
                continue
            media_decision = media_decisions[label]
            if media_decision["classification"] == "reference":
                events.append(
                    {
                        "event_id": f"{media_id}#reference",
                        "group_key": group_key,
                        "message_id": message_id,
                        "media_id": media_id,
                        "type": "irrelevant_media",
                        "note": media_decision.get("note"),
                    }
                )
                continue
            for position, entry in enumerate(media_decision["entries"], start=1):
                entry_id = f"{label}.{position}"
                event_id = f"{media_id}#entry:{position}"
                amount_state = str(entry["amount_state"])
                event = {
                    "event_id": event_id,
                    "group_key": group_key,
                    "message_id": message_id,
                    "media_id": media_id,
                    "type": _entry_event_type(entry),
                    "evidence_sha256": media_decision["evidence_sha256"],
                    "ocr": {
                        "amount": entry.get("amount"),
                        "amount_text": entry.get("amount_text"),
                        "currency": entry.get("currency"),
                        "payee": entry["payee"],
                        "status_text": entry.get("status_text"),
                        "status_class": STATUS_CLASS[str(entry["result"])],
                        "status_class_confidence": "high",
                        "amount_completeness": "complete" if amount_state == "clear" else amount_state,
                        "confidence": "high" if amount_state == "clear" else "low",
                    },
                }
                event["flow_side"] = entry["side"]
                side_by_entry[entry_id] = str(entry["side"])
                if entry.get("note"):
                    event["note"] = entry["note"]
                events.append(event)
                event_by_entry[entry_id] = event_id

        message_label_by_id, message_by_label = _message_labels(group)
        del message_label_by_id
        plan_orders: list[dict[str, Any]] = []
        case_by_order_id: dict[str, str] = {}
        for raw in decision["orders"]:
            short_order_id = str(raw["id"])
            case_id = f"{group_key}:{short_order_id}"
            case_by_order_id[short_order_id] = case_id
            entry_ids = [str(item) for item in raw["entry_ids"]]
            order: dict[str, Any] = {
                "case_id": case_id,
                "event_ids": [event_by_entry[item] for item in entry_ids],
                "source_message_ids": [
                    str(message_by_label[str(item)]["message_id"])
                    for item in raw["source_messages"]
                ],
            }
            for field in (
                "customer_id",
                "customer_nickname",
                "direction",
            ):
                order[field] = copy.deepcopy(raw.get(field))
            for field in (
                "rate",
                "rate_operator",
                "expected_payout",
                "note",
            ):
                if raw.get(field) not in (None, ""):
                    order[field] = copy.deepcopy(raw[field])
            event_sides = {
                event_by_entry[entry_id]: side_by_entry[entry_id]
                for entry_id in entry_ids
                if entry_id in side_by_entry
            }
            order["event_sides"] = event_sides
            order.update(_translated_advanced_order(raw, event_by_entry=event_by_entry))
            plan_orders.append(order)
        balance_links = []
        for raw_link in decision.get("balance_links", []):
            link = {
                key: copy.deepcopy(value)
                for key, value in raw_link.items()
                if key not in {"source_order_id", "target_order_id", "source_messages"}
            }
            link["source_case_id"] = case_by_order_id[core.clean_text(raw_link.get("source_order_id"))]
            link["target_case_id"] = case_by_order_id[core.clean_text(raw_link.get("target_order_id"))]
            if raw_link.get("source_messages"):
                link["source_message_ids"] = [
                    str(message_by_label[str(item)]["message_id"])
                    for item in raw_link["source_messages"]
                ]
            balance_links.append(link)
        plan_groups.append(
            {
                "group_key": group_key,
                "orders": plan_orders,
                "balance_links": balance_links,
            }
        )
    events_fingerprint = core.fingerprint_json(events)
    plan = {
        "contract_version": simple_ledger.SIMPLE_PLAN_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "events_fingerprint": events_fingerprint,
        "groups": plan_groups,
    }
    return events, plan


def finish_run(args: argparse.Namespace) -> dict[str, Any]:
    work = args.work.resolve()
    output = args.output.resolve()
    core.require(not output.exists(), f"output already exists; choose a new file: {output}")
    run = _load_run(work)
    normalized = _load_snapshot(work, run)
    groups = _group_index(normalized)
    decisions: dict[str, Mapping[str, Any]] = {}
    review_statistics: dict[str, dict[str, int]] = {}
    for run_group in run["groups"]:
        group_key = str(run_group["group_key"])
        decision = _load_json(_decision_path(work, run_group))
        core.require(decision.get("sealed") is True, f"group is not sealed: {run_group.get('run_label')}")
        current_fingerprint = _semantic_fingerprint(decision)
        core.require(
            decision.get("sealed_decision_fingerprint") == current_fingerprint,
            f"sealed decision changed; review and seal again: {run_group.get('run_label')}",
        )
        statistics = _validate_decision(
            normalized,
            groups[group_key],
            decision,
            require_complete=True,
            capture_hashes=False,
        )
        decisions[group_key] = decision
        review_statistics[group_key] = statistics

    events, plan = _compile_decisions_v2(normalized, decisions)
    events_fingerprint = core.fingerprint_json(events)
    core.require(plan["events_fingerprint"] == events_fingerprint, "internal events fingerprint mismatch")
    orders, ledger_statistics = simple_ledger.compile_simple_ledger(
        normalized,
        events,
        events_fingerprint,
        plan,
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="finish-", dir=work) as temporary_name:
        temporary = Path(temporary_name)
        orders_path = temporary / "orders.json"
        workbook_path = temporary / "workbook.xlsx"
        core.atomic_json(orders_path, orders)
        workbook = build_workbook.build_workbook(
            build_workbook.validate_orders(orders),
            args.template.resolve(),
        )
        workbook.save(workbook_path)
        workbook.close()
        errors = check_workbook.check(workbook_path, orders_path)
        core.require(not errors, "workbook verification failed: " + "; ".join(errors[:10]))
        workbook_path.replace(output)

    run["status"] = "finished"
    run["finished_at"] = datetime.now(timezone.utc).isoformat()
    run["output"] = str(output)
    core.atomic_json(_run_path(work), run)
    return {
        "output": str(output),
        "groups": ledger_statistics["groups"],
        "orders": ledger_statistics["orders"],
        "fund_entries": sum(item["fund_entries"] for item in review_statistics.values()),
        "pending_orders": ledger_statistics["pending_orders"],
        "warnings": ledger_statistics["warnings"],
    }


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        if args.command == "start":
            result = start_run(args)
        elif args.command == "review":
            result = review_command(args)
        else:
            result = finish_run(args)
    except (
        OSError,
        ValueError,
        KeyError,
        TypeError,
        json.JSONDecodeError,
        ZoneInfoNotFoundError,
    ) as exc:
        print(f"Reconciliation failed: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
