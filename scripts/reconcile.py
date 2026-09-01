#!/usr/bin/env python3
"""Run a fresh group-chat reconciliation through start, review, and finish."""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Iterable, Mapping
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import build_workbook
import check_workbook
import core
import extract_line_android_miui
import extract_line_ios
import normalize_exports
import simple_ledger


RUN_CONTRACT = "group-chat-reconcile-run/1.0"
DECISION_CONTRACT = "group-chat-decision/3.2"
REVIEW_PAGE_CONTRACT = "group-chat-review-page/2.0"
REVIEW_BATCH_CONTRACT = "group-chat-review-batch/1.1"
EDIT_CONTROL_MODE = "review-apply-batch/1.1"
LEGACY_DECISION_CONTRACT = "group-chat-decision/3.1"
LEGACY_EDIT_CONTROL_MODE = "review-apply-batch/1.0"
NORMALIZER_VERSION = "reconcile-start/1.0"
DEFAULT_PAGE_SIZE = 200
MAX_PAGE_SIZE = 500
DEFAULT_REVIEW_PAGE_OUTPUT_CHAR_BUDGET = 14_000
UNKNOWN_PAYEE_WARNING_MIN_TRANSFERS = 10
FLOW_SIDES = frozenset({"payment", "payment_refund", "payout", "recovery", "unknown"})
SIDE_EXCEPTION_SIDES = {
    "relayed_customer_payment": "payment",
    "relayed_internal_payout": "payout",
    "explicit_payment_refund": "payment_refund",
    "explicit_recovery": "recovery",
}
ENTRY_KINDS = frozenset({"transfer", "cash"})
ENTRY_RESULTS = frozenset({"completed", "failed", "pending", "not_shown", "unknown"})
AMOUNT_STATES = frozenset({"clear", "partial", "unreadable"})
NORMAL_PROCESSING_STATUS_MARKERS = (
    "确认中",
    "处理中",
    "等待确认",
    "待区块确认",
    "pending",
    "processing",
    "confirming",
    "unconfirmed",
)
FAILURE_STATUS_MARKERS = (
    "失败",
    "取消",
    "拒绝",
    "作废",
    "无效",
    "风控导致未完成",
    "风控未完成",
    "风控未通过",
    "风控拦截",
    "failed",
    "declined",
    "rejected",
    "cancelled",
    "canceled",
    "invalid",
)
MISSING_CONTEXT_NOTE_MARKERS = (
    "客户",
    "方向",
    "付款",
    "回款",
    "汇率",
    "计价",
    "金额",
    "币种",
    "收款方",
    "上下文",
    "无法辨认",
    "未说明",
    "缺少",
)
PRICING_EXPECTED_KINDS = frozenset(
    {"explicit", "calculated_from_terms", "unknown"}
)
PRICING_UNKNOWN_REASONS = frozenset(
    {"not_stated", "conflicting_authority", "incomplete_formula"}
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
    start.add_argument("--date", help="optional accounting date in YYYY-MM-DD")
    start.add_argument(
        "--from",
        dest="accounting_from",
        metavar="YYYY-MM-DD HH:MM",
        help="optional inclusive accounting start time",
    )
    start.add_argument(
        "--to",
        dest="accounting_to",
        metavar="YYYY-MM-DD HH:MM",
        help="optional exclusive accounting end time",
    )
    start.add_argument("--timezone", default="Asia/Bangkok")
    start.add_argument(
        "--roster",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "config" / "roster.yaml",
    )
    start.add_argument("--line-backup", action="append", default=[], type=Path)
    start.add_argument("--line-android-backup", action="append", default=[], type=Path)
    start.add_argument("--line-self-name", default="LINE_SELF")

    review = commands.add_parser("review", help="read, validate, or seal one group review")
    review.add_argument("work", type=Path)
    review_actions = review.add_subparsers(dest="review_action", required=True)
    next_page = review_actions.add_parser("next", help="return the next compact chronological page")
    next_page.add_argument("--group", help="group key, run label, or unique exact group name")
    next_page.add_argument(
        "--limit",
        type=int,
        default=None,
        help="legacy exact message count; omit for an adaptive complete page",
    )
    status = review_actions.add_parser("status", help="show review progress")
    status.add_argument("--group", help="optional group selector")
    audit = review_actions.add_parser("audit", help="report review-risk diagnostics without writing files")
    audit.add_argument("--group", help="optional group selector")
    check = review_actions.add_parser("check", help="validate the current decision data")
    check.add_argument("--group", required=True)
    upgrade_checkpoints = review_actions.add_parser(
        "upgrade-checkpoints",
        help="preserve decisions but reset legacy read progress for safe page commits",
    )
    upgrade_checkpoints.add_argument("--group", required=True)
    seal = review_actions.add_parser("seal", help="validate and seal a completely reviewed group")
    seal.add_argument("--group", required=True)
    apply_batch = review_actions.add_parser(
        "apply-batch",
        help="atomically merge one schema-validated semantic review batch",
    )
    apply_batch.add_argument("--group", required=True)
    apply_batch.add_argument("--input", required=True, type=Path)

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


def _discover_line_android_backups(
    inputs: Iterable[Path], explicit: Iterable[Path]
) -> list[Path]:
    candidates: set[Path] = set()
    for raw in explicit:
        candidates.add(extract_line_android_miui.find_backup_file(raw).resolve())
    for raw in inputs:
        path = raw.resolve()
        if path.is_file():
            if extract_line_android_miui.is_line_android_backup(path) and not _generated_path(path):
                candidates.add(path)
            continue
        if not path.is_dir():
            continue
        for backup in path.rglob("*.bak"):
            if not _generated_path(backup) and extract_line_android_miui.is_line_android_backup(backup):
                candidates.add(backup.resolve())
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


def _line_android_documents(
    work: Path,
    backups: Iterable[Path],
    *,
    contains: str,
    timezone_name: str,
    roster_path: Path,
) -> list[dict[str, Any]]:
    documents: list[dict[str, Any]] = []
    needle = contains.casefold()
    for index, backup in enumerate(backups, start=1):
        groups = extract_line_android_miui.available_groups_for_backup(backup)
        if not any(needle in str(group.get("group_name") or "").casefold() for group in groups):
            continue
        output = work / "snapshot" / "sources" / f"line_android_{index:02d}.json"
        media_dir = work / "snapshot" / "line_android_media" / f"backup_{index:02d}"
        exit_code = extract_line_android_miui.main(
            [
                str(backup),
                "-o",
                str(output),
                "--timezone",
                timezone_name,
                "--group-pattern",
                re.escape(contains),
                "--roster",
                str(roster_path),
                "--media-dir",
                str(media_dir),
            ]
        )
        core.require(exit_code == 0, f"LINE Android extraction failed for {backup}")
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
    decision = {
        "contract_version": DECISION_CONTRACT,
        "normalized_source_fingerprint": normalized.get("source_fingerprint"),
        "group_fingerprint": core.fingerprint_json(group),
        "group_key": group.get("group_key"),
        "group_name": group.get("group_name"),
        "platform": group.get("platform"),
        "message_count": len(group.get("messages", [])),
        "evidence_media_count": len(evidence),
        "reviewed_through": 0,
        "read_complete": not bool(group.get("messages", [])),
        "sealed": False,
        "sealed_decision_fingerprint": None,
        "media_decisions": {},
        "orders": [],
        "open_orders": [],
        "balance_links": [],
        "settlement_allocations": [],
        "unknown_payee_reviewed_entry_ids": [],
    }
    decision["edit_control"] = _new_edit_control(_semantic_fingerprint(decision))
    return decision


def _platform_prefix(platform: object) -> str:
    return {"Telegram": "TG", "WhatsApp": "WA", "LINE": "LINE"}.get(str(platform), "CHAT")


def _normalize_accounting_date(value: object) -> str | None:
    if value is None:
        return None
    cleaned = core.clean_text(value)
    error = "--date must be a valid date in YYYY-MM-DD format"
    core.require(bool(cleaned), error)
    try:
        parsed = datetime.strptime(cleaned, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(error) from exc
    core.require(parsed.isoformat() == cleaned, error)
    return cleaned


def _normalize_accounting_time(value: object, *, option: str) -> str | None:
    if value is None:
        return None
    cleaned = core.clean_text(value)
    error = f"{option} must be a valid time in YYYY-MM-DD HH:MM format"
    core.require(bool(cleaned), error)
    try:
        parsed = datetime.strptime(cleaned, "%Y-%m-%d %H:%M")
    except ValueError as exc:
        raise ValueError(error) from exc
    core.require(parsed.strftime("%Y-%m-%d %H:%M") == cleaned, error)
    return cleaned


def _filter_normalized_window(
    normalized: dict[str, Any],
    *,
    start_at: datetime,
    end_at: datetime,
    fingerprint_fields: Mapping[str, object],
) -> dict[str, Any]:
    core.require(start_at.tzinfo is not None and end_at.tzinfo is not None, "accounting window requires timezone")
    filtered = copy.deepcopy(normalized)
    for group in filtered.get("groups", []):
        group["messages"] = [
            message
            for message in group.get("messages", [])
            if start_at
            <= datetime.fromisoformat(str(message.get("timestamp"))).astimezone(start_at.tzinfo)
            < end_at
        ]
    messages = [message for group in filtered.get("groups", []) for message in group.get("messages", [])]
    media = [item for message in messages for item in message.get("media", [])]
    filtered["source_fingerprint"] = core.fingerprint_json(
        {
            "base_source_fingerprint": normalized.get("source_fingerprint"),
            **fingerprint_fields,
        }
    )
    statistics = dict(normalized.get("statistics", {}))
    statistics.update(
        {
            "groups": len(filtered.get("groups", [])),
            "messages": len(messages),
            "excluded_messages": sum(
                bool(message.get("excluded_from_accounting")) for message in messages
            ),
            "media": len(media),
            "available_media": sum(item.get("availability") == "available" for item in media),
            "missing_media": sum(item.get("availability") == "missing" for item in media),
        }
    )
    filtered["statistics"] = statistics
    return filtered


def _filter_normalized_date(
    normalized: dict[str, Any], *, accounting_date: str, timezone_name: str
) -> dict[str, Any]:
    zone = ZoneInfo(timezone_name)
    start_at = datetime.strptime(accounting_date, "%Y-%m-%d").replace(tzinfo=zone)
    return _filter_normalized_window(
        normalized,
        start_at=start_at,
        end_at=start_at + timedelta(days=1),
        fingerprint_fields={
            "accounting_date": accounting_date,
            "timezone": timezone_name,
        },
    )


def _filter_normalized_time_range(
    normalized: dict[str, Any],
    *,
    accounting_from: str,
    accounting_to: str,
    timezone_name: str,
) -> dict[str, Any]:
    zone = ZoneInfo(timezone_name)
    return _filter_normalized_window(
        normalized,
        start_at=datetime.strptime(accounting_from, "%Y-%m-%d %H:%M").replace(tzinfo=zone),
        end_at=datetime.strptime(accounting_to, "%Y-%m-%d %H:%M").replace(tzinfo=zone),
        fingerprint_fields={
            "accounting_from": accounting_from,
            "accounting_to": accounting_to,
            "timezone": timezone_name,
        },
    )


def start_run(args: argparse.Namespace) -> dict[str, Any]:
    work = args.work.resolve()
    core.require(not work.exists(), f"work directory already exists; start requires a new path: {work}")
    inputs = [path.resolve() for path in args.inputs]
    for path in inputs:
        core.require(path.exists(), f"input does not exist: {path}")
    core.require(bool(core.clean_text(args.contains)), "--contains cannot be empty")
    zone = ZoneInfo(args.timezone)
    core.require(args.timezone == "Asia/Bangkok", "accounting timezone must be Asia/Bangkok")
    accounting_date = _normalize_accounting_date(getattr(args, "date", None))
    accounting_from = _normalize_accounting_time(
        getattr(args, "accounting_from", None), option="--from"
    )
    accounting_to = _normalize_accounting_time(
        getattr(args, "accounting_to", None), option="--to"
    )
    core.require(
        (accounting_from is None) == (accounting_to is None),
        "--from and --to must be provided together",
    )
    core.require(
        accounting_date is None or accounting_from is None,
        "--date cannot be combined with --from or --to",
    )
    if accounting_from is not None and accounting_to is not None:
        core.require(
            datetime.strptime(accounting_from, "%Y-%m-%d %H:%M")
            < datetime.strptime(accounting_to, "%Y-%m-%d %H:%M"),
            "--from must be earlier than --to",
        )
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
    line_android_backups = _discover_line_android_backups(
        inputs, getattr(args, "line_android_backup", [])
    )
    documents.extend(
        _line_android_documents(
            work,
            line_android_backups,
            contains=args.contains,
            timezone_name=args.timezone,
            roster_path=roster_path,
        )
    )
    normalized = _merge_documents(documents, contains=args.contains, timezone_name=args.timezone)
    if accounting_date is not None:
        normalized = _filter_normalized_date(
            normalized,
            accounting_date=accounting_date,
            timezone_name=args.timezone,
        )
    elif accounting_from is not None and accounting_to is not None:
        normalized = _filter_normalized_time_range(
            normalized,
            accounting_from=accounting_from,
            accounting_to=accounting_to,
            timezone_name=args.timezone,
        )
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
                "edit_mode": EDIT_CONTROL_MODE,
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
        "line_android_backups": [str(path) for path in line_android_backups],
        "groups": group_reports,
    }
    if accounting_date is not None:
        run["accounting_date"] = accounting_date
    elif accounting_from is not None and accounting_to is not None:
        run["accounting_from"] = accounting_from
        run["accounting_to"] = accounting_to
    core.atomic_json(_run_path(work), run)
    result = {
        "work": str(work),
        "groups": len(group_reports),
        "messages": normalized["statistics"]["messages"],
        "evidence_media": sum(item["evidence_media"] for item in group_reports),
        "selected_groups": group_reports,
    }
    if accounting_date is not None:
        result["accounting_date"] = accounting_date
    elif accounting_from is not None and accounting_to is not None:
        result["accounting_from"] = accounting_from
        result["accounting_to"] = accounting_to
    return result


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
            "reviewed_through": decision.get("reviewed_through"),
            "read_complete": decision.get("read_complete"),
            "media_decisions": decision.get("media_decisions"),
            "orders": decision.get("orders"),
            "open_orders": decision.get("open_orders"),
            "balance_links": decision.get("balance_links"),
            "settlement_allocations": decision.get("settlement_allocations"),
            "unknown_payee_reviewed_entry_ids": decision.get(
                "unknown_payee_reviewed_entry_ids"
            ),
        }
    )


def _legacy_semantic_fingerprint_3_1(decision: Mapping[str, Any]) -> str:
    return core.fingerprint_json(
        {
            "media_decisions": decision.get("media_decisions"),
            "orders": decision.get("orders"),
            "balance_links": decision.get("balance_links"),
            "settlement_allocations": decision.get("settlement_allocations"),
            "unknown_payee_reviewed_entry_ids": decision.get(
                "unknown_payee_reviewed_entry_ids"
            ),
        }
    )


def _new_edit_control(approved_fingerprint: str) -> dict[str, Any]:
    return {
        "mode": EDIT_CONTROL_MODE,
        "approved_semantic_fingerprint": approved_fingerprint,
        "batch_count": 0,
        "last_batch": None,
    }


def _controlled_editing(run_group: Mapping[str, Any]) -> bool:
    return core.clean_text(run_group.get("edit_mode")) == EDIT_CONTROL_MODE


def _require_controlled_semantics(
    run_group: Mapping[str, Any],
    decision: Mapping[str, Any],
) -> bool:
    if not _controlled_editing(run_group):
        return False
    control = decision.get("edit_control")
    core.require(
        isinstance(control, Mapping),
        "controlled decision is missing edit_control; use review apply-batch",
    )
    unknown = sorted(
        set(control)
        - {"mode", "approved_semantic_fingerprint", "batch_count", "last_batch"}
    )
    core.require(
        not unknown,
        f"edit_control has unsupported fields: {', '.join(unknown)}",
    )
    core.require(
        core.clean_text(control.get("mode")) == EDIT_CONTROL_MODE,
        "unsupported decision edit_control mode",
    )
    batch_count = control.get("batch_count")
    core.require(
        isinstance(batch_count, int) and batch_count >= 0,
        "edit_control.batch_count must be a non-negative integer",
    )
    last_batch = control.get("last_batch")
    if batch_count == 0:
        core.require(last_batch is None, "edit_control.last_batch must be empty before the first batch")
    else:
        core.require(isinstance(last_batch, Mapping), "edit_control.last_batch is required")
        unknown_last = sorted(
            set(last_batch)
            - {
                "id",
                "input_fingerprint",
                "base_fingerprint",
                "result_fingerprint",
                "applied_at",
            }
        )
        core.require(
            not unknown_last,
            f"edit_control.last_batch has unsupported fields: {', '.join(unknown_last)}",
        )
    current = _semantic_fingerprint(decision)
    approved = core.clean_text(control.get("approved_semantic_fingerprint"))
    core.require(
        approved == current,
        "decision semantic fields were edited outside review apply-batch; "
        "restore the last approved decision or submit a batch from its approved fingerprint",
    )
    if isinstance(last_batch, Mapping):
        core.require(
            core.clean_text(last_batch.get("result_fingerprint")) == current,
            "edit_control.last_batch does not match the approved decision",
        )
    return True


def _load_review_batch(path: Path) -> dict[str, Any]:
    core.require(path.is_file(), f"review batch does not exist: {path}")
    batch = _load_json(path)
    allowed = {
        "contract_version",
        "batch_id",
        "base_fingerprint",
        "media_decisions",
        "remove_media_labels",
        "orders",
        "remove_order_ids",
        "open_orders",
        "page_commit",
        "balance_links",
        "settlement_allocations",
        "unknown_payee_reviewed_entry_ids",
    }
    unknown = sorted(set(batch) - allowed)
    core.require(
        not unknown,
        f"review batch has unsupported fields: {', '.join(unknown)}",
    )
    core.require(
        batch.get("contract_version") == REVIEW_BATCH_CONTRACT,
        f"unsupported review batch contract; expected {REVIEW_BATCH_CONTRACT}",
    )
    batch_id = core.clean_text(batch.get("batch_id"))
    core.require(
        bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,127}", batch_id)),
        "review batch_id must use 1..128 ASCII letters, digits, dot, underscore, or hyphen",
    )
    batch["batch_id"] = batch_id
    base_fingerprint = core.clean_text(batch.get("base_fingerprint"))
    core.require(
        bool(re.fullmatch(r"sha256:[0-9a-f]{64}", base_fingerprint)),
        "review batch base_fingerprint must be a sha256 fingerprint from review status, next, or check",
    )
    batch["base_fingerprint"] = base_fingerprint
    page_commit = batch.get("page_commit")
    if page_commit is not None:
        core.require(isinstance(page_commit, Mapping), "review batch page_commit must be an object")
        unknown_page_fields = sorted(
            set(page_commit) - {"page_start", "page_end", "page_token"}
        )
        core.require(
            not unknown_page_fields,
            f"review batch page_commit has unsupported fields: {', '.join(unknown_page_fields)}",
        )
        page_start = page_commit.get("page_start")
        page_end = page_commit.get("page_end")
        core.require(
            isinstance(page_start, int) and isinstance(page_end, int),
            "review batch page_commit page_start and page_end must be integers",
        )
        page_token = core.clean_text(page_commit.get("page_token"))
        core.require(
            bool(re.fullmatch(r"sha256:[0-9a-f]{64}", page_token)),
            "review batch page_commit page_token must come from review next",
        )
        page_commit["page_token"] = page_token
        core.require(
            "open_orders" in batch,
            "review batch that commits a page must include the complete open_orders list",
        )
    if "open_orders" in batch:
        core.require(isinstance(batch["open_orders"], list), "review batch open_orders must be a list")
    return batch


def _merge_review_batch(
    decision: Mapping[str, Any],
    batch: Mapping[str, Any],
) -> dict[str, Any]:
    candidate = copy.deepcopy(dict(decision))
    candidate["sealed"] = False
    candidate["sealed_decision_fingerprint"] = None

    media_updates = batch.get("media_decisions", {})
    core.require(isinstance(media_updates, Mapping), "review batch media_decisions must be an object")
    remove_media_value = batch.get("remove_media_labels", [])
    core.require(isinstance(remove_media_value, list), "review batch remove_media_labels must be a list")
    remove_media = [str(item) for item in remove_media_value]
    core.require(len(remove_media) == len(set(remove_media)), "review batch remove_media_labels repeats a label")
    core.require(
        not (set(media_updates) & set(remove_media)),
        "review batch cannot update and remove the same media label",
    )
    candidate_media = candidate.get("media_decisions")
    core.require(isinstance(candidate_media, dict), "decision media_decisions must be an object")
    missing_media = sorted(set(remove_media) - set(candidate_media))
    core.require(not missing_media, f"review batch removes unknown media labels: {missing_media[:20]}")
    for label in remove_media:
        del candidate_media[label]
    for label, value in media_updates.items():
        candidate_media[str(label)] = copy.deepcopy(value)

    order_updates = batch.get("orders", [])
    core.require(isinstance(order_updates, list), "review batch orders must be a list")
    update_order_ids: list[str] = []
    for position, order in enumerate(order_updates):
        core.require(isinstance(order, Mapping), f"review batch orders[{position}] must be an object")
        order_id = core.clean_text(order.get("id"))
        core.require(bool(order_id), f"review batch orders[{position}].id is required")
        update_order_ids.append(order_id)
    core.require(
        len(update_order_ids) == len(set(update_order_ids)),
        "review batch orders repeats an order id",
    )
    remove_orders_value = batch.get("remove_order_ids", [])
    core.require(isinstance(remove_orders_value, list), "review batch remove_order_ids must be a list")
    remove_order_ids = [str(item) for item in remove_orders_value]
    core.require(
        len(remove_order_ids) == len(set(remove_order_ids)),
        "review batch remove_order_ids repeats an order id",
    )
    core.require(
        not (set(update_order_ids) & set(remove_order_ids)),
        "review batch cannot update and remove the same order id",
    )
    candidate_orders = candidate.get("orders")
    core.require(isinstance(candidate_orders, list), "decision orders must be a list")
    existing_order_ids = {
        core.clean_text(order.get("id"))
        for order in candidate_orders
        if isinstance(order, Mapping)
    }
    missing_orders = sorted(set(remove_order_ids) - existing_order_ids)
    core.require(not missing_orders, f"review batch removes unknown order ids: {missing_orders[:20]}")
    if remove_order_ids:
        candidate_orders[:] = [
            order
            for order in candidate_orders
            if not isinstance(order, Mapping)
            or core.clean_text(order.get("id")) not in set(remove_order_ids)
        ]
    order_positions = {
        core.clean_text(order.get("id")): position
        for position, order in enumerate(candidate_orders)
        if isinstance(order, Mapping)
    }
    for order in order_updates:
        order_id = core.clean_text(order.get("id"))
        replacement = copy.deepcopy(dict(order))
        if order_id in order_positions:
            candidate_orders[order_positions[order_id]] = replacement
        else:
            order_positions[order_id] = len(candidate_orders)
            candidate_orders.append(replacement)

    for field in (
        "open_orders",
        "balance_links",
        "settlement_allocations",
        "unknown_payee_reviewed_entry_ids",
    ):
        if field in batch:
            core.require(isinstance(batch[field], list), f"review batch {field} must be a list")
            candidate[field] = copy.deepcopy(batch[field])
    return candidate


def _validate_entry(
    entry: dict[str, Any],
    *,
    field: str,
    sender_role: str,
    source_message_label: str,
    message_labels: set[str],
) -> None:
    allowed = {
        "amount",
        "amount_text",
        "currency",
        "payee",
        "payee_state",
        "kind",
        "side",
        "side_exception",
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
    exception = entry.get("side_exception")
    side = core.clean_text(entry.get("side")).casefold()
    if not side:
        exception_kind_hint = (
            core.clean_text(exception.get("kind")).casefold()
            if isinstance(exception, Mapping)
            else ""
        )
        side = SIDE_EXCEPTION_SIDES.get(
            exception_kind_hint,
            {
                "内部人员": "payout",
                "客户候选": "payment",
            }.get(sender_role, "unknown"),
        )
    core.require(side in FLOW_SIDES, f"{field}.side is unsupported")
    exception_kind: str | None = None
    if exception not in (None, ""):
        core.require(isinstance(exception, dict), f"{field}.side_exception must be an object")
        unknown_exception_fields = sorted(
            set(exception) - {"kind", "source_messages", "detail"}
        )
        core.require(
            not unknown_exception_fields,
            f"{field}.side_exception: unsupported fields: {', '.join(unknown_exception_fields)}",
        )
        exception_kind = core.clean_text(exception.get("kind")).casefold()
        core.require(
            exception_kind in SIDE_EXCEPTION_SIDES,
            f"{field}.side_exception.kind is unsupported",
        )
        core.require(
            SIDE_EXCEPTION_SIDES[exception_kind] == side,
            f"{field}.side_exception.kind does not support side={side}",
        )
        raw_exception_sources = exception.get("source_messages")
        core.require(
            isinstance(raw_exception_sources, list) and raw_exception_sources,
            f"{field}.side_exception.source_messages is required",
        )
        exception_sources = [str(item) for item in raw_exception_sources]
        core.require(
            len(exception_sources) == len(set(exception_sources)),
            f"{field}.side_exception.source_messages repeats a label",
        )
        core.require(
            set(exception_sources) <= message_labels,
            f"{field}.side_exception.source_messages contains an unknown label",
        )
        core.require(
            any(label != source_message_label for label in exception_sources),
            f"{field}.side_exception must cite chat context beyond the fund image itself",
        )
        detail = core.clean_text(exception.get("detail"))
        core.require(bool(detail), f"{field}.side_exception.detail is required")
        exception["kind"] = exception_kind
        exception["source_messages"] = exception_sources
        exception["detail"] = detail

    if side in {"payment", "payout"}:
        expected_side = {
            "内部人员": "payout",
            "客户候选": "payment",
        }.get(sender_role)
        if expected_side is None:
            core.require(
                side == "unknown",
                f"{field}: unknown sender role ordinary fund entry must use side=unknown",
            )
        elif side != expected_side:
            required_exception = (
                "relayed_customer_payment"
                if sender_role == "内部人员"
                else "relayed_internal_payout"
            )
            role_name = "internal" if sender_role == "内部人员" else "customer"
            core.require(
                exception_kind == required_exception,
                f"{field}: {role_name} sender ordinary fund entry must use "
                f"side={expected_side} unless side_exception.kind={required_exception}",
            )
        else:
            core.require(
                exception_kind is None,
                f"{field}.side_exception is only allowed when ordinary side differs from sender role",
            )
    elif side == "payment_refund":
        core.require(
            exception_kind == "explicit_payment_refund",
            f"{field}: side=payment_refund requires side_exception.kind=explicit_payment_refund",
        )
    elif side == "recovery":
        core.require(
            exception_kind == "explicit_recovery",
            f"{field}: side=recovery requires side_exception.kind=explicit_recovery",
        )
    else:
        core.require(
            exception_kind is None,
            f"{field}.side_exception is not allowed when side=unknown",
        )
    entry["side"] = side
    status_text = core.clean_text(entry.get("status_text"))
    if status_text:
        entry["status_text"] = status_text
    normalized_status_text = status_text.casefold()
    failure_text = f"{status_text} {core.clean_text(entry.get('note'))}".casefold()
    result = core.clean_text(entry.get("result")).casefold()
    if not result:
        result = (
            "failed"
            if any(marker in failure_text for marker in FAILURE_STATUS_MARKERS)
            else "completed"
        )
    core.require(result in ENTRY_RESULTS, f"{field}.result is unsupported")
    if any(marker in normalized_status_text for marker in NORMAL_PROCESSING_STATUS_MARKERS):
        core.require(
            result == "completed",
            f"{field}: normal processing status must use result=completed",
        )
    if any(marker in normalized_status_text for marker in FAILURE_STATUS_MARKERS):
        core.require(
            result == "failed",
            f"{field}: explicit failure status must use result=failed",
        )
    if result == "failed":
        core.require(
            any(marker in failure_text for marker in FAILURE_STATUS_MARKERS),
            f"{field}: result=failed requires explicit failure, cancellation, rejection, invalidation, or risk-control text",
        )
    entry["result"] = result
    amount_state = core.clean_text(entry.get("amount_state")).casefold()
    if not amount_state:
        has_amount = entry.get("amount") not in (None, "")
        has_currency = entry.get("currency") not in (None, "")
        amount_state = "clear" if has_amount and has_currency else "partial" if has_amount or has_currency else "unreadable"
    core.require(amount_state in AMOUNT_STATES, f"{field}.amount_state is unsupported")
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
    if entry.get("payee_state") in (None, ""):
        payee_text = core.clean_text(entry.get("payee"))
        entry["payee_state"] = (
            "cash"
            if kind == "cash"
            else "not_shown"
            if payee_text == "未显示"
            else "unreadable"
            if payee_text == "无法辨认"
            else "visible"
        )
    if kind == "transfer" and entry.get("currency") == "THB":
        entry["payee"] = core.validate_thai_bank_account_payee(
            entry.get("payee"),
            field=f"{field}.payee",
            payee_state=entry.get("payee_state"),
        )
    else:
        entry["payee"] = core.validate_payee(
            entry.get("payee"),
            field=f"{field}.payee",
            cash=kind == "cash",
            payee_state=entry.get("payee_state"),
            require_state=True,
        )
    if entry.get("payee_state") not in (None, ""):
        entry["payee_state"] = core.clean_text(entry.get("payee_state")).casefold()


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


def _validate_pricing_scope(
    value: object,
    *,
    field: str,
    source_labels: set[str],
    payout_currency: str | None,
    require_complete: bool,
) -> bool:
    """Validate one v3 pricing authority. Return True only when it is missing."""

    if value in (None, ""):
        core.require(not require_complete, f"{field} is required before seal")
        return True
    core.require(isinstance(value, dict), f"{field} must be an object")
    pricing = value
    unknown = sorted(set(pricing) - {"source_messages", "terms", "expected"})
    core.require(not unknown, f"{field}: unsupported fields: {', '.join(unknown)}")

    raw_sources = pricing.get("source_messages")
    core.require(
        isinstance(raw_sources, list) and raw_sources,
        f"{field}.source_messages is required",
    )
    cited = [str(item) for item in raw_sources]
    core.require(
        len(cited) == len(set(cited)),
        f"{field}.source_messages repeats a label",
    )
    core.require(
        set(cited) <= source_labels,
        f"{field}.source_messages must cite labels already present in order.source_messages",
    )

    terms_value = pricing.get("terms")
    terms: dict[str, Any] | None = None
    if terms_value not in (None, ""):
        core.require(isinstance(terms_value, dict), f"{field}.terms must be an object")
        terms = terms_value
        unknown_terms = sorted(set(terms) - {"rate", "operator", "fees", "rounding"})
        core.require(
            not unknown_terms,
            f"{field}.terms: unsupported fields: {', '.join(unknown_terms)}",
        )
        rate = core.parse_decimal(terms.get("rate"), field=f"{field}.terms.rate")
        core.require(rate is not None and rate > 0, f"{field}.terms.rate must be positive")
        terms["rate"] = core.decimal_text(rate)
        operator = core.clean_text(terms.get("operator")).casefold()
        core.require(
            operator in simple_ledger.RATE_OPERATORS,
            f"{field}.terms.operator is required and must be multiply or divide",
        )
        terms["operator"] = operator
        _validate_network_fees(terms.get("fees"), field=f"{field}.terms.fees")
        rounding = terms.get("rounding")
        if rounding not in (None, ""):
            core.require(isinstance(rounding, dict), f"{field}.terms.rounding must be an object")
            unknown_rounding = sorted(set(rounding) - {"unit", "mode", "currency"})
            core.require(
                not unknown_rounding,
                f"{field}.terms.rounding: unsupported fields: {', '.join(unknown_rounding)}",
            )
            unit = core.parse_decimal(
                rounding.get("unit"),
                field=f"{field}.terms.rounding.unit",
            )
            core.require(unit is not None and unit > 0, f"{field}.terms.rounding.unit must be positive")
            rounding["unit"] = core.decimal_text(unit)
            mode = core.clean_text(rounding.get("mode")).casefold() or "half_up"
            core.require(
                mode in simple_ledger.ROUNDING_MODES,
                f"{field}.terms.rounding.mode is unsupported",
            )
            rounding["mode"] = mode
            if rounding.get("currency") not in (None, ""):
                rounding["currency"] = core.normalize_currency(
                    rounding.get("currency"),
                    field=f"{field}.terms.rounding.currency",
                )
                if payout_currency:
                    core.require(
                        rounding["currency"] == payout_currency,
                        f"{field}.terms.rounding.currency must use {payout_currency}",
                    )

    expected = pricing.get("expected")
    core.require(isinstance(expected, dict), f"{field}.expected is required")
    kind = core.clean_text(expected.get("kind")).casefold()
    core.require(kind in PRICING_EXPECTED_KINDS, f"{field}.expected.kind is unsupported")
    expected["kind"] = kind
    if kind == "explicit":
        unknown_expected = sorted(set(expected) - {"kind", "amount"})
        core.require(
            not unknown_expected,
            f"{field}.expected: unsupported fields: {', '.join(unknown_expected)}",
        )
        amount = core.parse_decimal(expected.get("amount"), field=f"{field}.expected.amount")
        core.require(amount is not None and amount >= 0, f"{field}.expected.amount cannot be negative")
        expected["amount"] = core.decimal_text(amount)
    elif kind == "calculated_from_terms":
        unknown_expected = sorted(set(expected) - {"kind"})
        core.require(
            not unknown_expected,
            f"{field}.expected: unsupported fields: {', '.join(unknown_expected)}",
        )
        core.require(terms is not None, f"{field}.terms is required for calculated_from_terms")
    else:
        unknown_expected = sorted(set(expected) - {"kind", "reason"})
        core.require(
            not unknown_expected,
            f"{field}.expected: unsupported fields: {', '.join(unknown_expected)}",
        )
        reason = core.clean_text(expected.get("reason")).casefold()
        core.require(
            reason in PRICING_UNKNOWN_REASONS,
            f"{field}.expected.reason is required and unsupported",
        )
        expected["reason"] = reason
    if require_complete and kind != "unknown":
        core.require(
            terms is not None,
            f"{field}.terms with the confirmed exchange rate is required before seal",
        )
    return False


def _validate_open_orders(
    group: Mapping[str, Any],
    decision: dict[str, Any],
    *,
    reviewed_through: int,
    require_complete: bool,
) -> int:
    open_orders = decision.get("open_orders")
    core.require(isinstance(open_orders, list), "open_orders must be a list")
    if require_complete:
        core.require(
            not open_orders,
            "open_orders must be resolved, cancelled, or converted to final orders before seal",
        )

    _, message_by_label = _message_labels(group)
    message_positions = {
        label: position
        for position, label in enumerate(message_by_label, start=1)
    }
    label_by_message_id, _ = _message_labels(group)
    inventory = _media_inventory(group)
    media_source_labels = {
        media_label: label_by_message_id.get(str(message.get("message_id") or ""), "")
        for media_label, (message, _) in inventory.items()
    }
    seen_ids: set[str] = set()
    allowed_fields = {
        "id",
        "start_message",
        "source_messages",
        "media_labels",
        "customer_nickname",
        "direction",
        "rate",
        "operator",
        "summary",
        "unresolved",
    }
    for position, item in enumerate(open_orders):
        field = f"open_orders[{position}]"
        core.require(isinstance(item, dict), f"{field} must be an object")
        unknown_fields = sorted(set(item) - allowed_fields)
        core.require(
            not unknown_fields,
            f"{field} has unsupported fields: {', '.join(unknown_fields)}",
        )
        open_id = core.clean_text(item.get("id"))
        core.require(bool(open_id), f"{field}.id is required")
        core.require(open_id not in seen_ids, f"open_orders repeats id {open_id}")
        seen_ids.add(open_id)
        item["id"] = open_id

        start_message = core.clean_text(item.get("start_message"))
        core.require(start_message in message_positions, f"{field}.start_message is unknown")
        source_value = item.get("source_messages")
        core.require(
            isinstance(source_value, list) and source_value,
            f"{field}.source_messages must be a non-empty list",
        )
        source_messages = [str(value) for value in source_value]
        core.require(
            len(source_messages) == len(set(source_messages)),
            f"{field}.source_messages repeats a label",
        )
        core.require(
            set(source_messages) <= set(message_positions),
            f"{field}.source_messages contains an unknown label",
        )
        core.require(
            start_message in source_messages,
            f"{field}.source_messages must include start_message",
        )
        core.require(
            all(message_positions[label] <= reviewed_through for label in source_messages),
            f"{field}.source_messages cannot cite an uncommitted page",
        )
        item["start_message"] = start_message
        item["source_messages"] = source_messages

        media_value = item.get("media_labels", [])
        core.require(isinstance(media_value, list), f"{field}.media_labels must be a list")
        media_labels = [str(value) for value in media_value]
        core.require(
            len(media_labels) == len(set(media_labels)),
            f"{field}.media_labels repeats a label",
        )
        core.require(
            set(media_labels) <= set(inventory),
            f"{field}.media_labels contains an unknown label",
        )
        core.require(
            all(media_source_labels[label] in source_messages for label in media_labels),
            f"{field}.source_messages must include every media source message",
        )
        item["media_labels"] = media_labels

        if "customer_nickname" in item:
            item["customer_nickname"] = core.clean_text(item.get("customer_nickname"))
        direction = core.clean_text(item.get("direction"))
        if direction:
            item["direction"] = core.canonical_direction(direction, field=f"{field}.direction")
        elif "direction" in item:
            item["direction"] = ""

        rate_text = core.clean_text(item.get("rate"))
        operator = core.clean_text(item.get("operator")).casefold()
        if rate_text:
            rate = core.parse_decimal(rate_text, field=f"{field}.rate")
            core.require(rate is not None and rate > 0, f"{field}.rate must be positive")
            core.require(operator in {"multiply", "divide"}, f"{field}.operator is required with rate")
            item["rate"] = core.decimal_text(rate)
            item["operator"] = operator
        else:
            core.require(not operator, f"{field}.operator requires rate")
            item.pop("rate", None)
            item.pop("operator", None)

        summary = core.clean_text(item.get("summary"))
        core.require(bool(summary), f"{field}.summary is required")
        item["summary"] = summary
        unresolved_value = item.get("unresolved")
        core.require(
            isinstance(unresolved_value, list) and unresolved_value,
            f"{field}.unresolved must explain why the order is still open",
        )
        unresolved = [core.clean_text(value) for value in unresolved_value]
        core.require(all(unresolved), f"{field}.unresolved cannot contain blank items")
        item["unresolved"] = unresolved
    return len(open_orders)


def _order_continuity_review_candidates(
    facts: list[dict[str, Any]],
    *,
    advanced_relation_order_ids: set[str],
) -> list[dict[str, Any]]:
    """Return narrow, non-mutating candidates for a possible split continuation."""

    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for fact in facts:
        customer_key = core.clean_text(fact.get("customer_key"))
        direction = core.clean_text(fact.get("direction"))
        if not customer_key or not direction:
            continue
        grouped.setdefault((customer_key, direction), []).append(fact)

    candidates: list[dict[str, Any]] = []
    for related_orders in grouped.values():
        related_orders.sort(
            key=lambda item: (
                int(item["source_start_position"]),
                int(item["source_end_position"]),
                int(item["order_position"]),
            )
        )
        for earlier, later in zip(related_orders, related_orders[1:]):
            earlier_order_id = str(earlier["order_id"])
            later_order_id = str(later["order_id"])
            if (
                earlier_order_id in advanced_relation_order_ids
                or later_order_id in advanced_relation_order_ids
                or earlier["has_complex_flow"]
                or later["has_complex_flow"]
            ):
                continue
            if not (
                earlier["has_completed_payment"]
                and not earlier["has_completed_payout"]
                and earlier["has_confirmed_pricing"]
                and later["has_completed_payment"]
                and later["has_completed_payout"]
                and later["has_unknown_pricing"]
            ):
                continue
            candidates.append(
                {
                    "earlier_order_id": earlier_order_id,
                    "later_order_id": later_order_id,
                    "customer_nickname": earlier["customer_nickname"],
                    "direction": earlier["direction"],
                    "source_start": f"S{min(int(earlier['source_start_position']), int(later['source_start_position'])):05d}",
                    "source_end": f"S{max(int(earlier['source_end_position']), int(later['source_end_position'])):05d}",
                    "reason_codes": [
                        "earlier_payment_without_payout",
                        "earlier_confirmed_pricing",
                        "later_payment_and_payout",
                        "later_unknown_pricing",
                    ],
                }
            )
    return candidates


def _validate_decision(
    normalized: Mapping[str, Any],
    group: Mapping[str, Any],
    decision: dict[str, Any],
    *,
    require_complete: bool,
    capture_hashes: bool,
) -> dict[str, Any]:
    expected = _decision_template(normalized, group)
    decision_contract = core.clean_text(decision.get("contract_version"))
    core.require(
        decision_contract == DECISION_CONTRACT,
        f"unsupported decision contract; expected {DECISION_CONTRACT}",
    )
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
    core.require(
        decision.get("read_complete") is (reviewed_through == len(messages)),
        "read_complete must exactly match reviewed_through",
    )
    if require_complete:
        core.require(decision.get("read_complete") is True and reviewed_through == len(messages), "group chronology has not been read completely")
    open_order_count = _validate_open_orders(
        group,
        decision,
        reviewed_through=reviewed_through,
        require_complete=require_complete,
    )

    inventory = _media_inventory(group)
    label_by_message_id, message_by_label = _message_labels(group)
    message_labels = set(message_by_label)
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
    entry_records: dict[str, dict[str, Any]] = {}
    entry_source_labels: dict[str, str] = {}
    transfer_payees = 0
    unknown_payee_entry_ids: set[str] = set()
    unreadable_payee_entry_ids: set[str] = set()
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
        message, media = inventory[label]
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
            source_message_label = label_by_message_id[
                str(message.get("message_id") or "")
            ]
            _validate_entry(
                entry,
                field=f"{field}.entries[{position - 1}]",
                sender_role=core.clean_text(message.get("role")),
                source_message_label=source_message_label,
                message_labels=message_labels,
            )
            entry_id = f"{label}.{position}"
            core.require(entry_id not in entry_ids, f"duplicate entry id: {entry_id}")
            entry_ids.add(entry_id)
            entry_records[entry_id] = entry
            entry_source_labels[entry_id] = source_message_label
            if entry["kind"] == "transfer":
                transfer_payees += 1
                if entry["payee"] in core.UNKNOWN_PAYEES:
                    unknown_payee_entry_ids.add(entry_id)
                if entry.get("payee_state") == "unreadable":
                    unreadable_payee_entry_ids.add(entry_id)
        fund_media_count += 1

    reviewed_unknown_value = decision.get("unknown_payee_reviewed_entry_ids", [])
    core.require(
        isinstance(reviewed_unknown_value, list),
        "unknown_payee_reviewed_entry_ids must be a list",
    )
    reviewed_unknown = [str(item) for item in reviewed_unknown_value]
    core.require(
        len(reviewed_unknown) == len(set(reviewed_unknown)),
        "unknown_payee_reviewed_entry_ids repeats an entry",
    )
    core.require(
        set(reviewed_unknown) <= unknown_payee_entry_ids,
        "unknown_payee_reviewed_entry_ids may only cite transfer entries marked 未显示 or 无法辨认",
    )
    unknown_payee_warning = (
        transfer_payees >= UNKNOWN_PAYEE_WARNING_MIN_TRANSFERS
        and len(unknown_payee_entry_ids) * 2 > transfer_payees
    )
    unknown_payee_review_required = bool(unreadable_payee_entry_ids)
    unreviewed_unknown_payees = unreadable_payee_entry_ids - set(reviewed_unknown)
    if require_complete and unreviewed_unknown_payees:
        core.require(
            not unreviewed_unknown_payees,
            "unreadable payee review is required: reopen only the cited unclear originals and list "
            "the still-unreadable entry IDs in unknown_payee_reviewed_entry_ids; remaining: "
            f"{sorted(unreviewed_unknown_payees)[:40]}",
        )

    del label_by_message_id
    orders = decision.get("orders")
    core.require(isinstance(orders, list), "orders must be a list")
    order_ids: set[str] = set()
    order_by_id: dict[str, dict[str, Any]] = {}
    assigned_entries: set[str] = set()
    pricing_scopes = 0
    missing_pricing_scopes: list[str] = []
    single_entry_orders = 0
    blank_direction_orders = 0
    unknown_pricing_orders = 0
    own_fund_source_only_orders = 0
    mass_degenerate_orders = 0
    single_unknown_pricing_orders = 0
    order_degradation_signals: list[tuple[bool, bool, bool, bool, str]] = []
    order_continuity_facts: list[dict[str, Any]] = []
    order_note_counts: dict[str, int] = {}
    order_fields = {
        "id",
        "entry_ids",
        "source_messages",
        "customer_nickname",
        "direction",
        "pricing",
        "note",
        "same_transactions",
        "legs",
    }
    for position, order in enumerate(orders):
        field = f"{group.get('group_key')}.orders[{position}]"
        core.require(isinstance(order, dict), f"{field} must be an object")
        unknown = sorted(set(order) - order_fields)
        core.require(not unknown, f"{field}: unsupported fields: {', '.join(unknown)}")
        for required_field in ("customer_nickname", "direction"):
            core.require(
                required_field in order,
                f"{field}.{required_field} must be explicitly supplied by the model",
            )
        order_note = core.clean_text(order.get("note"))
        if order_note:
            order["note"] = order_note
            order_note_counts[order_note] = order_note_counts.get(order_note, 0) + 1
        order_id = core.clean_text(order.get("id"))
        core.require(bool(order_id), f"{field}.id is required")
        core.require(order_id not in order_ids, f"{field}.id is duplicated")
        order_ids.add(order_id)
        order_by_id[order_id] = order
        raw_entries = order.get("entry_ids")
        core.require(isinstance(raw_entries, list) and raw_entries, f"{field}.entry_ids is required")
        refs = [str(item) for item in raw_entries]
        is_single_entry_order = len(refs) == 1
        if is_single_entry_order:
            single_entry_orders += 1
        core.require(len(refs) == len(set(refs)), f"{field}.entry_ids repeats an entry")
        for entry_id in refs:
            core.require(entry_id in entry_ids, f"{field}: unknown fund entry {entry_id}")
            core.require(entry_id not in assigned_entries, f"fund entry assigned to multiple orders: {entry_id}")
            assigned_entries.add(entry_id)
        raw_messages = order.get("source_messages", [])
        core.require(isinstance(raw_messages, list), f"{field}.source_messages must be a list")
        source_labels = [str(item) for item in raw_messages]
        core.require(len(source_labels) == len(set(source_labels)), f"{field}.source_messages repeats a label")
        derived_source_labels = [entry_source_labels[entry_id] for entry_id in refs]
        for entry_id in refs:
            side_exception = entry_records[entry_id].get("side_exception")
            if isinstance(side_exception, Mapping):
                derived_source_labels.extend(
                    str(item) for item in side_exception.get("source_messages", [])
                )
        pricing_candidates: list[object] = [order.get("pricing")]
        if isinstance(order.get("legs"), list):
            pricing_candidates.extend(
                leg.get("pricing")
                for leg in order["legs"]
                if isinstance(leg, Mapping)
            )
        for pricing_candidate in pricing_candidates:
            if isinstance(pricing_candidate, Mapping):
                derived_source_labels.extend(
                    str(item)
                    for item in pricing_candidate.get("source_messages", [])
                )
        for label in derived_source_labels:
            if label not in source_labels:
                source_labels.append(label)
        order["source_messages"] = source_labels
        core.require(set(source_labels) <= set(message_by_label), f"{field}.source_messages contains an unknown label")
        source_label_set = set(source_labels)
        own_fund_source_only = source_label_set == {
            entry_source_labels[entry_id] for entry_id in refs
        }
        if own_fund_source_only:
            own_fund_source_only_orders += 1
        for entry_id in refs:
            side_exception = entry_records[entry_id].get("side_exception")
            if isinstance(side_exception, Mapping):
                exception_sources = {
                    str(item) for item in side_exception.get("source_messages", [])
                }
                core.require(
                    exception_sources <= source_label_set,
                    f"{field}: side_exception.source_messages must belong to the order",
                )
        order_payout_currency: str | None = None
        blank_direction = not core.clean_text(order.get("direction"))
        if blank_direction:
            blank_direction_orders += 1
        else:
            order["direction"] = core.canonical_direction(order.get("direction"), field=f"{field}.direction")
            _, order_payout_currency = core.direction_currencies(
                order["direction"],
                field=f"{field}.direction",
            )
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
            core.require(
                order.get("pricing") in (None, ""),
                f"{field}.pricing must be empty when pricing is defined per leg",
            )
            referenced_settlements: set[str] = set()
            display_labels_by_direction: dict[str, list[str]] = {}
            for leg_position, leg in enumerate(legs):
                leg_field = f"{field}.legs[{leg_position}]"
                core.require(isinstance(leg, dict), f"{leg_field} must be an object")
                allowed_leg_fields = {
                    "leg_id",
                    "display_label",
                    "direction",
                    "allocation_amount",
                    "pricing",
                    "payout_entry_ids",
                    "recovery_entry_ids",
                }
                unknown_leg_fields = sorted(set(leg) - allowed_leg_fields)
                core.require(not unknown_leg_fields, f"{leg_field}: unsupported fields: {', '.join(unknown_leg_fields)}")
                leg_direction = core.canonical_direction(leg.get("direction"), field=f"{leg_field}.direction")
                _, leg_payout_currency = core.direction_currencies(
                    leg_direction,
                    field=f"{leg_field}.direction",
                )
                if isinstance(leg, dict):
                    leg["direction"] = leg_direction
                display_label = core.clean_text(leg.get("display_label"))
                if "display_label" in leg:
                    core.require(bool(display_label), f"{leg_field}.display_label cannot be blank")
                    core.require(
                        len(display_label) <= 80 and "\n" not in display_label and "\r" not in display_label,
                        f"{leg_field}.display_label must be a single line of at most 80 characters",
                    )
                    leg["display_label"] = display_label
                display_labels_by_direction.setdefault(leg_direction, []).append(display_label)
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
                if _validate_pricing_scope(
                    leg.get("pricing"),
                    field=f"{leg_field}.pricing",
                    source_labels=set(source_labels),
                    payout_currency=leg_payout_currency,
                    require_complete=require_complete,
                ):
                    missing_pricing_scopes.append(f"{leg_field}.pricing")
            for repeated_direction, display_labels in display_labels_by_direction.items():
                if len(display_labels) < 2:
                    continue
                core.require(
                    all(display_labels),
                    f"{field}.legs sharing direction {repeated_direction} require display_label on every leg",
                )
                core.require(
                    len({label.casefold() for label in display_labels}) == len(display_labels),
                    f"{field}.legs sharing direction {repeated_direction} require distinct display_label values",
                )
        else:
            pricing_scopes += 1
            if _validate_pricing_scope(
                order.get("pricing"),
                field=f"{field}.pricing",
                source_labels=set(source_labels),
                payout_currency=order_payout_currency,
                require_complete=require_complete,
            ):
                missing_pricing_scopes.append(f"{field}.pricing")

        if isinstance(legs, list) and legs:
            unknown_pricing = all(
                isinstance(leg.get("pricing"), Mapping)
                and isinstance(leg["pricing"].get("expected"), Mapping)
                and core.clean_text(leg["pricing"]["expected"].get("kind")).casefold()
                == "unknown"
                for leg in legs
            )
        else:
            pricing = order.get("pricing")
            unknown_pricing = (
                isinstance(pricing, Mapping)
                and isinstance(pricing.get("expected"), Mapping)
                and core.clean_text(pricing["expected"].get("kind")).casefold()
                == "unknown"
            )
        if unknown_pricing:
            unknown_pricing_orders += 1
        if is_single_entry_order and unknown_pricing:
            single_unknown_pricing_orders += 1
        meaningful_chat_context = any(
            bool(core.clean_text(message_by_label[label].get("text")))
            for label in source_label_set
        )
        if (
            own_fund_source_only
            and not meaningful_chat_context
            and (blank_direction or unknown_pricing)
        ):
            core.require(
                bool(order_note)
                and any(marker in order_note for marker in MISSING_CONTEXT_NOTE_MARKERS),
                f"{field}: isolated fund evidence requires a specific missing-context note",
            )
        completed_ordinary_sides = {
            core.clean_text(entry_records[entry_id].get("side")).casefold()
            for entry_id in refs
            if entry_records[entry_id].get("result") == "completed"
        }
        if (
            not blank_direction
            and not unknown_pricing
            and {"payment", "payout"} <= completed_ordinary_sides
        ):
            core.require(
                meaningful_chat_context,
                f"{field}: completed ordinary order requires meaningful chat context, not fund images alone",
            )
        source_positions = [
            int(match.group(1))
            for label in source_labels
            if (match := re.fullmatch(r"S(\d+)", label)) is not None
        ]
        pricing_value = order.get("pricing")
        has_confirmed_pricing = (
            isinstance(pricing_value, Mapping)
            and isinstance(pricing_value.get("terms"), Mapping)
            and isinstance(pricing_value.get("expected"), Mapping)
            and core.clean_text(pricing_value["expected"].get("kind")).casefold()
            in {"explicit", "calculated_from_terms"}
        )
        entry_sides = {
            core.clean_text(entry_records[entry_id].get("side")).casefold()
            for entry_id in refs
        }
        if source_positions:
            order_continuity_facts.append(
                {
                    "order_id": order_id,
                    "order_position": position,
                    "customer_nickname": core.clean_text(order.get("customer_nickname")),
                    "customer_key": core.clean_text(order.get("customer_nickname")).casefold(),
                    "direction": core.clean_text(order.get("direction")),
                    "source_start_position": min(source_positions),
                    "source_end_position": max(source_positions),
                    "has_completed_payment": "payment" in completed_ordinary_sides,
                    "has_completed_payout": "payout" in completed_ordinary_sides,
                    "has_confirmed_pricing": has_confirmed_pricing,
                    "has_unknown_pricing": unknown_pricing,
                    "has_complex_flow": bool(legs)
                    or bool(same_transactions)
                    or bool(entry_sides & {"payment_refund", "recovery"}),
                }
            )
        if (
            is_single_entry_order
            and blank_direction
            and unknown_pricing
            and own_fund_source_only
        ):
            mass_degenerate_orders += 1
        order_degradation_signals.append(
            (
                is_single_entry_order,
                blank_direction,
                unknown_pricing,
                own_fund_source_only,
                order_note,
            )
        )

    order_count = len(orders)
    repeated_note_values = {
        note for note, count in order_note_counts.items() if count > 1
    }
    repeated_note_orders = sum(
        bool(note) and note in repeated_note_values
        for *_, note in order_degradation_signals
    )
    multi_signal_degenerate_orders = sum(
        sum((single, blank, unknown_pricing, own_source, note in repeated_note_values))
        >= 3
        for single, blank, unknown_pricing, own_source, note in order_degradation_signals
    )
    mass_degenerate_order_ratio = (
        mass_degenerate_orders / order_count if order_count else 0.0
    )
    multi_signal_degenerate_order_ratio = (
        multi_signal_degenerate_orders / order_count if order_count else 0.0
    )
    suspected_bulk_order_creation = int(
        order_count >= 10
        and (
            mass_degenerate_order_ratio >= 0.8
            or multi_signal_degenerate_order_ratio >= 0.7
        )
    )

    advanced_relation_order_ids: set[str] = set()
    settlement_allocations = decision.get("settlement_allocations", [])
    core.require(
        isinstance(settlement_allocations, list),
        "settlement_allocations must be a list",
    )
    allocated_source_entries: set[str] = set()
    for position, relation in enumerate(settlement_allocations):
        field = f"{group.get('group_key')}.settlement_allocations[{position}]"
        core.require(isinstance(relation, dict), f"{field} must be an object")
        unknown_relation_fields = sorted(
            set(relation) - {"entry_id", "allocations", "source_messages"}
        )
        core.require(
            not unknown_relation_fields,
            f"{field}: unsupported fields: {', '.join(unknown_relation_fields)}",
        )
        source_entry_id = core.clean_text(relation.get("entry_id"))
        core.require(source_entry_id in entry_records, f"{field}.entry_id is unknown")
        core.require(
            source_entry_id not in assigned_entries,
            f"{field}.entry_id must not also appear in an order.entry_ids",
        )
        core.require(
            source_entry_id not in allocated_source_entries,
            f"{field}.entry_id is already used by another settlement allocation",
        )
        source_entry = entry_records[source_entry_id]
        core.require(
            source_entry.get("side") in {"payout", "recovery"},
            f"{field}.entry_id must identify a payout or recovery entry",
        )
        core.require(
            source_entry.get("result") == "completed"
            and source_entry.get("amount_state") == "clear",
            f"{field}.entry_id must be a completed entry with a clear amount",
        )
        source_amount = core.parse_decimal(
            source_entry.get("amount"),
            field=f"{field}.source_amount",
        )
        core.require(
            source_amount is not None and source_amount > 0,
            f"{field}.source_amount must be positive",
        )
        source_currency = core.normalize_currency(
            source_entry.get("currency"),
            field=f"{field}.source_currency",
        )
        allocations = relation.get("allocations")
        core.require(
            isinstance(allocations, list) and len(allocations) >= 2,
            f"{field}.allocations must contain at least two target orders",
        )
        target_order_ids: set[str] = set()
        target_customer_nicknames: set[str] = set()
        allocated_amounts = []
        for allocation_position, allocation in enumerate(allocations):
            allocation_field = f"{field}.allocations[{allocation_position}]"
            core.require(
                isinstance(allocation, dict),
                f"{allocation_field} must be an object",
            )
            unknown_allocation_fields = sorted(set(allocation) - {"order_id", "amount"})
            core.require(
                not unknown_allocation_fields,
                f"{allocation_field}: unsupported fields: {', '.join(unknown_allocation_fields)}",
            )
            target_order_id = core.clean_text(allocation.get("order_id"))
            core.require(
                target_order_id in order_by_id,
                f"{allocation_field}.order_id is unknown",
            )
            core.require(
                target_order_id not in target_order_ids,
                f"{field}.allocations repeats target order {target_order_id}",
            )
            target_order_ids.add(target_order_id)
            target_order = order_by_id[target_order_id]
            core.require(
                target_order.get("legs") in (None, "", []),
                f"{allocation_field}: settlement allocation does not support multi-leg target orders",
            )
            target_customer_nickname = core.clean_text(target_order.get("customer_nickname"))
            core.require(
                bool(target_customer_nickname),
                f"{allocation_field}: target customer_nickname must be known",
            )
            target_customer_nicknames.add(target_customer_nickname.casefold())
            _, target_payout_currency = core.direction_currencies(
                target_order.get("direction"),
                field=f"{allocation_field}.target_direction",
            )
            core.require(
                target_payout_currency == source_currency,
                f"{allocation_field}: target payout currency must be {source_currency}",
            )
            target_entries = [
                entry_records[str(item)]
                for item in target_order.get("entry_ids", [])
                if str(item) in entry_records
            ]
            core.require(
                any(item.get("side") in {"payment", "payment_refund"} for item in target_entries),
                f"{allocation_field}: target order must contain its own payment evidence",
            )
            amount = core.parse_decimal(
                allocation.get("amount"),
                field=f"{allocation_field}.amount",
            )
            core.require(
                amount is not None and amount > 0,
                f"{allocation_field}.amount must be positive",
            )
            allocation["order_id"] = target_order_id
            allocation["amount"] = core.decimal_text(amount)
            allocated_amounts.append(amount)
        core.require(
            len(target_customer_nicknames) == 1,
            f"{field}.allocations must target orders for the same customer",
        )
        advanced_relation_order_ids.update(target_order_ids)
        allocated_total = sum(allocated_amounts, start=source_amount * 0)
        core.require(
            allocated_total == source_amount,
            f"{field}.allocations total {core.decimal_text(allocated_total)} "
            f"must equal source amount {core.decimal_text(source_amount)}",
        )
        source_messages = relation.get("source_messages", [])
        core.require(isinstance(source_messages, list), f"{field}.source_messages must be a list")
        source_labels = [str(item) for item in source_messages]
        core.require(
            len(source_labels) == len(set(source_labels)),
            f"{field}.source_messages repeats a label",
        )
        core.require(
            set(source_labels) <= set(message_by_label),
            f"{field}.source_messages contains an unknown label",
        )
        relation["entry_id"] = source_entry_id
        allocated_source_entries.add(source_entry_id)
        assigned_entries.add(source_entry_id)

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
        advanced_relation_order_ids.update({source_order_id, target_order_id})
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

    order_continuity_review_candidates = _order_continuity_review_candidates(
        order_continuity_facts,
        advanced_relation_order_ids=advanced_relation_order_ids,
    )

    return {
        "available_media": len(available),
        "missing_media": len(missing),
        "classified_media": len(media_decisions),
        "reference_media": reference_count,
        "fund_media": fund_media_count,
        "fund_entries": len(entry_ids),
        "orders": len(orders),
        "open_orders": open_order_count,
        "single_entry_orders": single_entry_orders,
        "single_entry_order_ratio": single_entry_orders / order_count if order_count else 0.0,
        "blank_direction_orders": blank_direction_orders,
        "blank_direction_order_ratio": blank_direction_orders / order_count if order_count else 0.0,
        "unknown_pricing_orders": unknown_pricing_orders,
        "unknown_pricing_order_ratio": unknown_pricing_orders / order_count if order_count else 0.0,
        "own_fund_source_only_orders": own_fund_source_only_orders,
        "own_fund_source_only_order_ratio": own_fund_source_only_orders / order_count if order_count else 0.0,
        "repeated_order_notes": sum(
            count - 1 for count in order_note_counts.values() if count > 1
        ),
        "repeated_note_orders": repeated_note_orders,
        "repeated_note_order_ratio": repeated_note_orders / order_count if order_count else 0.0,
        "role_side_conflicts": 0,
        "mass_degenerate_orders": mass_degenerate_orders,
        "mass_degenerate_order_ratio": mass_degenerate_order_ratio,
        "single_unknown_pricing_orders": single_unknown_pricing_orders,
        "single_unknown_pricing_order_ratio": (
            single_unknown_pricing_orders / order_count if order_count else 0.0
        ),
        "multi_signal_degenerate_orders": multi_signal_degenerate_orders,
        "multi_signal_degenerate_order_ratio": multi_signal_degenerate_order_ratio,
        "suspected_bulk_order_creation": suspected_bulk_order_creation,
        "order_continuity_review_candidate_count": len(
            order_continuity_review_candidates
        ),
        "order_continuity_review_candidates": order_continuity_review_candidates,
        "order_to_fund_entry_ratio": order_count / len(entry_ids) if entry_ids else 0.0,
        "pricing_scopes": pricing_scopes,
        "missing_pricing_scopes": len(missing_pricing_scopes),
        "unassigned_entries": len(entry_ids - assigned_entries),
        "settlement_allocations": len(settlement_allocations),
        "transfer_payees": transfer_payees,
        "unknown_payees": len(unknown_payee_entry_ids),
        "unknown_payee_warning": int(unknown_payee_warning),
        "unknown_payee_review_required": int(unknown_payee_review_required),
        "unreviewed_unknown_payees": len(unreviewed_unknown_payees),
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


def _review_page_token(
    group: Mapping[str, Any],
    start: int,
    end: int,
    compact: list[dict[str, Any]] | None = None,
    collapsed: int | None = None,
    group_fingerprint: str | None = None,
) -> str:
    if compact is None or collapsed is None:
        compact, collapsed = _compact_page(group, start, end)
    return core.fingerprint_json(
        {
            "contract_version": REVIEW_PAGE_CONTRACT,
            "group_fingerprint": group_fingerprint or core.fingerprint_json(group),
            "page_start": start,
            "page_end": end,
            "messages": compact,
            "collapsed_system_or_empty_messages": collapsed,
        }
    )


def _commit_review_page(
    decision: dict[str, Any],
    batch: Mapping[str, Any],
    group: Mapping[str, Any],
) -> tuple[int, int] | None:
    page_commit = batch.get("page_commit")
    if page_commit is None:
        return None
    assert isinstance(page_commit, Mapping)
    start = page_commit["page_start"]
    end = page_commit["page_end"]
    messages = list(group.get("messages", []))
    committed_through = int(decision.get("reviewed_through") or 0)
    core.require(
        start == committed_through,
        f"page_commit must start at reviewed_through={committed_through}; fetch review next again",
    )
    core.require(start < end <= len(messages), "page_commit range is empty or outside the group")
    core.require(
        end - start <= MAX_PAGE_SIZE,
        f"page_commit cannot exceed {MAX_PAGE_SIZE} messages",
    )
    expected_token = _review_page_token(group, start, end)
    core.require(
        page_commit["page_token"] == expected_token,
        "page_commit token does not match the exact page returned by review next",
    )
    decision["reviewed_through"] = end
    decision["read_complete"] = end == len(messages)
    return start, end


def _open_order_carry_messages(
    group: Mapping[str, Any],
    open_orders: object,
) -> list[dict[str, Any]]:
    if not isinstance(open_orders, list):
        return []
    labels: set[str] = set()
    for item in open_orders:
        if not isinstance(item, Mapping):
            continue
        source_messages = item.get("source_messages", [])
        if isinstance(source_messages, list):
            labels.update(str(value) for value in source_messages)
    positions = sorted(
        int(match.group(1)) - 1
        for label in labels
        if (match := re.fullmatch(r"S(\d{5})", label)) is not None
    )
    result: list[dict[str, Any]] = []
    for position in positions:
        if 0 <= position < len(group.get("messages", [])):
            compact, _ = _compact_page(group, position, position + 1)
            result.extend(compact)
    return result


def _render_result_json(result: Mapping[str, Any], *, compact: bool) -> str:
    if compact:
        return json.dumps(result, ensure_ascii=False, separators=(",", ":")) + "\n"
    return json.dumps(result, ensure_ascii=False, indent=2) + "\n"


def _review_page_result(
    *,
    group: Mapping[str, Any],
    group_key: str,
    run_group: Mapping[str, Any],
    decision: Mapping[str, Any],
    start: int,
    end: int,
    group_fingerprint: str,
    semantic_fingerprint: str,
    open_orders: list[Any],
    carry_messages: list[dict[str, Any]],
) -> dict[str, Any]:
    messages = list(group.get("messages", []))
    compact, collapsed = _compact_page(group, start, end)
    page_token = (
        _review_page_token(
            group,
            start,
            end,
            compact,
            collapsed,
            group_fingerprint,
        )
        if start < end
        else None
    )
    return {
        "contract_version": REVIEW_PAGE_CONTRACT,
        "group_key": group_key,
        "run_label": run_group.get("run_label"),
        "page_start": start,
        "page_end": end,
        "message_count": len(messages),
        "messages": compact,
        "collapsed_system_or_empty_messages": collapsed,
        "returned_messages": len(compact),
        "page_reaches_group_end": end == len(messages),
        "done": bool(decision.get("read_complete")),
        "commit_required": start < end,
        "open_orders": open_orders,
        "carry_messages": carry_messages,
        "controlled_editing": _controlled_editing(run_group),
        "page_token": page_token,
        "semantic_fingerprint": semantic_fingerprint,
    }


RISK_DIAGNOSTIC_FIELDS = (
    "blank_direction_orders",
    "unknown_pricing_orders",
    "single_unknown_pricing_orders",
    "own_fund_source_only_orders",
    "repeated_note_orders",
    "multi_signal_degenerate_orders",
    "suspected_bulk_order_creation",
    "order_continuity_review_candidate_count",
    "unassigned_entries",
    "unknown_payees",
    "unknown_payee_warning",
)
RISK_FLAG_FIELDS = (
    "blank_direction_orders",
    "unknown_pricing_orders",
    "repeated_note_orders",
    "multi_signal_degenerate_orders",
    "suspected_bulk_order_creation",
    "order_continuity_review_candidate_count",
    "unassigned_entries",
    "unknown_payee_warning",
)


def _review_risk_summary(reports: list[dict[str, Any]]) -> dict[str, Any]:
    totals = {
        field: sum(int(report.get(field) or 0) for report in reports)
        for field in RISK_DIAGNOSTIC_FIELDS
    }
    flagged_groups = sum(
        any(int(report.get(field) or 0) for field in RISK_FLAG_FIELDS)
        for report in reports
    )
    return {
        "groups": len(reports),
        "flagged_groups": flagged_groups,
        "totals": totals,
    }


def _upgrade_legacy_checkpoints(
    work: Path,
    run: dict[str, Any],
    run_group: Mapping[str, Any],
    normalized: Mapping[str, Any],
    group: Mapping[str, Any],
    decision_path: Path,
    decision: Mapping[str, Any],
) -> dict[str, Any]:
    core.require(
        decision.get("contract_version") == LEGACY_DECISION_CONTRACT,
        f"upgrade-checkpoints only accepts {LEGACY_DECISION_CONTRACT}",
    )
    core.require(
        core.clean_text(run_group.get("edit_mode")) == LEGACY_EDIT_CONTROL_MODE,
        "legacy task is not using the controlled 3.1 edit protocol",
    )
    control = decision.get("edit_control")
    core.require(isinstance(control, Mapping), "legacy decision is missing edit_control")
    core.require(
        core.clean_text(control.get("mode")) == LEGACY_EDIT_CONTROL_MODE,
        "legacy decision has an unsupported edit_control mode",
    )
    legacy_fingerprint = _legacy_semantic_fingerprint_3_1(decision)
    core.require(
        core.clean_text(control.get("approved_semantic_fingerprint"))
        == legacy_fingerprint,
        "legacy decision semantic fields changed outside its approved batches",
    )

    previous_reviewed_through = int(decision.get("reviewed_through") or 0)
    candidate = copy.deepcopy(dict(decision))
    candidate["contract_version"] = DECISION_CONTRACT
    candidate["reviewed_through"] = 0
    candidate["read_complete"] = not bool(group.get("messages", []))
    candidate["open_orders"] = []
    candidate["sealed"] = False
    candidate["sealed_decision_fingerprint"] = None
    statistics = _validate_decision(
        normalized,
        group,
        candidate,
        require_complete=False,
        capture_hashes=False,
    )
    current_fingerprint = _semantic_fingerprint(candidate)
    candidate["edit_control"] = _new_edit_control(current_fingerprint)
    core.atomic_json(decision_path, candidate)
    run_group["edit_mode"] = EDIT_CONTROL_MODE
    run_group["sealed"] = False
    run["status"] = "reviewing"
    core.atomic_json(_run_path(work), run)
    return {
        "group_key": run_group.get("group_key"),
        "run_label": run_group.get("run_label"),
        "upgraded_from": LEGACY_DECISION_CONTRACT,
        "contract_version": DECISION_CONTRACT,
        "previous_reviewed_through": previous_reviewed_through,
        "reviewed_through": candidate["reviewed_through"],
        "preserved_media_decisions": len(candidate.get("media_decisions", {})),
        "preserved_orders": len(candidate.get("orders", [])),
        "semantic_fingerprint": current_fingerprint,
        **statistics,
    }


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
            semantic_fingerprint = _semantic_fingerprint(decision)
            control = decision.get("edit_control")
            approved_fingerprint = (
                core.clean_text(control.get("approved_semantic_fingerprint"))
                if isinstance(control, Mapping)
                else ""
            )
            reports.append(
                {
                    "group_key": run_group.get("group_key"),
                    "run_label": run_group.get("run_label"),
                    "reviewed_through": decision.get("reviewed_through"),
                    "read_complete": bool(decision.get("read_complete")),
                    "message_count": decision.get("message_count"),
                    "classified_media": len(decision.get("media_decisions", {})),
                    "evidence_media_count": decision.get("evidence_media_count"),
                    "open_orders": copy.deepcopy(decision.get("open_orders", [])),
                    "sealed": bool(decision.get("sealed")),
                    "controlled_editing": _controlled_editing(run_group),
                    "semantic_fingerprint": semantic_fingerprint,
                    "uncontrolled_semantic_changes": bool(
                        _controlled_editing(run_group)
                        and approved_fingerprint != semantic_fingerprint
                    ),
                }
            )
        return {"groups": reports}

    if args.review_action == "audit":
        selected = (
            [_select_run_group(run, args.group)]
            if args.group
            else [item for item in run["groups"] if isinstance(item, Mapping)]
        )
        audit_reports: list[dict[str, Any]] = []
        for run_group in selected:
            group_key = str(run_group["group_key"])
            decision = _load_json(_decision_path(work, run_group))
            _require_controlled_semantics(run_group, decision)
            statistics = _validate_decision(
                normalized,
                groups[group_key],
                copy.deepcopy(decision),
                require_complete=False,
                capture_hashes=False,
            )
            flags = {
                field: statistics[field]
                for field in RISK_FLAG_FIELDS
                if int(statistics.get(field) or 0)
            }
            audit_reports.append(
                {
                    "group_key": group_key,
                    "run_label": run_group.get("run_label"),
                    "sealed": bool(decision.get("sealed")),
                    "flags": flags,
                    **statistics,
                }
            )
        return {
            "summary": _review_risk_summary(audit_reports),
            "groups": audit_reports,
        }

    run_group = _select_run_group(
        run,
        args.group,
        prefer_incomplete=args.review_action == "next",
    )
    group_key = str(run_group["group_key"])
    group = groups[group_key]
    decision_path = _decision_path(work, run_group)
    decision = _load_json(decision_path)

    if args.review_action == "upgrade-checkpoints":
        return _upgrade_legacy_checkpoints(
            work,
            run,
            run_group,
            normalized,
            group,
            decision_path,
            decision,
        )

    if args.review_action == "apply-batch":
        core.require(decision.get("sealed") is not True, "sealed group cannot accept a review batch")
        controlled = _controlled_editing(run_group)
        if controlled:
            _require_controlled_semantics(run_group, decision)
        current_fingerprint = _semantic_fingerprint(decision)
        batch = _load_review_batch(args.input.resolve())
        input_fingerprint = core.fingerprint_json(batch)
        control = decision.get("edit_control")
        last_batch = (
            control.get("last_batch")
            if controlled and isinstance(control, Mapping)
            else None
        )
        if (
            isinstance(last_batch, Mapping)
            and core.clean_text(last_batch.get("id")) == batch["batch_id"]
        ):
            core.require(
                core.clean_text(last_batch.get("input_fingerprint"))
                == input_fingerprint,
                f"review batch id was already used with different content: {batch['batch_id']}",
            )
            core.require(
                core.clean_text(last_batch.get("result_fingerprint"))
                == current_fingerprint,
                "last applied batch no longer matches the current decision",
            )
            statistics = _validate_decision(
                normalized,
                group,
                copy.deepcopy(decision),
                require_complete=False,
                capture_hashes=False,
            )
            return {
                "group_key": group_key,
                "run_label": run_group.get("run_label"),
                "sealed": False,
                "controlled_editing": True,
                "batch_id": batch["batch_id"],
                "idempotent_replay": True,
                "semantic_fingerprint": current_fingerprint,
                "reviewed_through": decision.get("reviewed_through"),
                "read_complete": bool(decision.get("read_complete")),
                **statistics,
            }
        core.require(
            batch["base_fingerprint"] == current_fingerprint,
            "review batch is stale: base_fingerprint does not match the current decision; "
            "refresh review status and rebuild the batch",
        )
        candidate = _merge_review_batch(decision, batch)
        committed_page = _commit_review_page(candidate, batch, group)
        statistics = _validate_decision(
            normalized,
            group,
            candidate,
            require_complete=False,
            capture_hashes=True,
        )
        result_fingerprint = _semantic_fingerprint(candidate)
        prior_batch_count = (
            control.get("batch_count")
            if controlled
            and isinstance(control, Mapping)
            and isinstance(control.get("batch_count"), int)
            else 0
        )
        candidate["edit_control"] = {
            "mode": EDIT_CONTROL_MODE,
            "approved_semantic_fingerprint": result_fingerprint,
            "batch_count": prior_batch_count + 1,
            "last_batch": {
                "id": batch["batch_id"],
                "input_fingerprint": input_fingerprint,
                "base_fingerprint": current_fingerprint,
                "result_fingerprint": result_fingerprint,
                "applied_at": datetime.now(timezone.utc).isoformat(),
            },
        }
        core.atomic_json(decision_path, candidate)
        run_group["edit_mode"] = EDIT_CONTROL_MODE
        run_group["sealed"] = False
        core.atomic_json(_run_path(work), run)
        return {
            "group_key": group_key,
            "run_label": run_group.get("run_label"),
            "sealed": False,
            "controlled_editing": True,
            "batch_id": batch["batch_id"],
            "idempotent_replay": False,
            "semantic_fingerprint": result_fingerprint,
            "committed_page": (
                {"page_start": committed_page[0], "page_end": committed_page[1]}
                if committed_page is not None
                else None
            ),
            "reviewed_through": candidate.get("reviewed_through"),
            "read_complete": bool(candidate.get("read_complete")),
            **statistics,
        }

    if args.review_action == "next":
        core.require(
            decision.get("contract_version") == DECISION_CONTRACT,
            f"review next requires {DECISION_CONTRACT}; start a fresh work directory for the new safe paging protocol",
        )
        core.require(
            _require_controlled_semantics(run_group, decision),
            "review next requires the controlled apply-batch protocol",
        )
        exact_limit = args.limit
        if exact_limit is not None:
            core.require(
                1 <= exact_limit <= MAX_PAGE_SIZE,
                f"--limit must be 1..{MAX_PAGE_SIZE}",
            )
        start = int(decision.get("reviewed_through") or 0)
        messages = list(group.get("messages", []))
        maximum_count = exact_limit if exact_limit is not None else DEFAULT_PAGE_SIZE
        maximum_end = min(start + maximum_count, len(messages))
        group_fingerprint = core.fingerprint_json(group)
        semantic_fingerprint = _semantic_fingerprint(decision)
        open_orders = copy.deepcopy(decision.get("open_orders", []))
        carry_messages = _open_order_carry_messages(group, open_orders)

        def build_page(end: int) -> dict[str, Any]:
            return _review_page_result(
                group=group,
                group_key=group_key,
                run_group=run_group,
                decision=decision,
                start=start,
                end=end,
                group_fingerprint=group_fingerprint,
                semantic_fingerprint=semantic_fingerprint,
                open_orders=open_orders,
                carry_messages=carry_messages,
            )

        if exact_limit is not None or start == maximum_end:
            return build_page(maximum_end)

        selected_page: dict[str, Any] | None = None
        first_page_chars: int | None = None
        for end in range(start + 1, maximum_end + 1):
            candidate = build_page(end)
            output_chars = len(_render_result_json(candidate, compact=True))
            if first_page_chars is None:
                first_page_chars = output_chars
            if output_chars <= DEFAULT_REVIEW_PAGE_OUTPUT_CHAR_BUDGET:
                selected_page = candidate

        core.require(
            selected_page is not None,
            "review next cannot return even one complete message within the "
            f"{DEFAULT_REVIEW_PAGE_OUTPUT_CHAR_BUDGET}-character output budget "
            f"(first complete page needs {first_page_chars} characters); "
            "do not commit a truncated page",
        )
        return selected_page

    controlled = _require_controlled_semantics(run_group, decision)
    prior_fingerprint = _semantic_fingerprint(decision)
    statistics = _validate_decision(
        normalized,
        group,
        decision,
        require_complete=args.review_action == "seal",
        capture_hashes=True,
    )
    current_fingerprint = _semantic_fingerprint(decision)
    if controlled:
        control = decision["edit_control"]
        control["approved_semantic_fingerprint"] = current_fingerprint
        last_batch = control.get("last_batch")
        if (
            isinstance(last_batch, dict)
            and last_batch.get("result_fingerprint") == prior_fingerprint
        ):
            last_batch["result_fingerprint"] = current_fingerprint
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
        "controlled_editing": controlled,
        "semantic_fingerprint": current_fingerprint,
        **statistics,
    }


def _entry_event_type(entry: Mapping[str, Any]) -> str:
    side = core.clean_text(entry.get("side")).casefold()
    if core.clean_text(entry.get("kind")).casefold() == "cash":
        return "cash_payment" if side in {"payment", "payment_refund"} else "cash_payout"
    if side == "recovery":
        return "payout_recovery"
    if side == "payout":
        return "payout_screenshot"
    return "payment_screenshot"


def _translated_pricing(
    value: object,
    *,
    message_by_label: Mapping[str, Mapping[str, Any]],
) -> dict[str, Any]:
    core.require(isinstance(value, Mapping), "pricing must be an object")
    pricing = {
        key: copy.deepcopy(item)
        for key, item in value.items()
        if key != "source_messages"
    }
    pricing["source_message_ids"] = [
        str(message_by_label[str(label)]["message_id"])
        for label in value.get("source_messages", [])
    ]
    return pricing


def _translated_advanced_order(
    raw: Mapping[str, Any],
    *,
    event_by_entry: Mapping[str, str],
    message_by_label: Mapping[str, Mapping[str, Any]],
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
                    "pricing",
                }
            }
            leg["pricing"] = _translated_pricing(
                item.get("pricing"),
                message_by_label=message_by_label,
            )
            leg["payout_event_ids"] = [
                event_by_entry[str(value)] for value in item.get("payout_entry_ids", [])
            ]
            if item.get("recovery_entry_ids") not in (None, []):
                leg["recovery_event_ids"] = [
                    event_by_entry[str(value)] for value in item.get("recovery_entry_ids", [])
                ]
            compiled_legs.append(leg)
        result["legs"] = compiled_legs
    return result


def _compile_decisions_v3(
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
        message_label_by_id, message_by_label = _message_labels(group)
        del message_label_by_id
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
                        "payee_state": entry.get("payee_state"),
                        "status_text": entry.get("status_text"),
                        "status_class": STATUS_CLASS[str(entry["result"])],
                        "status_class_confidence": "high",
                        "amount_completeness": "complete" if amount_state == "clear" else amount_state,
                        "confidence": "high" if amount_state == "clear" else "low",
                    },
                }
                event["flow_side"] = entry["side"]
                if isinstance(entry.get("side_exception"), Mapping):
                    side_exception = entry["side_exception"]
                    event["side_exception"] = {
                        "kind": side_exception["kind"],
                        "source_message_ids": [
                            str(message_by_label[str(source_label)]["message_id"])
                            for source_label in side_exception["source_messages"]
                        ],
                        "detail": side_exception["detail"],
                    }
                side_by_entry[entry_id] = str(entry["side"])
                if entry.get("note"):
                    event["note"] = entry["note"]
                events.append(event)
                event_by_entry[entry_id] = event_id

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
                "customer_nickname",
                "direction",
            ):
                order[field] = copy.deepcopy(raw.get(field))
            for field in ("note",):
                if raw.get(field) not in (None, ""):
                    order[field] = copy.deepcopy(raw[field])
            if raw.get("pricing") not in (None, ""):
                order["pricing"] = _translated_pricing(
                    raw.get("pricing"),
                    message_by_label=message_by_label,
                )
            event_sides = {
                event_by_entry[entry_id]: side_by_entry[entry_id]
                for entry_id in entry_ids
                if entry_id in side_by_entry
            }
            order["event_sides"] = event_sides
            order.update(
                _translated_advanced_order(
                    raw,
                    event_by_entry=event_by_entry,
                    message_by_label=message_by_label,
                )
            )
            plan_orders.append(order)
        settlement_allocations = []
        for raw_relation in decision.get("settlement_allocations", []):
            relation = {
                "source_event_id": event_by_entry[str(raw_relation["entry_id"])],
                "allocations": [
                    {
                        "target_case_id": case_by_order_id[str(item["order_id"])],
                        "amount": copy.deepcopy(item["amount"]),
                    }
                    for item in raw_relation["allocations"]
                ],
            }
            if raw_relation.get("source_messages"):
                relation["source_message_ids"] = [
                    str(message_by_label[str(item)]["message_id"])
                    for item in raw_relation["source_messages"]
                ]
            settlement_allocations.append(relation)
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
                "settlement_allocations": settlement_allocations,
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
    finish_risk_reports: list[dict[str, Any]] = []
    for run_group in run["groups"]:
        group_key = str(run_group["group_key"])
        decision = _load_json(_decision_path(work, run_group))
        _require_controlled_semantics(run_group, decision)
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
        finish_risk_reports.append(
            {
                "group_key": group_key,
                "run_label": run_group.get("run_label"),
                **statistics,
            }
        )

    events, plan = _compile_decisions_v3(normalized, decisions)
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
        "risk_report": _review_risk_summary(finish_risk_reports),
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
    compact = args.command == "review" and args.review_action == "next"
    sys.stdout.write(_render_result_json(result, compact=compact))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
