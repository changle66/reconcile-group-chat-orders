#!/usr/bin/env python3
"""Normalize Telegram JSON and WhatsApp text exports without external skills."""

from __future__ import annotations

import argparse
import json
import mimetypes
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import core


NORMALIZER_VERSION = "1.1.0"
WHATSAPP_HEADER_RE = re.compile(
    r"^(?P<date>\d{4}/\d{1,2}/\d{1,2})\s+(?P<time>\d{1,2}:\d{2})\s+-\s+(?P<body>.*)$"
)
WHATSAPP_ATTACHMENT_RE = re.compile(
    r"(?P<name>[^\\/:*?\"<>|\r\n]+?\.(?:jpe?g|png|webp|gif|heic|pdf|mp4|mov|m4a|mp3|opus|ogg|wav|was))"
    r"(?:\s*\((?:文件附件|file attached)\))?",
    re.IGNORECASE,
)
MENTION_RE = re.compile(r"@([^@\n\r，。；、]{1,64})")
WHATSAPP_SENDER_RE = re.compile(r"^(?P<sender>[^:\n]{1,120}):(?:\s(?P<text>[\s\S]*))?$")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", type=Path, help="Export files or directories")
    parser.add_argument("-o", "--output", required=True, type=Path)
    parser.add_argument("--timezone", default="Asia/Bangkok")
    parser.add_argument(
        "--roster",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "config" / "roster.yaml",
    )
    parser.add_argument("--force", action="store_true")
    return parser.parse_args(argv)


def discover_sources(inputs: Iterable[Path]) -> tuple[list[Path], list[Path]]:
    telegram: set[Path] = set()
    whatsapp: set[Path] = set()
    for raw in inputs:
        path = raw.resolve()
        if path.is_file():
            if path.name.casefold() == "result.json":
                telegram.add(path)
            elif path.suffix.casefold() == ".txt":
                whatsapp.add(path)
            continue
        if not path.is_dir():
            raise ValueError(f"input does not exist: {path}")
        direct_json = path / "result.json"
        if direct_json.is_file():
            telegram.add(direct_json.resolve())
        for candidate in path.glob("*.txt"):
            if "whatsapp" in candidate.name.casefold():
                whatsapp.add(candidate.resolve())
    if not telegram and not whatsapp:
        raise ValueError("no Telegram result.json or WhatsApp .txt exports found")
    return (
        sorted(telegram, key=lambda item: str(item).casefold()),
        sorted(whatsapp, key=lambda item: str(item).casefold()),
    )


def media_kind(path: Path, media_type: object = None) -> tuple[str, str | None]:
    mime, _ = mimetypes.guess_type(path.name)
    hint = str(media_type or "").casefold()
    compact_name = path.name.casefold()
    if hint in {"sticker", "animated_sticker", "animation"} or re.search(r"(?:^|[-_])(?:sticker|stk)(?:[-_.]|$)", compact_name):
        return "sticker", mime or "image/webp"
    if mime == "image/gif":
        return "animation", mime
    if hint in {"video", "video_file", "video_message", "round_video_message"}:
        return "video", mime or "video/mp4"
    if hint in {"audio", "audio_file", "voice_message"}:
        return "audio", mime or "audio/ogg"
    if (mime and mime.startswith("image/")) or hint in {"photo", "image"}:
        return "image", mime or "image/jpeg"
    if mime == "application/pdf":
        return "document", mime
    if mime and mime.startswith("video/"):
        return "video", mime
    if mime and mime.startswith("audio/"):
        return "audio", mime
    return "file", mime


def telegram_media(message: dict[str, Any], source: Path) -> tuple[list[dict[str, Any]], list[str], list[Path]]:
    media: list[dict[str, Any]] = []
    warnings: list[str] = []
    fingerprint_paths: list[Path] = []
    # Telegram may export both a primary file and its thumbnail.  They are one
    # message attachment, not two accounting proofs, so choose one source in
    # priority order and keep a missing primary as missing evidence.
    selected_fields: list[str] = []
    for candidate_field in ("photo", "file", "thumbnail"):
        if isinstance(message.get(candidate_field), str) and message.get(candidate_field, "").strip():
            selected_fields = [candidate_field]
            break
    for field in selected_fields:
        raw_reference = message.get(field)
        if not isinstance(raw_reference, str) or not raw_reference.strip():
            continue
        reference = core.clean_text(raw_reference)
        candidate = (source.parent / reference).resolve()
        kind, mime = media_kind(candidate, message.get("media_type"))
        # Voice, video, stickers and animations are deliberately outside the
        # accounting evidence contract. Keep the chat message, but do not hash,
        # dispatch or require classification of the attachment.
        if kind in core.NON_EVIDENCE_MEDIA_KINDS:
            continue
        available = candidate.is_file()
        media.append(
            {
                "kind": kind,
                "mime_type": mime,
                "availability": "available" if available else "missing",
                "source_field": field,
                "path": str(candidate) if available else None,
                "original_reference": reference,
                "blob_sha256": core.sha256_file(candidate) if available else None,
                "byte_size": candidate.stat().st_size if available else None,
                "missing_kind": "not_exported" if not available and reference.startswith("(File not included") else ("missing" if not available else None),
            }
        )
        if available:
            fingerprint_paths.append(candidate)
        else:
            warnings.append(f"Telegram media missing: {source} :: {reference}")
    return media, warnings, fingerprint_paths


def telegram_timestamp(message: dict[str, Any], zone: ZoneInfo) -> str:
    raw_epoch = message.get("date_unixtime")
    if raw_epoch not in (None, ""):
        return datetime.fromtimestamp(int(raw_epoch), tz=timezone.utc).astimezone(zone).isoformat()
    raw_date = str(message.get("date") or "")
    if not raw_date:
        raise ValueError("Telegram message has no date")
    parsed = datetime.fromisoformat(raw_date)
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=zone)
    return parsed.astimezone(zone).isoformat()


def normalize_telegram(
    source: Path,
    *,
    zone: ZoneInfo,
    roster: dict[str, Any],
) -> tuple[dict[str, Any], list[str], list[Path]]:
    data = json.loads(source.read_text(encoding="utf-8-sig"))
    group_id = str(data.get("id") or core.stable_token(source.name, data.get("name"), length=16))
    group_key = f"telegram:{group_id}"
    group_name = core.clean_text(data.get("name")) or group_key
    warnings: list[str] = []
    fingerprint_paths: list[Path] = [source]
    messages: list[dict[str, Any]] = []
    for raw_sequence, raw_message in enumerate(data.get("messages", []), start=1):
        if not isinstance(raw_message, dict):
            continue
        local_id = str(raw_message.get("id") or f"seq-{raw_sequence}")
        message_id = f"{group_key}:{local_id}"
        text = core.flatten_telegram_text(raw_message.get("text"))
        sender_name = core.clean_text(raw_message.get("from")) or None
        sender_id = core.clean_text(raw_message.get("from_id")) or None
        role = core.classify_role(sender_name or sender_id, roster=roster)
        media, media_warnings, media_paths = telegram_media(raw_message, source)
        for media_index, media_item in enumerate(media):
            media_item["media_id"] = f"{message_id}#media:{media_index}"
        warnings.extend(f"{message_id}: {item}" for item in media_warnings)
        fingerprint_paths.extend(media_paths)
        raw_reply = raw_message.get("reply_to_message_id")
        reply_to = f"{group_key}:{raw_reply}" if raw_reply not in (None, "") else None
        message_type = str(raw_message.get("type") or "message")
        excluded = message_type != "message" or role == "未知"
        mentions = [core.clean_text(match.group(1)) for match in MENTION_RE.finditer(text)]
        messages.append(
            {
                "message_id": message_id,
                "timestamp": telegram_timestamp(raw_message, zone),
                "sender_name": sender_name,
                "sender_id": sender_id,
                "role": role,
                "text": text,
                "reply_to_message_id": reply_to,
                "mentions": list(dict.fromkeys(item for item in mentions if item)),
                "media": media,
                "source_file": str(source),
                "source_sequence": raw_sequence,
                "excluded_from_accounting": excluded,
            }
        )
    return (
        {
            "group_key": group_key,
            "platform": "Telegram",
            "timestamp_basis": "epoch",
            "source_timezone": None,
            "group_name": group_name,
            "source_files": [str(source)],
            "messages": messages,
            "rate_candidates": [],
        },
        warnings,
        fingerprint_paths,
    )


def whatsapp_group_name(source: Path) -> str:
    name = source.stem
    if name.startswith("与") and "的 WhatsApp 聊天" in name:
        name = name[1 : name.rfind("的 WhatsApp 聊天")]
    return core.clean_text(name) or source.stem


def split_whatsapp_blocks(source: Path) -> list[tuple[datetime, str, int]]:
    blocks: list[tuple[datetime, str, int]] = []
    current_time: datetime | None = None
    current_lines: list[str] = []
    current_sequence = 0
    for line_number, line in enumerate(source.read_text(encoding="utf-8-sig").splitlines(), start=1):
        match = WHATSAPP_HEADER_RE.match(line)
        if match:
            if current_time is not None:
                blocks.append((current_time, "\n".join(current_lines), current_sequence))
            current_time = datetime.strptime(
                f"{match.group('date')} {match.group('time')}", "%Y/%m/%d %H:%M"
            )
            current_lines = [match.group("body")]
            current_sequence = line_number
        elif current_time is not None:
            current_lines.append(line)
    if current_time is not None:
        blocks.append((current_time, "\n".join(current_lines), current_sequence))
    return blocks


def whatsapp_sender_and_text(body: str) -> tuple[str | None, str]:
    cleaned = core.clean_text(body)
    match = WHATSAPP_SENDER_RE.match(cleaned)
    if match is None:
        return None, cleaned
    sender = core.clean_text(match.group("sender"))
    text = core.clean_text(match.group("text") or "")
    if not sender or len(sender) > 120:
        return None, cleaned
    return sender, text


def whatsapp_media(text: str, source: Path) -> tuple[list[dict[str, Any]], list[str], list[Path]]:
    media: list[dict[str, Any]] = []
    warnings: list[str] = []
    fingerprint_paths: list[Path] = []
    seen: set[str] = set()
    for match in WHATSAPP_ATTACHMENT_RE.finditer(text):
        reference = core.clean_text(match.group("name")).strip()
        if reference in seen:
            continue
        seen.add(reference)
        candidate = (source.parent / reference).resolve()
        kind, mime = media_kind(candidate)
        if kind in core.NON_EVIDENCE_MEDIA_KINDS:
            continue
        available = candidate.is_file()
        media.append(
            {
                "kind": kind,
                "mime_type": mime,
                "availability": "available" if available else "missing",
                "source_field": "whatsapp_attachment",
                "path": str(candidate) if available else None,
                "original_reference": reference,
                "blob_sha256": core.sha256_file(candidate) if available else None,
                "byte_size": candidate.stat().st_size if available else None,
                "missing_kind": None if available else "missing",
            }
        )
        if available:
            fingerprint_paths.append(candidate)
        else:
            warnings.append(f"WhatsApp media missing: {source} :: {reference}")
    return media, warnings, fingerprint_paths


def sender_identifier(sender: str | None) -> str | None:
    if sender is None:
        return None
    digits = re.sub(r"\D", "", sender)
    if sender.lstrip().startswith("+") and digits:
        return "+" + digits
    return "name:" + core.stable_token(core.normalize_name(sender), length=20)


def normalize_whatsapp(
    source: Path,
    *,
    zone: ZoneInfo,
    roster: dict[str, Any],
) -> tuple[dict[str, Any], list[str], list[Path]]:
    group_name = whatsapp_group_name(source)
    blocks = split_whatsapp_blocks(source)
    if not blocks:
        raise ValueError(f"{source}: no WhatsApp timestamped messages found")
    group_token = core.stable_token("whatsapp-group", core.normalize_name(group_name), length=16)
    group_key = f"whatsapp:{group_token}"
    warnings: list[str] = []
    fingerprint_paths: list[Path] = [source]
    staged: list[dict[str, Any]] = []
    snapshot_hash = core.sha256_file(source)
    for naive_time, body, raw_sequence in blocks:
        sender_name, text = whatsapp_sender_and_text(body)
        sender_id = sender_identifier(sender_name)
        role = core.classify_role(sender_name, roster=roster)
        media, local_warnings, local_paths = whatsapp_media(text, source)
        fingerprint_paths.extend(local_paths)
        timestamp = naive_time.replace(tzinfo=zone).isoformat()
        message_id = f"{group_key}:snapshot:{snapshot_hash[:12]}:line:{raw_sequence}"
        for media_index, media_item in enumerate(media):
            media_item["media_id"] = f"{message_id}#media:{media_index}"
        warnings.extend(f"{message_id}: {item}" for item in local_warnings)
        mentions = [core.clean_text(match.group(1)) for match in MENTION_RE.finditer(text)]
        staged.append(
            {
                "message_id": message_id,
                "timestamp": timestamp,
                "sender_name": sender_name,
                "sender_id": sender_id,
                "role": role,
                "text": text,
                "reply_to_message_id": None,
                "mentions": list(dict.fromkeys(item for item in mentions if item)),
                "media": media,
                "source_file": str(source),
                "source_sequence": raw_sequence,
                "excluded_from_accounting": sender_name is None or role == "未知",
            }
        )
    for index, message in enumerate(staged, start=1):
        message["source_sequence"] = index
    return (
        {
            "group_key": group_key,
            "platform": "WhatsApp",
            "timestamp_basis": "assumed_local",
            "source_timezone": str(zone.key),
            "group_name": group_name,
            "source_files": [str(source)],
            "messages": staged,
            "rate_candidates": [],
        },
        warnings,
        fingerprint_paths,
    )


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    output = args.output.resolve()
    if output.exists() and not args.force:
        print(f"Output already exists (use --force): {output}", file=sys.stderr)
        return 2
    try:
        zone = ZoneInfo(args.timezone)
        if args.timezone != "Asia/Bangkok":
            raise ValueError("this skill's accounting timezone must be Asia/Bangkok unless its rules are revised")
        roster = core.load_roster(args.roster.resolve())
        telegram_sources, whatsapp_sources = discover_sources(args.inputs)
        groups: list[dict[str, Any]] = []
        warnings: list[str] = []
        fingerprint_paths: list[Path] = []
        for source in telegram_sources:
            group, local_warnings, local_paths = normalize_telegram(source, zone=zone, roster=roster)
            groups.append(group)
            warnings.extend(local_warnings)
            fingerprint_paths.extend(local_paths)
        for source in whatsapp_sources:
            group, local_warnings, local_paths = normalize_whatsapp(source, zone=zone, roster=roster)
            groups.append(group)
            warnings.extend(local_warnings)
            fingerprint_paths.extend(local_paths)
        keys = [group["group_key"] for group in groups]
        if len(keys) != len(set(keys)):
            raise ValueError("duplicate group_key; provide one export per group")
        groups.sort(key=lambda item: (item["platform"], item["group_name"], item["group_key"]))
        source_fingerprint = core.fingerprint_files(
            fingerprint_paths,
            context={
                "normalizer_version": NORMALIZER_VERSION,
                "timezone": args.timezone,
                "roster": roster,
                "groups": [(group["platform"], group["group_name"], group["group_key"]) for group in groups],
            },
        )
        all_messages = [message for group in groups for message in group["messages"]]
        result = {
            "contract_version": core.NORMALIZED_CONTRACT,
            "timezone": args.timezone,
            "source_fingerprint": source_fingerprint,
            "groups": groups,
            "statistics": {
                "groups": len(groups),
                "messages": len(all_messages),
                "excluded_messages": sum(bool(message["excluded_from_accounting"]) for message in all_messages),
                "media": sum(len(message["media"]) for message in all_messages),
                "available_media": sum(
                    media.get("availability") == "available"
                    for message in all_messages
                    for media in message["media"]
                ),
                "missing_media": sum(
                    media.get("availability") == "missing"
                    for message in all_messages
                    for media in message["media"]
                ),
                "rate_candidates": sum(len(group["rate_candidates"]) for group in groups),
                "source_files": len(telegram_sources) + len(whatsapp_sources),
                "fingerprinted_files": len({path.resolve() for path in fingerprint_paths}),
            },
            "warnings": warnings,
            "skipped_files": [],
        }
        core.atomic_json(output, result)
        core.load_normalized(output)
    except (OSError, ValueError, json.JSONDecodeError, ZoneInfoNotFoundError) as exc:
        print(f"Normalization failed: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result["statistics"], ensure_ascii=False))
    if warnings:
        print(f"Warnings: {len(warnings)} (see normalized output)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
