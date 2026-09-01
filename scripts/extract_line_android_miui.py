#!/usr/bin/env python3
"""Extract LINE group chats from an unencrypted MIUI Android app backup."""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
import sys
import tarfile
import tempfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import core as local_core


PACKAGE_NAME = "jp.naver.line.android"
APP_PREFIX = f"apps/{PACKAGE_NAME}"
DATABASE_PREFIX = f"{APP_PREFIX}/db"
MEDIA_PREFIX = f"{APP_PREFIX}/ef/chats"
ANDROID_BACKUP_MARKER = b"ANDROID BACKUP\n"
MIUI_BACKUP_MARKER = b"MIUI BACKUP\n"
EXTRACTOR_VERSION = "1.0.0"
DEFAULT_GROUP_PATTERNS = (r"小额", r"钱庄", r"出[🐱😺🐷]")
IMAGE_ATTACHMENT_TYPE = 1
TEXT_ATTACHMENT_TYPE = 0
MENTION_RE = re.compile(r"@([^@\n\r]{1,64}?)(?=\s{2,}|$|[，。；、])")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("backup", type=Path, help="MIUI backup folder or LINE .bak file")
    parser.add_argument("-o", "--output", type=Path, help="normalized.json output")
    parser.add_argument("--timezone", default="Asia/Bangkok", help="IANA accounting timezone")
    parser.add_argument(
        "--group-pattern",
        action="append",
        default=None,
        help="Regex selecting group names; repeatable (defaults to small-FX naming patterns)",
    )
    parser.add_argument("--group-id", action="append", default=[], help="Exact LINE group MID; repeatable")
    parser.add_argument("--all-group-chats", action="store_true", help="Include every LINE group chat")
    parser.add_argument("--list-groups", action="store_true", help="List available LINE group chats and exit")
    parser.add_argument(
        "--roster",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "config" / "roster.yaml",
        help="Staff roster YAML",
    )
    parser.add_argument(
        "--media-dir",
        type=Path,
        default=None,
        help="Extracted media directory (default: <output-stem>_media beside output)",
    )
    parser.add_argument("--force", action="store_true", help="Replace output and extracted media files")
    return parser.parse_args(argv)


def _backup_header(handle: Any) -> dict[str, Any]:
    handle.seek(0)
    prefix = handle.read(8192)
    if not prefix.startswith(MIUI_BACKUP_MARKER):
        raise ValueError("not a supported MIUI application backup")
    marker_at = prefix.find(ANDROID_BACKUP_MARKER)
    if marker_at < 0:
        raise ValueError("not an Android Backup stream: ANDROID BACKUP header is missing")
    if PACKAGE_NAME.encode("ascii") not in prefix[:marker_at]:
        raise ValueError(f"MIUI backup is not for {PACKAGE_NAME}")
    handle.seek(marker_at + len(ANDROID_BACKUP_MARKER))
    version = handle.readline().strip().decode("ascii", errors="strict")
    compressed = handle.readline().strip().decode("ascii", errors="strict")
    encryption = handle.readline().strip().decode("ascii", errors="strict")
    if not version.isdigit():
        raise ValueError(f"invalid Android Backup version: {version!r}")
    if compressed != "0":
        raise ValueError("compressed Android Backup streams are not supported; expected compression flag 0")
    if encryption != "none":
        raise ValueError("encrypted Android Backup streams are not supported")
    return {
        "android_backup_version": version,
        "compressed": compressed,
        "encryption": encryption,
        "tar_offset": handle.tell(),
        "miui_wrapped": True,
    }


def backup_header(path: Path) -> dict[str, Any]:
    resolved = path.resolve()
    if not resolved.is_file():
        raise ValueError(f"LINE Android backup does not exist: {resolved}")
    with resolved.open("rb") as handle:
        return _backup_header(handle)


def is_line_android_backup(path: Path) -> bool:
    try:
        backup_header(path)
        return True
    except (OSError, UnicodeError, ValueError):
        return False


def find_backup_file(input_path: Path) -> Path:
    path = input_path.resolve()
    if path.is_file():
        if not is_line_android_backup(path):
            raise ValueError(f"not a supported LINE Android backup: {path}")
        return path
    if not path.is_dir():
        raise ValueError(f"backup path does not exist: {path}")
    matches = sorted(
        (candidate.resolve() for candidate in path.rglob("*.bak") if is_line_android_backup(candidate)),
        key=lambda item: str(item).casefold(),
    )
    if not matches:
        raise ValueError(f"no supported LINE Android .bak file found under {path}")
    if len(matches) > 1:
        details = ", ".join(str(item) for item in matches)
        raise ValueError(f"multiple LINE Android backups found under {path}; pass one .bak file: {details}")
    return matches[0]


@contextmanager
def open_backup_tar(path: Path) -> Iterator[tarfile.TarFile]:
    with path.resolve().open("rb") as handle:
        header = _backup_header(handle)
        handle.seek(int(header["tar_offset"]))
        with tarfile.open(fileobj=handle, mode="r:") as archive:
            yield archive


def _member_index(archive: tarfile.TarFile) -> dict[str, tarfile.TarInfo]:
    return {member.name: member for member in archive.getmembers()}


def _write_member(archive: tarfile.TarFile, member: tarfile.TarInfo, destination: Path) -> None:
    source = archive.extractfile(member)
    if source is None:
        raise ValueError(f"backup member cannot be read: {member.name}")
    destination.write_bytes(source.read())


def sqlite_ro(path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(f"{path.resolve().as_uri()}?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    return connection


def _require_columns(connection: sqlite3.Connection, table: str, required: Iterable[str]) -> None:
    columns = {str(row[1]) for row in connection.execute(f'PRAGMA table_info("{table}")')}
    missing = sorted(set(required) - columns)
    if missing:
        raise ValueError(f"LINE Android database table {table!r} is missing columns: {', '.join(missing)}")


@contextmanager
def open_line_databases(
    backup: Path,
) -> Iterator[tuple[sqlite3.Connection, sqlite3.Connection | None]]:
    with tempfile.TemporaryDirectory(prefix="line-android-miui-") as temporary:
        root = Path(temporary)
        with open_backup_tar(backup) as archive:
            members = _member_index(archive)
            for stem in ("naver_line", "contact"):
                for suffix in ("", "-wal", "-shm", "-journal"):
                    name = f"{DATABASE_PREFIX}/{stem}{suffix}"
                    member = members.get(name)
                    if member is not None and member.isfile():
                        _write_member(archive, member, root / f"{stem}{suffix}")
        line_path = root / "naver_line"
        if not line_path.is_file():
            raise ValueError(f"LINE Android database is missing from backup: {DATABASE_PREFIX}/naver_line")
        line_db = sqlite_ro(line_path)
        contact_db: sqlite3.Connection | None = None
        try:
            _require_columns(line_db, "groups", {"id", "name"})
            _require_columns(
                line_db,
                "chat_history",
                {
                    "id",
                    "server_id",
                    "type",
                    "chat_id",
                    "from_mid",
                    "content",
                    "created_time",
                    "status",
                    "attachement_type",
                    "parameter",
                },
            )
            contact_path = root / "contact"
            if contact_path.is_file():
                contact_db = sqlite_ro(contact_path)
                _require_columns(
                    contact_db,
                    "contacts",
                    {"mid", "profile_name", "overridden_name"},
                )
            yield line_db, contact_db
        finally:
            if contact_db is not None:
                contact_db.close()
            line_db.close()


def message_timestamp(raw_value: object, zone: ZoneInfo) -> str:
    try:
        value = int(str(raw_value))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"invalid LINE Android message timestamp: {raw_value!r}") from exc
    if value > 10_000_000_000_000:
        seconds = value / 1_000_000
    elif value > 10_000_000_000:
        seconds = value / 1_000
    else:
        seconds = value
    return datetime.fromtimestamp(seconds, tz=timezone.utc).astimezone(zone).isoformat()


def available_groups(line_db: sqlite3.Connection) -> list[dict[str, Any]]:
    rows = list(
        line_db.execute(
            """
            SELECT g.id AS group_mid, COALESCE(NULLIF(g.name, ''), g.id) AS group_name,
                   COUNT(h.id) AS message_count,
                   MIN(CAST(h.created_time AS INTEGER)) AS first_timestamp,
                   MAX(CAST(h.created_time AS INTEGER)) AS last_timestamp
            FROM groups g
            LEFT JOIN chat_history h ON h.chat_id = g.id
            WHERE g.id LIKE 'c%'
            GROUP BY g.id, g.name
            ORDER BY g.id
            """
        )
    )
    return [
        {
            "group_mid": str(row["group_mid"]),
            "group_name": str(row["group_name"]),
            "message_count": int(row["message_count"] or 0),
            "first_timestamp": row["first_timestamp"],
            "last_timestamp": row["last_timestamp"],
        }
        for row in rows
    ]


def available_groups_for_backup(backup: Path) -> list[dict[str, Any]]:
    resolved = find_backup_file(backup)
    with open_line_databases(resolved) as (line_db, _):
        return available_groups(line_db)


def selected_groups(groups: list[dict[str, Any]], args: argparse.Namespace) -> list[dict[str, Any]]:
    if args.all_group_chats:
        return groups
    exact = set(args.group_id)
    unknown = exact - {group["group_mid"] for group in groups}
    if unknown:
        raise ValueError(f"Unknown --group-id value(s): {', '.join(sorted(unknown))}")
    raw_patterns = (
        args.group_pattern
        if args.group_pattern is not None
        else ([] if exact else list(DEFAULT_GROUP_PATTERNS))
    )
    try:
        patterns = [re.compile(pattern, re.IGNORECASE) for pattern in raw_patterns]
    except re.error as exc:
        raise ValueError(f"Invalid --group-pattern: {exc}") from exc
    return [
        group
        for group in groups
        if group["group_mid"] in exact
        or any(pattern.search(group["group_name"]) for pattern in patterns)
    ]


def contact_names(contact_db: sqlite3.Connection | None) -> dict[str, str]:
    if contact_db is None:
        return {}
    result: dict[str, str] = {}
    for row in contact_db.execute("SELECT mid, profile_name, overridden_name FROM contacts"):
        mid = local_core.clean_text(row["mid"])
        name = local_core.clean_text(row["overridden_name"]) or local_core.clean_text(row["profile_name"])
        if mid and name:
            result[mid] = name
    return result


def parse_parameters(raw_value: object) -> dict[str, str]:
    raw = str(raw_value or "")
    parts = raw.split("\t")
    return {
        local_core.clean_text(parts[index]): parts[index + 1]
        for index in range(0, len(parts) - 1, 2)
        if local_core.clean_text(parts[index])
    }


def parameter_mentions(
    parameters: Mapping[str, str],
    text: str,
    names: Mapping[str, str],
) -> list[str]:
    mentions: list[str] = []
    raw_mention = parameters.get("MENTION")
    if raw_mention:
        try:
            parsed = json.loads(raw_mention)
        except json.JSONDecodeError:
            parsed = None
        if isinstance(parsed, dict) and isinstance(parsed.get("MENTIONEES"), list):
            for item in parsed["MENTIONEES"]:
                if not isinstance(item, dict):
                    continue
                mid = local_core.clean_text(item.get("M"))
                value = names.get(mid) or mid
                if value:
                    mentions.append(value)
    mentions.extend(local_core.clean_text(match.group(1)) for match in MENTION_RE.finditer(text))
    return list(dict.fromkeys(item for item in mentions if item))


def system_message_text(
    message_type: int,
    attachment_type: int,
    parameters: Mapping[str, str],
    names: Mapping[str, str],
) -> str:
    location_key = local_core.clean_text(parameters.get("LOC_KEY"))
    participants = [
        names.get(mid, mid)
        for mid in str(parameters.get("LOC_ARGS") or "").split("\x1e")
        if local_core.clean_text(mid)
    ]
    label = f"LINE系统消息 {location_key}" if location_key else f"LINE消息类型 {message_type}/{attachment_type}"
    if participants:
        label += "：" + "、".join(participants)
    return f"[{label}]"


def media_signature(data: bytes) -> tuple[str, str | None, str]:
    if data.startswith(b"\xff\xd8\xff"):
        return ".jpg", "image/jpeg", "image"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png", "image/png", "image"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return ".gif", "image/gif", "animation"
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return ".webp", "image/webp", "image"
    if data.startswith(b"%PDF-"):
        return ".pdf", "application/pdf", "document"
    if len(data) >= 12 and data[4:8] == b"ftyp":
        brand = data[8:12]
        if brand in {b"heic", b"heix", b"hevc", b"hevx", b"mif1", b"msf1"}:
            return ".heic", "image/heic", "image"
        return ".mp4", "video/mp4", "video"
    return ".bin", None, "image"


def _media_member(
    members: Mapping[str, tarfile.TarInfo], group_mid: str, local_id: int
) -> tuple[tarfile.TarInfo | None, str | None, str | None]:
    prefix = f"{MEDIA_PREFIX}/{group_mid}/messages/{local_id}"
    for suffix, source_field, variant in (
        (".original", "line_android_attachment", "original"),
        ("", "line_android_attachment", "primary"),
        (".thumb", "line_android_thumbnail", "thumbnail"),
    ):
        member = members.get(prefix + suffix)
        if member is not None and member.isfile():
            return member, source_field, variant
    return None, None, None


def copied_media_path(
    data: bytes,
    *,
    media_root: Path,
    group_mid: str,
    local_id: int,
    extension: str,
    force: bool,
) -> Path:
    safe_group = re.sub(r"[^A-Za-z0-9._-]", "_", group_mid) or local_core.stable_token(group_mid)
    group_dir = media_root / safe_group
    group_dir.mkdir(parents=True, exist_ok=True)
    destination = group_dir / f"{local_id}{extension}"
    if destination.exists() and not force:
        if destination.read_bytes() != data:
            raise ValueError(f"Media destination exists with different content: {destination}")
        return destination.resolve()
    destination.write_bytes(data)
    return destination.resolve()


def normalize_group(
    *,
    line_db: sqlite3.Connection,
    archive: tarfile.TarFile,
    members: Mapping[str, tarfile.TarInfo],
    group: Mapping[str, Any],
    zone: ZoneInfo,
    roster: Mapping[str, Any],
    names: Mapping[str, str],
    media_root: Path,
    force: bool,
    source_files: list[str],
) -> tuple[dict[str, Any], list[str]]:
    group_mid = str(group["group_mid"])
    group_name = str(group["group_name"])
    rows = list(
        line_db.execute(
            """
            SELECT id, server_id, type, from_mid, content, created_time,
                   status, attachement_type, parameter
            FROM chat_history
            WHERE chat_id = ?
            ORDER BY CAST(created_time AS INTEGER), id
            """,
            (group_mid,),
        )
    )
    message_ids: list[str] = []
    server_to_message: dict[str, str] = {}
    used: set[str] = set()
    for row in rows:
        server_id = local_core.clean_text(row["server_id"])
        base = server_id or f"local-{int(row['id'])}"
        local_key = base if base not in used else f"{base}:local-{int(row['id'])}"
        used.add(local_key)
        message_id = f"line:{group_mid}:{local_key}"
        message_ids.append(message_id)
        if server_id and server_id not in server_to_message:
            server_to_message[server_id] = message_id

    warnings: list[str] = []
    messages: list[dict[str, Any]] = []
    for sequence, (row, message_id) in enumerate(zip(rows, message_ids), start=1):
        local_id = int(row["id"])
        message_type = int(row["type"] or 0)
        attachment_type = int(row["attachement_type"] or 0)
        parameters = parse_parameters(row["parameter"])
        text = local_core.clean_text(row["content"])
        is_context_only = message_type != 1 or attachment_type not in {
            TEXT_ATTACHMENT_TYPE,
            IMAGE_ATTACHMENT_TYPE,
        }
        if is_context_only and not text:
            text = system_message_text(message_type, attachment_type, parameters, names)

        sender_mid = local_core.clean_text(row["from_mid"]) or None
        if is_context_only:
            sender_name = None
            role = "未知"
        else:
            sender_name = names.get(sender_mid or "") or None
            role = local_core.classify_role(sender_name or sender_mid, roster=roster)

        media: list[dict[str, Any]] = []
        if attachment_type == IMAGE_ATTACHMENT_TYPE:
            member, source_field, variant = _media_member(members, group_mid, local_id)
            if member is None:
                media.append(
                    {
                        "media_id": f"{message_id}#media:0",
                        "kind": "image",
                        "mime_type": None,
                        "availability": "missing",
                        "source_field": "line_android_attachment",
                        "path": None,
                        "original_reference": f"LINE Android message {local_id}",
                        "blob_sha256": None,
                        "byte_size": None,
                        "variant": None,
                        "missing_kind": "missing_from_backup",
                    }
                )
                warnings.append(f"{message_id}: image attachment is missing from the backup")
            else:
                source = archive.extractfile(member)
                if source is None:
                    raise ValueError(f"backup media member cannot be read: {member.name}")
                data = source.read()
                extension, mime_type, kind = media_signature(data)
                if kind in {"image", "document"}:
                    final_path = copied_media_path(
                        data,
                        media_root=media_root,
                        group_mid=group_mid,
                        local_id=local_id,
                        extension=extension,
                        force=force,
                    )
                    media.append(
                        {
                            "media_id": f"{message_id}#media:0",
                            "kind": kind,
                            "mime_type": mime_type,
                            "availability": "available",
                            "source_field": source_field,
                            "path": str(final_path),
                            "original_reference": member.name,
                            "blob_sha256": None,
                            "byte_size": len(data),
                            "variant": variant,
                            "missing_kind": None,
                        }
                    )
                    if variant == "thumbnail":
                        warnings.append(f"{message_id}: only a thumbnail is present in the backup")

        raw_reply = local_core.clean_text(parameters.get("message_relation_server_message_id"))
        reply_to = server_to_message.get(raw_reply) if raw_reply else None
        excluded = is_context_only or bool(local_core.BOT_NAME_RE.search(sender_name or ""))
        messages.append(
            {
                "message_id": message_id,
                "timestamp": message_timestamp(row["created_time"], zone),
                "sender_name": sender_name,
                "sender_id": sender_mid,
                "role": role,
                "text": text,
                "reply_to_message_id": reply_to,
                "mentions": parameter_mentions(parameters, text, names),
                "media": media,
                "source_file": source_files[0],
                "source_sequence": sequence,
                "excluded_from_accounting": excluded,
            }
        )

    return (
        {
            "group_key": f"line:{group_mid}",
            "platform": "LINE",
            "timestamp_basis": "epoch",
            "source_timezone": None,
            "group_name": group_name,
            "source_files": source_files,
            "messages": messages,
            "rate_candidates": [],
        },
        warnings,
    )


def descriptor_for_backup(backup: Path) -> Path | None:
    candidate = backup.resolve().parent / "descript.xml"
    return candidate if candidate.is_file() else None


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        zone = ZoneInfo(args.timezone)
        if args.timezone != "Asia/Bangkok":
            raise ValueError("accounting timezone must be Asia/Bangkok")
        backup = find_backup_file(args.backup)
        roster = local_core.load_roster(args.roster.resolve())
        with open_line_databases(backup) as (line_db, contact_db):
            groups = available_groups(line_db)
            chosen = selected_groups(groups, args)
            if args.list_groups:
                for group in groups:
                    first = message_timestamp(group["first_timestamp"], zone) if group["first_timestamp"] else ""
                    last = message_timestamp(group["last_timestamp"], zone) if group["last_timestamp"] else ""
                    marker = "*" if group in chosen else " "
                    print(
                        f"{marker}\t{group['group_mid']}\t{group['group_name']}\t"
                        f"{group['message_count']}\t{first}\t{last}"
                    )
                return 0
            if args.output is None:
                raise ValueError("--output is required unless --list-groups is used")
            if not chosen:
                raise ValueError("No LINE groups matched the selection")

            output = args.output.resolve()
            if output.exists() and not args.force:
                raise ValueError(f"Output already exists (use --force): {output}")
            media_root = (args.media_dir or output.parent / f"{output.stem}_media").resolve()
            source_paths = [backup]
            descriptor = descriptor_for_backup(backup)
            if descriptor is not None:
                source_paths.append(descriptor)
            source_files = [str(path.resolve()) for path in source_paths]
            names = contact_names(contact_db)
            normalized_groups: list[dict[str, Any]] = []
            warnings: list[str] = []
            with open_backup_tar(backup) as archive:
                members = _member_index(archive)
                for group in chosen:
                    normalized, local_warnings = normalize_group(
                        line_db=line_db,
                        archive=archive,
                        members=members,
                        group=group,
                        zone=zone,
                        roster=roster,
                        names=names,
                        media_root=media_root,
                        force=args.force,
                        source_files=source_files,
                    )
                    normalized_groups.append(normalized)
                    warnings.extend(local_warnings)
    except (OSError, UnicodeError, ValueError, sqlite3.DatabaseError, tarfile.TarError, ZoneInfoNotFoundError) as exc:
        print(f"LINE Android extraction failed: {exc}", file=sys.stderr)
        return 2

    normalized_groups.sort(key=lambda item: item["group_key"])
    selection = {
        "extractor_version": EXTRACTOR_VERSION,
        "timezone": args.timezone,
        "group_ids": [item["group_key"] for item in normalized_groups],
        "roster": roster,
    }
    try:
        source_fingerprint = local_core.fingerprint_files(source_paths, context=selection)
        all_messages = [message for group in normalized_groups for message in group["messages"]]
        result = {
            "contract_version": local_core.NORMALIZED_CONTRACT,
            "timezone": args.timezone,
            "source_timezone": None,
            "source_fingerprint": source_fingerprint,
            "groups": normalized_groups,
            "statistics": {
                "groups": len(normalized_groups),
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
                "rate_candidates": 0,
                "source_files": len(source_files),
                "fingerprinted_files": len(source_paths),
            },
            "warnings": warnings,
            "skipped_files": [],
        }
        local_core.atomic_json(output, result)
        local_core.load_normalized(output)
    except (OSError, ValueError) as exc:
        print(f"LINE Android normalization failed: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result["statistics"], ensure_ascii=False))
    if warnings:
        print(f"Warnings: {len(warnings)} (see normalized output)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
