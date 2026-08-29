#!/usr/bin/env python3
"""Extract selected LINE group chats from an unencrypted Apple backup.

The output follows this skill's self-contained normalized contract. The script
is read-only with respect to the backup. It optionally copies selected media
next to the output so hashed iOS-backup filenames become ordinary image paths.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import plistlib
import re
import shutil
import sqlite3
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import core as local_core


LINE_DOMAINS = (
    "AppDomain-jp.naver.line",
    "AppDomainGroup-group.com.linecorp.line",
)
EXTRACTOR_VERSION = "1.1.0"
IMAGE_CONTENT_TYPES = {1, 112}
ALBUM_CONTAINER_CONTENT_TYPE = 111
SYSTEM_CONTENT_TYPE = 18
IGNORED_MEDIA_CONTENT_TYPES = {2, 3, 6, 7, 8, 9, 10, 13}
DEFAULT_GROUP_PATTERNS = (r"小额", r"钱庄", r"出[🐱😺🐷]")
WINDOWS_ILLEGAL = re.compile(r"[<>:\"/\\|?*\x00-\x1f]")
MENTION_RE = re.compile(r"@([^@\n\r]{1,64}?)(?=\s{2,}|$|[，。；、])")
CONTEXT_ONLY_NOTIFICATION_RE = re.compile(
    r"(?:撤回了一条(?:消息|信息)|已将.+添加至群|加入(?:了)?群聊|退出(?:了)?群聊|已离开群聊)"
)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("backup", type=Path, help="Backup folder or device folder containing Manifest.db")
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
    parser.add_argument("--self-name", default="LINE_SELF", help="Display name for outgoing messages with no sender row")
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
        help="Copied media directory (default: <output-stem>_media beside output)",
    )
    parser.add_argument("--no-copy-media", action="store_true", help="Reference hashed backup files directly")
    parser.add_argument("--force", action="store_true", help="Replace output and copied media files")
    return parser.parse_args(argv)


def sqlite_ro(path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(f"{path.resolve().as_uri()}?mode=ro&immutable=1", uri=True)
    connection.row_factory = sqlite3.Row
    return connection


def find_device_dir(input_path: Path) -> Path:
    path = input_path.resolve()
    if path.is_file() and path.name.casefold() == "manifest.db":
        return path.parent
    if (path / "Manifest.db").is_file():
        return path
    matches = sorted(path.glob("*/Manifest.db"), key=lambda item: str(item).casefold())
    if len(matches) == 1:
        return matches[0].parent
    if not matches:
        raise ValueError(f"No Manifest.db found under {path}")
    raise ValueError(f"Multiple device backups found under {path}; pass one device folder")


def physical_backup_file(device_dir: Path, file_id: str) -> Path:
    candidates = (device_dir / file_id[:2] / file_id, device_dir / file_id)
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    return candidates[0].resolve()


def manifest_row_for_suffix(
    manifest: sqlite3.Connection, suffix: str, *, domains: Iterable[str] = LINE_DOMAINS
) -> sqlite3.Row:
    placeholders = ",".join("?" for _ in domains)
    rows = list(
        manifest.execute(
            f"""
            SELECT fileID, domain, relativePath, flags
            FROM Files
            WHERE domain IN ({placeholders}) AND relativePath LIKE ?
            ORDER BY CASE WHEN domain = 'AppDomainGroup-group.com.linecorp.line' THEN 0 ELSE 1 END,
                     relativePath
            """,
            (*domains, f"%{suffix}"),
        )
    )
    exact = [row for row in rows if str(row["relativePath"]).endswith(suffix)]
    if len(exact) != 1:
        details = ", ".join(str(row["relativePath"]) for row in exact[:5])
        raise ValueError(f"Expected one LINE backup file ending {suffix!r}; found {len(exact)}: {details}")
    return exact[0]


def safe_filename(value: str, *, fallback: str) -> str:
    cleaned = WINDOWS_ILLEGAL.sub("_", unicodedata.normalize("NFKC", value)).strip(" .")
    return (cleaned or fallback)[:80]


def clean_text(value: object) -> str:
    text = str(value or "")
    for marker in ("\u2066", "\u2067", "\u2068", "\u2069", "\ufeff"):
        text = text.replace(marker, "")
    return text.replace("\r\n", "\n").replace("\r", "\n").strip("\u3000 \t\n")


def message_timestamp(raw: object, zone: ZoneInfo) -> str:
    milliseconds = int(raw)
    return datetime.fromtimestamp(milliseconds / 1000, tz=timezone.utc).astimezone(zone).isoformat()


def resolve_keyed_archive(raw: bytes | None) -> object:
    """Best-effort decoder for LINE's NSKeyedArchiver metadata."""
    if not raw:
        return None
    try:
        archive = plistlib.loads(raw)
    except Exception:
        return None
    if not isinstance(archive, dict) or "$objects" not in archive or "$top" not in archive:
        return archive
    objects = archive.get("$objects")
    if not isinstance(objects, list):
        return archive
    resolving: set[int] = set()
    memo: dict[int, object] = {}

    def resolve(value: object) -> object:
        if isinstance(value, plistlib.UID):
            index = int(value.data)
            if index in memo:
                return memo[index]
            if index in resolving or not (0 <= index < len(objects)):
                return None
            resolving.add(index)
            result = resolve(objects[index])
            resolving.remove(index)
            memo[index] = result
            return result
        if isinstance(value, list):
            return [resolve(item) for item in value]
        if isinstance(value, dict):
            if "NS.keys" in value and "NS.objects" in value:
                keys = resolve(value["NS.keys"])
                values = resolve(value["NS.objects"])
                if isinstance(keys, list) and isinstance(values, list):
                    return {str(key): item for key, item in zip(keys, values)}
            if "NS.objects" in value and set(value).issubset({"NS.objects", "$class"}):
                return resolve(value["NS.objects"])
            return {
                str(key): resolve(item)
                for key, item in value.items()
                if key not in {"$class", "$classes", "$classname"}
            }
        if isinstance(value, bytes):
            return None
        return value

    top = archive.get("$top")
    if isinstance(top, dict) and "root" in top:
        return resolve(top["root"])
    return resolve(top)


def backup_source_timezone(manifest: sqlite3.Connection) -> str | None:
    row = manifest.execute(
        "SELECT file FROM Files WHERE domain = ? AND relativePath = ? LIMIT 1",
        ("DatabaseDomain", "timezone/localtime"),
    ).fetchone()
    metadata = resolve_keyed_archive(row["file"] if row is not None else None)
    target = metadata.get("Target") if isinstance(metadata, dict) else None
    marker = "/zoneinfo/"
    if isinstance(target, str) and marker in target:
        return target.split(marker, 1)[1]
    return None


def find_reply_id(metadata: object) -> str | None:
    if isinstance(metadata, list):
        for value in metadata:
            nested = find_reply_id(value)
            if nested:
                return nested
        return None
    if not isinstance(metadata, dict):
        return None
    for key, value in metadata.items():
        compact = re.sub(r"[^a-z]", "", str(key).casefold())
        if "reply" in compact and compact.endswith(("id", "messageid")):
            if isinstance(value, (str, int)) and str(value).strip():
                return str(value).strip()
        nested = find_reply_id(value)
        if nested:
            return nested
    return None


def extract_mentions(text: str) -> list[str]:
    mentions: list[str] = []
    for match in MENTION_RE.finditer(text):
        name = match.group(1).strip()
        if name and name not in mentions:
            mentions.append(name)
    return mentions


def attachment_rows(manifest: sqlite3.Connection, chat_mid: str) -> dict[str, list[sqlite3.Row]]:
    rows = list(
        manifest.execute(
            """
            SELECT fileID, domain, relativePath, flags
            FROM Files
            WHERE domain IN ('AppDomain-jp.naver.line', 'AppDomainGroup-group.com.linecorp.line')
              AND (
                relativePath LIKE ? OR relativePath LIKE ?
              )
            ORDER BY CASE WHEN relativePath LIKE '%/Message Attachments/%' THEN 0 ELSE 1 END,
                     relativePath
            """,
            (f"%/Message Attachments/{chat_mid}/%", f"%/Message Thumbnails/{chat_mid}/%"),
        )
    )
    by_message: dict[str, list[sqlite3.Row]] = {}
    for row in rows:
        filename = Path(str(row["relativePath"])).name
        message_id = filename.split(".", 1)[0]
        by_message.setdefault(message_id, []).append(row)
    return by_message


def mime_and_kind(relative_path: str, content_type: int) -> tuple[str | None, str]:
    suffix = Path(relative_path).suffix.casefold()
    if suffix == ".thumb" or content_type in IMAGE_CONTENT_TYPES:
        return "image/jpeg", "image"
    mime, _ = mimetypes.guess_type(relative_path)
    if mime and mime.startswith("image/"):
        return mime, "image"
    if mime == "application/pdf":
        return mime, "document"
    if mime and mime.startswith("video/"):
        return mime, "video"
    if mime and mime.startswith("audio/"):
        return mime, "audio"
    return mime, "file"


def select_attachment(
    rows: list[sqlite3.Row], device_dir: Path
) -> tuple[sqlite3.Row | None, Path | None]:
    for row in rows:
        candidate = physical_backup_file(device_dir, str(row["fileID"]))
        if candidate.is_file() and int(row["flags"] or 0) == 1:
            return row, candidate
    return (rows[0], None) if rows else (None, None)


def copied_media_path(
    source: Path,
    row: sqlite3.Row,
    *,
    media_root: Path,
    group_name: str,
    group_mid: str,
    message_id: str,
    force: bool,
) -> Path:
    relative = str(row["relativePath"])
    extension = Path(relative).suffix.casefold()
    if extension == ".thumb":
        extension = ".jpg"
    if not extension:
        extension = ".bin"
    group_dir = media_root / safe_filename(group_name, fallback=group_mid[:12])
    group_dir.mkdir(parents=True, exist_ok=True)
    destination = group_dir / f"{safe_filename(message_id, fallback='message')}{extension}"
    if destination.exists() and not force:
        if destination.stat().st_size != source.stat().st_size:
            raise ValueError(f"Media destination exists with different size: {destination}")
        return destination.resolve()
    shutil.copy2(source, destination)
    return destination.resolve()


def hash_files(paths: Iterable[Path], selection: object) -> str:
    digests: list[str] = []
    for path in sorted({item.resolve() for item in paths}, key=lambda item: str(item).casefold()):
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        digests.append(digest.hexdigest())
    digests.append(
        hashlib.sha256(
            json.dumps(selection, ensure_ascii=False, sort_keys=True).encode("utf-8")
        ).hexdigest()
    )
    return "sha256:" + hashlib.sha256("\n".join(sorted(digests)).encode("ascii")).hexdigest()


def available_groups(line_db: sqlite3.Connection, group_db: sqlite3.Connection) -> list[dict[str, Any]]:
    names = {
        str(row["ZID"]): str(row["ZNAME"] or row["ZID"])
        for row in group_db.execute("SELECT ZID, ZNAME FROM ZUNIFIEDGROUP")
        if row["ZID"]
    }
    rows = list(
        line_db.execute(
            """
            SELECT c.Z_PK, c.ZMID, c.ZTYPE, COUNT(m.Z_PK) AS message_count,
                   MIN(m.ZTIMESTAMP) AS first_timestamp, MAX(m.ZTIMESTAMP) AS last_timestamp
            FROM ZCHAT c
            LEFT JOIN ZMESSAGE m ON m.ZCHAT = c.Z_PK
            WHERE c.ZTYPE = 2 OR c.ZMID LIKE 'c%'
            GROUP BY c.Z_PK
            ORDER BY c.ZMID
            """
        )
    )
    return [
        {
            "chat_pk": int(row["Z_PK"]),
            "group_mid": str(row["ZMID"]),
            "group_name": names.get(str(row["ZMID"]), str(row["ZMID"])),
            "message_count": int(row["message_count"] or 0),
            "first_timestamp": row["first_timestamp"],
            "last_timestamp": row["last_timestamp"],
        }
        for row in rows
        if row["ZMID"]
    ]


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


def normalize_group(
    *,
    line_db: sqlite3.Connection,
    manifest: sqlite3.Connection,
    device_dir: Path,
    group: dict[str, Any],
    zone: ZoneInfo,
    core: Any,
    roster: dict[str, tuple[str, ...]],
    self_name: str,
    media_root: Path | None,
    force: bool,
    source_files: list[str],
    source_timezone: str | None,
) -> tuple[dict[str, Any], list[str], list[Path]]:
    group_mid = group["group_mid"]
    group_name = group["group_name"]
    attachments = attachment_rows(manifest, group_mid)
    users = {
        int(row["Z_PK"]): (str(row["ZMID"] or ""), str(row["ZNAME"] or row["ZMID"] or ""))
        for row in line_db.execute("SELECT Z_PK, ZMID, ZNAME FROM ZUSER")
    }
    warnings: list[str] = []
    fingerprint_paths: list[Path] = []
    rows = list(
        line_db.execute(
            """
            SELECT Z_PK, ZID, ZTIMESTAMP, ZCONTENTTYPE, ZSENDSTATUS,
                   ZTEXT, ZSENDER, ZCONTENTMETADATA
            FROM ZMESSAGE
            WHERE ZCHAT = ?
            ORDER BY ZTIMESTAMP, Z_PK
            """,
            (group["chat_pk"],),
        )
    )
    messages: list[dict[str, Any]] = []
    snapshot_id = core.sha256_file(Path(source_files[1]))[:12]
    for sequence, row in enumerate(rows, start=1):
        attachment_local_id = str(row["ZID"] if row["ZID"] is not None else f"pk-{row['Z_PK']}")
        local_id = str(
            row["ZID"]
            if row["ZID"] is not None
            else f"snapshot-{snapshot_id}-pk-{row['Z_PK']}"
        )
        message_id = f"line:{group_mid}:{local_id}"
        content_type = int(row["ZCONTENTTYPE"] or 0)
        text = clean_text(row["ZTEXT"])
        sender_pk = row["ZSENDER"]
        send_status = int(row["ZSENDSTATUS"] or 0)
        is_system = content_type == SYSTEM_CONTENT_TYPE or (
            sender_pk is None
            and (send_status != 1 or bool(CONTEXT_ONLY_NOTIFICATION_RE.search(text)))
        )
        if is_system:
            sender_id = None
            sender_name = None
            role = "未知"
        elif sender_pk is None and send_status == 1:
            sender_id = "line:self"
            sender_name = self_name
            role = "内部人员"
        elif sender_pk is None:
            sender_id = None
            sender_name = None
            role = "未知"
        else:
            sender_id, sender_name = users.get(int(sender_pk), (f"line:user-pk-{sender_pk}", ""))
            role = core.classify_role(sender_name or sender_id, roster=roster)

        metadata = resolve_keyed_archive(row["ZCONTENTMETADATA"])
        raw_reply = find_reply_id(metadata)
        reply_to = f"line:{group_mid}:{raw_reply}" if raw_reply else None
        media: list[dict[str, Any]] = []
        candidates = attachments.get(attachment_local_id, [])
        selected_row, selected_path = select_attachment(candidates, device_dir)
        should_have_image = content_type in IMAGE_CONTENT_TYPES
        if selected_row is not None and selected_path is not None:
            mime_type, kind = mime_and_kind(str(selected_row["relativePath"]), content_type)
            if kind in {"image", "document"} and content_type not in IGNORED_MEDIA_CONTENT_TYPES:
                final_path = selected_path
                if media_root is not None:
                    final_path = copied_media_path(
                        selected_path,
                        selected_row,
                        media_root=media_root,
                        group_name=group_name,
                        group_mid=group_mid,
                        message_id=local_id,
                        force=force,
                    )
                media.append(
                    {
                        "media_id": f"{message_id}#media:0",
                        "kind": kind,
                        "mime_type": mime_type,
                        "availability": "available",
                        "source_field": (
                            "line_attachment"
                            if "/Message Attachments/" in str(selected_row["relativePath"])
                            else "line_thumbnail"
                        ),
                        "path": str(final_path),
                        "original_reference": str(selected_row["relativePath"]),
                        "blob_sha256": None,
                        "byte_size": selected_path.stat().st_size,
                        "variant": (
                            "primary"
                            if "/Message Attachments/" in str(selected_row["relativePath"])
                            else "thumbnail"
                        ),
                        "missing_kind": None,
                    }
                )
        elif should_have_image:
            media.append(
                {
                    "media_id": f"{message_id}#media:0",
                    "kind": "image",
                    "mime_type": None,
                    "availability": "missing",
                    "source_field": "line_attachment",
                    "path": None,
                    "original_reference": f"LINE message {local_id}",
                    "blob_sha256": None,
                    "byte_size": None,
                    "variant": None,
                    "missing_kind": "missing_from_backup",
                }
            )
            warnings.append(f"{message_id}: image attachment is missing from the backup")

        excluded = (
            is_system
            or content_type == ALBUM_CONTAINER_CONTENT_TYPE
            or bool(core.BOT_NAME_RE.search(sender_name or ""))
        )
        messages.append(
            {
                "message_id": message_id,
                "timestamp": message_timestamp(row["ZTIMESTAMP"], zone),
                "sender_name": sender_name,
                "sender_id": sender_id,
                "role": role,
                "text": text,
                "reply_to_message_id": reply_to,
                "mentions": extract_mentions(text),
                "media": media,
                "source_file": source_files[1],
                "source_sequence": sequence,
                "excluded_from_accounting": excluded,
            }
        )

    return (
        {
            "group_key": f"line:{group_mid}",
            "platform": "LINE",
            "timestamp_basis": "epoch",
            "source_timezone": source_timezone,
            "group_name": group_name,
            "source_files": source_files,
            "messages": messages,
            "rate_candidates": [],
        },
        warnings,
        fingerprint_paths,
    )


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        zone = ZoneInfo(args.timezone)
        device_dir = find_device_dir(args.backup)
        manifest_path = device_dir / "Manifest.db"
        manifest = sqlite_ro(manifest_path)
        source_timezone = backup_source_timezone(manifest)
        line_row = manifest_row_for_suffix(manifest, "/Messages/Line.sqlite")
        group_row = manifest_row_for_suffix(manifest, "/Messages/UnifiedGroup.sqlite")
        line_path = physical_backup_file(device_dir, str(line_row["fileID"]))
        group_path = physical_backup_file(device_dir, str(group_row["fileID"]))
        if not line_path.is_file() or not group_path.is_file():
            raise ValueError("LINE database files listed in Manifest.db are missing")
        core = local_core
        roster = core.load_roster(args.roster.resolve())
        line_db = sqlite_ro(line_path)
        group_db = sqlite_ro(group_path)
        groups = available_groups(line_db, group_db)
        chosen = selected_groups(groups, args)
    except (OSError, ValueError, sqlite3.DatabaseError, ZoneInfoNotFoundError) as exc:
        print(f"LINE extraction failed: {exc}", file=sys.stderr)
        return 2

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
        print("--output is required unless --list-groups is used", file=sys.stderr)
        return 2
    if not chosen:
        print("No LINE groups matched the selection", file=sys.stderr)
        return 2

    output = args.output.resolve()
    if output.exists() and not args.force:
        print(f"Output already exists (use --force): {output}", file=sys.stderr)
        return 2
    media_root = None
    if not args.no_copy_media:
        media_root = (args.media_dir or output.parent / f"{output.stem}_media").resolve()

    source_files = [str(manifest_path.resolve()), str(line_path), str(group_path)]
    normalized_groups: list[dict[str, Any]] = []
    warnings: list[str] = []
    fingerprint_paths: list[Path] = [manifest_path, line_path, group_path]
    try:
        for group in chosen:
            normalized, local_warnings, local_paths = normalize_group(
                line_db=line_db,
                manifest=manifest,
                device_dir=device_dir,
                group=group,
                zone=zone,
                core=core,
                roster=roster,
                self_name=args.self_name,
                media_root=media_root,
                force=args.force,
                source_files=source_files,
                source_timezone=source_timezone,
            )
            normalized_groups.append(normalized)
            warnings.extend(local_warnings)
            fingerprint_paths.extend(local_paths)
        normalized_groups.sort(key=lambda item: item["group_key"])
        selection = {
            "extractor_version": EXTRACTOR_VERSION,
            "timezone": args.timezone,
            "group_ids": [item["group_key"] for item in normalized_groups],
            "roster": core.load_roster(args.roster.resolve()),
            "self_name": args.self_name,
            "source_timezone": source_timezone,
        }
        source_fingerprint = hash_files(fingerprint_paths, selection)
    except (OSError, ValueError, sqlite3.DatabaseError) as exc:
        print(f"LINE normalization failed: {exc}", file=sys.stderr)
        return 2

    all_messages = [message for group in normalized_groups for message in group["messages"]]
    result = {
        "contract_version": core.NORMALIZED_CONTRACT,
        "timezone": args.timezone,
        "source_timezone": source_timezone,
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
            "rate_candidates": sum(len(group["rate_candidates"]) for group in normalized_groups),
            "source_files": len(source_files),
            "fingerprinted_files": len({path.resolve() for path in fingerprint_paths}),
        },
        "warnings": warnings,
        "skipped_files": [],
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    core.load_normalized(output)
    print(json.dumps(result["statistics"], ensure_ascii=False))
    if warnings:
        print(f"Warnings: {len(warnings)} (see normalized.json)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
