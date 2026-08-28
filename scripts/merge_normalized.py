#!/usr/bin/env python3
"""Merge compatible normalized chat documents without doing any accounting."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any

import core


MERGER_VERSION = "1.0.0"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", type=Path, help="normalized JSON documents")
    parser.add_argument("-o", "--output", required=True, type=Path)
    parser.add_argument("--timezone", default="Asia/Bangkok", help="Required common accounting timezone")
    parser.add_argument("--force", action="store_true")
    return parser.parse_args(argv)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.parent / f".{path.name}.tmp"
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    output = args.output.resolve()
    if output.exists() and not args.force:
        print(f"Output already exists (use --force): {output}", file=sys.stderr)
        return 2

    documents: list[tuple[Path, dict[str, Any], str]] = []
    try:
        for raw_path in args.inputs:
            path = raw_path.resolve()
            require(path != output, "output cannot also be an input")
            data = core.load_normalized(path)
            require(data.get("timezone") == args.timezone, f"{path}: timezone must be {args.timezone}")
            documents.append((path, data, file_sha256(path)))

        contract_versions = {data["contract_version"] for _, data, _ in documents}
        require(len(contract_versions) == 1, "all inputs must use the same contract_version")
        groups: list[dict[str, Any]] = []
        group_keys: set[str] = set()
        warnings: list[str] = []
        skipped_files: list[object] = []
        for path, data, _ in documents:
            for group in data["groups"]:
                require(isinstance(group, dict), f"{path}: each group must be an object")
                key = group.get("group_key")
                require(isinstance(key, str) and key, f"{path}: group_key is required")
                require(key not in group_keys, f"duplicate group_key across inputs: {key}")
                group_keys.add(key)
                groups.append(group)
            warnings.extend(str(item) for item in data.get("warnings", []))
            skipped_files.extend(data.get("skipped_files", []))

        groups.sort(key=lambda item: (str(item.get("platform", "")), str(item.get("group_name", "")), item["group_key"]))
        fingerprint_material = [
            f"{digest}:{data['source_fingerprint']}" for _, data, digest in documents
        ]
        fingerprint_material.extend((f"timezone:{args.timezone}", f"merger:{MERGER_VERSION}"))
        source_fingerprint = "sha256:" + hashlib.sha256(
            "\n".join(sorted(fingerprint_material)).encode("utf-8")
        ).hexdigest()
        messages = [message for group in groups for message in group.get("messages", [])]
        source_files = {
            str(source)
            for group in groups
            for source in group.get("source_files", [])
            if source is not None
        }
        result = {
            "contract_version": next(iter(contract_versions)),
            "timezone": args.timezone,
            "source_timezones": {
                path.name: data.get("source_timezone")
                for path, data, _ in documents
                if data.get("source_timezone")
            },
            "source_fingerprint": source_fingerprint,
            "groups": groups,
            "statistics": {
                "groups": len(groups),
                "messages": len(messages),
                "excluded_messages": sum(bool(message.get("excluded_from_accounting")) for message in messages),
                "media": sum(len(message.get("media", [])) for message in messages),
                "available_media": sum(
                    media.get("availability") == "available"
                    for message in messages
                    for media in message.get("media", [])
                ),
                "missing_media": sum(
                    media.get("availability") == "missing"
                    for message in messages
                    for media in message.get("media", [])
                ),
                "rate_candidates": sum(len(group.get("rate_candidates", [])) for group in groups),
                "source_files": len(source_files),
                "fingerprinted_files": sum(
                    int(data.get("statistics", {}).get("fingerprinted_files", 0))
                    for _, data, _ in documents
                ),
                "normalized_inputs": len(documents),
            },
            "warnings": warnings,
            "skipped_files": skipped_files,
        }
        atomic_json(output, result)
        core.load_normalized(output)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"Merge failed: {exc}", file=sys.stderr)
        return 2

    print(json.dumps(result["statistics"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
