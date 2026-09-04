#!/usr/bin/env python3
"""Run optional local OCR sequentially and return non-authoritative text hints."""

from __future__ import annotations

import contextlib
import importlib.metadata
import io
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Callable


WORKER_CONTRACT = "group-chat-ocr-worker/1.0"
CANDIDATE_CONTRACT = "group-chat-ocr-candidate/1.0"
MAX_REQUESTS = 20
IMAGE_TIMEOUT_SECONDS = 15


def _package_version(name: str) -> str | None:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return None


def _clean_lines(values: list[tuple[str, float | None]]) -> list[tuple[str, float | None]]:
    result: list[tuple[str, float | None]] = []
    for raw_text, confidence in values:
        text = " ".join(str(raw_text).replace("\x00", " ").split())
        if not text:
            continue
        normalized_confidence = None
        if confidence is not None:
            try:
                normalized_confidence = min(1.0, max(0.0, float(confidence)))
            except (TypeError, ValueError):
                normalized_confidence = None
        result.append((text, normalized_confidence))
    return result


def _rapidocr_backend() -> tuple[str, str | None, Callable[[Path], list[tuple[str, float | None]]]] | None:
    try:
        noise = io.StringIO()
        with contextlib.redirect_stdout(noise), contextlib.redirect_stderr(noise):
            from rapidocr_onnxruntime import RapidOCR

            engine = RapidOCR()
    except Exception:
        return None

    def recognize(path: Path) -> list[tuple[str, float | None]]:
        noise = io.StringIO()
        with contextlib.redirect_stdout(noise), contextlib.redirect_stderr(noise):
            output = engine(str(path))
        raw_result = output[0] if isinstance(output, tuple) else output
        if not isinstance(raw_result, list):
            return []
        lines: list[tuple[str, float | None]] = []
        for item in raw_result:
            if not isinstance(item, (list, tuple)) or len(item) < 2:
                continue
            confidence = item[2] if len(item) > 2 else None
            lines.append((str(item[1]), confidence))
        return _clean_lines(lines)

    return "rapidocr_onnxruntime", _package_version("rapidocr-onnxruntime"), recognize


def _tesseract_backend() -> tuple[str, str | None, Callable[[Path], list[tuple[str, float | None]]]] | None:
    executable = shutil.which("tesseract")
    if not executable:
        return None

    def recognize(path: Path) -> list[tuple[str, float | None]]:
        completed = subprocess.run(
            [executable, str(path), "stdout", "--psm", "6"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=IMAGE_TIMEOUT_SECONDS,
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError("tesseract returned a non-zero status")
        return _clean_lines([(line, None) for line in completed.stdout.splitlines()])

    return "tesseract", None, recognize


def _select_backend() -> tuple[str, str | None, Callable[[Path], list[tuple[str, float | None]]]] | None:
    return _rapidocr_backend() or _tesseract_backend()


def _candidate(
    *,
    digest: str,
    backend: str,
    backend_version: str | None,
    lines: list[tuple[str, float | None]],
    elapsed_ms: int,
    text_limit: int,
    error: str | None = None,
) -> dict[str, Any]:
    text = "\n".join(line for line, _ in lines)
    truncated = len(text) > text_limit
    if truncated:
        text = text[:text_limit].rstrip()
    confidences = [value for _, value in lines if value is not None]
    return {
        "contract_version": CANDIDATE_CONTRACT,
        "content_sha256": digest,
        "status": "error" if error else "ok" if text else "empty",
        "backend": backend,
        "backend_version": backend_version,
        "authoritative": False,
        "text": text,
        "average_confidence": (
            sum(confidences) / len(confidences) if confidences else None
        ),
        "line_count": len(lines),
        "elapsed_ms": max(0, elapsed_ms),
        "truncated": truncated,
        **({"error": error[:300]} if error else {}),
    }


def process(payload: object) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("worker input must be an object")
    if payload.get("contract_version") != WORKER_CONTRACT:
        raise ValueError("unsupported worker contract")
    if payload.get("candidate_contract") != CANDIDATE_CONTRACT:
        raise ValueError("unsupported OCR candidate contract")
    text_limit = payload.get("text_limit", 800)
    if not isinstance(text_limit, int) or isinstance(text_limit, bool) or not 1 <= text_limit <= 4000:
        raise ValueError("text_limit must be 1..4000")
    requests = payload.get("requests")
    if not isinstance(requests, list) or len(requests) > MAX_REQUESTS:
        raise ValueError(f"requests must be a list with at most {MAX_REQUESTS} items")
    normalized_requests: list[tuple[str, Path]] = []
    seen: set[str] = set()
    for item in requests:
        if not isinstance(item, dict):
            raise ValueError("each OCR request must be an object")
        digest = str(item.get("content_sha256") or "")
        path = Path(str(item.get("path") or "")).resolve()
        if len(digest) != 64 or any(character not in "0123456789abcdef" for character in digest):
            raise ValueError("OCR request has an invalid content hash")
        if digest in seen:
            continue
        if not path.is_file():
            raise ValueError("OCR request path is unavailable")
        seen.add(digest)
        normalized_requests.append((digest, path))

    backend = _select_backend()
    if backend is None:
        return {
            "contract_version": WORKER_CONTRACT,
            "status": "unavailable",
            "reason": "no_supported_local_ocr_backend",
            "backend": None,
            "backend_version": None,
            "results": {},
        }
    backend_name, backend_version, recognize = backend
    results: dict[str, Any] = {}
    for digest, path in normalized_requests:
        started = time.perf_counter()
        try:
            lines = recognize(path)
            results[digest] = _candidate(
                digest=digest,
                backend=backend_name,
                backend_version=backend_version,
                lines=lines,
                elapsed_ms=round((time.perf_counter() - started) * 1000),
                text_limit=text_limit,
            )
        except Exception as exc:
            results[digest] = _candidate(
                digest=digest,
                backend=backend_name,
                backend_version=backend_version,
                lines=[],
                elapsed_ms=round((time.perf_counter() - started) * 1000),
                text_limit=text_limit,
                error=f"{type(exc).__name__}: {str(exc)}",
            )
    return {
        "contract_version": WORKER_CONTRACT,
        "status": "ready",
        "reason": None,
        "backend": backend_name,
        "backend_version": backend_version,
        "results": results,
    }


def main() -> int:
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    try:
        payload = json.loads(sys.stdin.read())
        noise = io.StringIO()
        with contextlib.redirect_stdout(noise), contextlib.redirect_stderr(noise):
            response = process(payload)
    except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
        print(json.dumps({"contract_version": WORKER_CONTRACT, "status": "error", "reason": str(exc), "results": {}}))
        return 2
    print(json.dumps(response, ensure_ascii=False, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
