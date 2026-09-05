#!/usr/bin/env python3
"""Small shared primitives for the simple group-chat ledger."""

from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal, InvalidOperation, getcontext
from pathlib import Path
from typing import Any, Iterable, Mapping

try:
    import yaml
except ImportError:  # The roster subset also has a small stdlib fallback parser.
    yaml = None


NORMALIZED_CONTRACT = "small-group-normalized/1.0"
ORDERS_CONTRACT = "small-group-simple-orders/2.0"
RULE_VERSION = "simple-ledger/2.0"
getcontext().prec = 60

HEADERS = [
    "记录类型",
    "订单编号",
    "客户昵称",
    "换汇方向",
    "付款合计",
    "汇率",
    "收款方实际到账金额",
    "流水币种",
    "应回金额",
    "内部实际回款合计",
    "核对结果",
    "备注",
    "订单状态",
    "完成时间",
    "收款方",
    "聊天消息时间",
]

EVENT_TYPES = {
    "payment_screenshot",
    "payout_screenshot",
    "cash_payment",
    "cash_payout",
    "payout_recovery",
    "irrelevant_media",
    "missing_evidence_media",
}
FUND_EVENT_TYPES = {
    "payment_screenshot",
    "payout_screenshot",
    "cash_payment",
    "cash_payout",
    "payout_recovery",
}
MEDIA_EVENT_TYPES = EVENT_TYPES
NON_EVIDENCE_MEDIA_KINDS = {"audio", "video", "sticker", "animation"}

CURRENCY_TOLERANCES = {
    "CNY": Decimal("0"),
    "THB": Decimal("0"),
    "USDT": Decimal("0"),
    "TRX": Decimal("0"),
    "USD": Decimal("0"),
    "JPY": Decimal("0"),
    "GBP": Decimal("0"),
    "SGD": Decimal("0"),
}
CURRENCIES = frozenset(CURRENCY_TOLERANCES)
DIRECTIONS = {
    f"{payment}->{payout}": (payment, payout)
    for payment in CURRENCIES
    for payout in CURRENCIES
    if payment != payout
}

FAILED_STATUS_RE = re.compile(
    r"(?:失败|未完成|无法完成|姓名不匹配|已取消|取消成功|已退回|退回成功|不支持|暂时不支持|受限|限额|超限|"
    r"风控|风险管理|declined|failed|failure|cancelled|canceled|reverted|returned)",
    re.IGNORECASE,
)
FAILED_NEGATED_COMPLETION_RE = re.compile(
    r"(?:未成功|(?:无法|不能|未能|不(?:会|能)?|not\s+)(?:[^\n]{0,12})?"
    r"(?:成功|完成|到账|successful|completed|received))",
    re.IGNORECASE,
)
PENDING_STATUS_RE = re.compile(
    r"(?:待区块|等待确认|待确认|正在确认|确认中|处理中|处理当中|正在进行|进行中|等待中|"
    r"pending|processing|confirming|unconfirmed|not\s+confirmed)",
    re.IGNORECASE,
)
PENDING_NEGATED_COMPLETION_RE = re.compile(
    r"(?:尚未|还未|未|没有|没|等待|待)(?:[^\n]{0,12})?(?:到账|入账|收到|确认|完成|"
    r"credited|received|confirmed|completed)",
    re.IGNORECASE,
)
COMPLETED_STATUS_RE = re.compile(
    r"(?:支付成功|转账成功|交易成功|已完成|成功|已到账|到账|completed|successful|success)",
    re.IGNORECASE,
)
STATUS_CLASSES = frozenset({"failed", "pending", "completed", "blank", "unknown"})
CONFIDENCE_LEVELS = frozenset({"high", "medium", "low"})
AMOUNT_COMPLETENESS = frozenset({"complete", "partial", "unreadable"})
UNKNOWN_PAYEES = frozenset({"未显示", "无法辨认"})
PAYEE_STATES = frozenset({"visible", "not_shown", "unreadable", "cash"})
GENERIC_PAYEES = frozenset(
    {
        "群内收款方",
        "泰铢收款账户",
        "usdt收款钱包",
        "银行卡收款方",
        "客户退款钱包",
    }
)
GENERIC_PAYEE_RE = re.compile(
    r"^(?:截图所示|聊天指定|固定金额二维码|支付宝截图所示).*(?:收款方|收款账户|收款地址|账户)?$",
    re.IGNORECASE,
)
THAI_BANK_ACCOUNT_RE = re.compile(r"^[0-9Xx*#•·×.\- –—]+$")
THAI_BANK_ACCOUNT_TOKEN_RE = re.compile(r"[0-9Xx*#•×]")
BOT_NAME_RE = re.compile(r"(?:自动统计机器人|统计机器人|机器人A\d+)", re.IGNORECASE)
CASH_AMOUNT_NOTATION_RE = re.compile(
    r"(?P<sign>[+-])?\s*(?:(?P<number>(?:\d+(?:\.\d*)?|\.\d+))\s*)?(?P<suffix>[wk])",
    re.IGNORECASE,
)
CASH_AMOUNT_MULTIPLIERS = {"w": Decimal("10000"), "k": Decimal("1000")}


def media_is_evidence(media: Mapping[str, Any]) -> bool:
    return str(media.get("kind") or "").casefold() not in NON_EVIDENCE_MEDIA_KINDS


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def sha256_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_cached_file_hash(
    path: Path,
    recorded_hash: object,
    *,
    verify: bool,
    changed_message: str,
) -> tuple[str, bool]:
    """Return a trusted prior hash or verify the file and return a fresh hash."""

    recorded = clean_text(recorded_hash)
    if not verify:
        return recorded, False
    actual = sha256_file(path)
    if recorded:
        require(recorded == actual, changed_message)
    return actual, True


def fingerprint_files(paths: Iterable[Path], *, context: object) -> str:
    material: list[str] = []
    for path in sorted({item.resolve() for item in paths}, key=lambda item: str(item).casefold()):
        material.append(f"{path.name}:{sha256_file(path)}")
    material.append(f"context:{sha256_bytes(json_bytes(context))}")
    return "sha256:" + sha256_bytes("\n".join(material).encode("utf-8"))


def fingerprint_json(value: object) -> str:
    return "sha256:" + sha256_bytes(json_bytes(value))


def atomic_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.parent / f".{path.name}.tmp"
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def normalize_name(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    text = text.replace("～", "~").replace("〜", "~")
    return "".join(
        character
        for character in text
        if not character.isspace() and unicodedata.category(character) != "Cf"
    )


def load_roster(path: Path) -> dict[str, tuple[dict[str, str], ...]]:
    raw_text = path.read_text(encoding="utf-8-sig")
    if yaml is not None:
        data = yaml.safe_load(raw_text) or {}
        raw_patterns = data.get("staff", {}).get("patterns", [])
    else:
        raw_patterns = []
        current: dict[str, str] | None = None
        for raw_line in raw_text.splitlines():
            line = raw_line.split("#", 1)[0].strip()
            if line.startswith("- kind:"):
                if current:
                    raw_patterns.append(current)
                current = {"kind": line.split(":", 1)[1].strip().strip("\"'")}
            elif current is not None and line.startswith("value:"):
                current["value"] = line.split(":", 1)[1].strip().strip("\"'")
        if current:
            raw_patterns.append(current)
    require(isinstance(raw_patterns, list), f"{path}: staff.patterns must be a list")
    patterns: list[dict[str, str]] = []
    for index, item in enumerate(raw_patterns):
        require(isinstance(item, dict), f"{path}: staff.patterns[{index}] must be an object")
        kind = str(item.get("kind") or "")
        value = normalize_name(item.get("value"))
        require(kind in {"prefix", "contains", "exact"}, f"{path}: unsupported roster kind {kind!r}")
        require(bool(value), f"{path}: empty roster value")
        patterns.append({"kind": kind, "value": value})
    require(bool(patterns), f"{path}: at least one staff pattern is required")
    return {"patterns": tuple(patterns)}


def classify_role(name: object, *, roster: Mapping[str, Any], is_self: bool = False) -> str:
    if is_self:
        return "内部人员"
    normalized = normalize_name(name)
    if not normalized:
        return "未知"
    for pattern in roster.get("patterns", ()):
        kind = pattern["kind"]
        value = pattern["value"]
        if (
            (kind == "prefix" and normalized.startswith(value))
            or (kind == "contains" and value in normalized)
            or (kind == "exact" and normalized == value)
        ):
            return "内部人员"
    if BOT_NAME_RE.search(str(name or "")):
        return "未知"
    return "客户候选"


def clean_text(value: object) -> str:
    text = str(value or "")
    for marker in (
        "\u200e",
        "\u200f",
        "\u202a",
        "\u202b",
        "\u202c",
        "\u2066",
        "\u2067",
        "\u2068",
        "\u2069",
        "\ufeff",
    ):
        text = text.replace(marker, "")
    return text.replace("\r\n", "\n").replace("\r", "\n").strip("\u3000 \t\n")


def validate_payee(
    value: object,
    *,
    field: str,
    cash: bool = False,
    payee_state: object = None,
    require_state: bool = False,
) -> str:
    """Return a payee that is consistent with the reviewed evidence state."""
    state = clean_text(payee_state).casefold()
    if require_state:
        require(bool(state), f"{field.rsplit('.', 1)[0]}.payee_state is required")
    if state:
        require(state in PAYEE_STATES, f"{field.rsplit('.', 1)[0]}.payee_state is unsupported")

    if cash:
        if state:
            require(state == "cash", f"{field.rsplit('.', 1)[0]}.payee_state must be cash")
        return "现金"

    require(state != "cash", f"{field.rsplit('.', 1)[0]}.payee_state cash requires kind cash")
    require(
        isinstance(value, str),
        f"{field} must be text so account numbers and wallet addresses are preserved exactly",
    )
    payee = clean_text(value)
    require(
        bool(payee),
        f"{field} is required; use 未显示 or 无法辨认 only when that is true of the evidence",
    )
    require(
        payee.casefold() not in GENERIC_PAYEES and GENERIC_PAYEE_RE.fullmatch(payee) is None,
        f"{field} must be the exact visible name, masked account, or address; "
        f"generic placeholder is forbidden: {payee!r}",
    )
    if state == "visible":
        require(
            payee not in UNKNOWN_PAYEES,
            f"{field} cannot be {payee} when payee_state is visible",
        )
    elif state == "not_shown":
        require(payee == "未显示", f"{field} must be 未显示 when payee_state is not_shown")
    elif state == "unreadable":
        require(payee == "无法辨认", f"{field} must be 无法辨认 when payee_state is unreadable")
    return payee


def validate_thai_bank_account_payee(
    value: object,
    *,
    field: str,
    payee_state: object,
) -> str:
    """Require a Thai-bank recipient to be recorded as the visible account only."""
    payee = validate_payee(
        value,
        field=field,
        payee_state=payee_state,
        require_state=True,
    )
    state = clean_text(payee_state).casefold()
    if state != "visible":
        return payee
    require(
        THAI_BANK_ACCOUNT_RE.fullmatch(payee) is not None
        and THAI_BANK_ACCOUNT_TOKEN_RE.search(payee) is not None,
        f"{field} must contain only the exact visible Thai bank account or masked account; "
        "do not record the Thai recipient name or bank name",
    )
    return payee


def normalize_currency(
    value: object,
    *,
    field: str = "currency",
    allow_none: bool = False,
) -> str | None:
    code = clean_text(value).upper()
    if not code:
        require(allow_none, f"{field} is required")
        return None
    require(code in CURRENCIES, f"{field} has unsupported currency {code!r}")
    return code


def direction_currencies(value: object, *, field: str = "direction") -> tuple[str, str]:
    text = clean_text(value)
    match = re.fullmatch(r"([A-Za-z]+)\s*->\s*([A-Za-z]+)", text)
    require(match is not None, f"{field} must use PAYMENT->PAYOUT syntax")
    assert match is not None
    payment = normalize_currency(match.group(1), field=f"{field}.payment_currency")
    payout = normalize_currency(match.group(2), field=f"{field}.payout_currency")
    assert payment is not None and payout is not None
    require(payment != payout, f"{field} must exchange two different currencies")
    direction = f"{payment}->{payout}"
    require(direction in DIRECTIONS, f"{field} is unsupported: {direction}")
    return DIRECTIONS[direction]


def canonical_direction(value: object, *, field: str = "direction") -> str:
    payment, payout = direction_currencies(value, field=field)
    return f"{payment}->{payout}"


def currency_tolerance(value: object, *, field: str = "currency") -> Decimal:
    code = normalize_currency(value, field=field)
    assert code is not None
    return CURRENCY_TOLERANCES[code]


def flatten_telegram_text(value: object) -> str:
    if isinstance(value, str):
        return clean_text(value)
    if isinstance(value, list):
        chunks: list[str] = []
        for item in value:
            if isinstance(item, str):
                chunks.append(item)
            elif isinstance(item, dict):
                chunks.append(str(item.get("text") or ""))
        return clean_text("".join(chunks))
    return clean_text(value)


def stable_token(*parts: object, length: int = 20) -> str:
    material = "\u241f".join(unicodedata.normalize("NFKC", str(part or "")) for part in parts)
    return hashlib.sha256(material.encode("utf-8")).hexdigest()[:length]


def parse_decimal(value: object, *, field: str, allow_none: bool = False) -> Decimal | None:
    if value in (None, ""):
        if allow_none:
            return None
        raise ValueError(f"{field} is required")
    if isinstance(value, bool) or isinstance(value, float):
        raise ValueError(f"{field} must be an ordinary decimal string, not a floating-point value")
    raw = str(value).strip()
    if re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", raw) is None:
        raise ValueError(f"{field} must be an ordinary decimal without commas or exponents: {value!r}")
    try:
        result = Decimal(raw)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"{field} is not a decimal: {value!r}") from exc
    if not result.is_finite():
        raise ValueError(f"{field} must be finite")
    return result


def parse_cash_amount(value: object, *, field: str, allow_none: bool = False) -> Decimal | None:
    if value in (None, ""):
        if allow_none:
            return None
        raise ValueError(f"{field} is required")
    if isinstance(value, bool) or isinstance(value, float):
        raise ValueError(f"{field} must be handwritten text or an ordinary decimal string")
    raw = unicodedata.normalize("NFKC", str(value)).strip()
    match = CASH_AMOUNT_NOTATION_RE.fullmatch(raw)
    if match is None:
        result = parse_decimal(raw, field=field)
    else:
        base = parse_decimal(match.group("number") or "1", field=f"{field}.coefficient")
        assert base is not None
        if match.group("sign") == "-":
            base = -base
        result = base * CASH_AMOUNT_MULTIPLIERS[match.group("suffix").casefold()]
    assert result is not None
    return result


def decimal_text(value: Decimal | None) -> str | None:
    if value is None:
        return None
    if value == 0:
        return "0"
    rendered = format(value.normalize(), "f")
    return rendered.rstrip("0").rstrip(".") if "." in rendered else rendered


def classify_status(event_type: str, status_text: object, status_class: object = None) -> str:
    text = clean_text(status_text)
    if text and (FAILED_STATUS_RE.search(text) or FAILED_NEGATED_COMPLETION_RE.search(text)):
        return "failed"
    if text and (PENDING_STATUS_RE.search(text) or PENDING_NEGATED_COMPLETION_RE.search(text)):
        return "pending"
    if text and COMPLETED_STATUS_RE.search(text):
        return "completed"
    declared = clean_text(status_class).casefold()
    if declared in STATUS_CLASSES:
        return declared
    if event_type in {"cash_payment", "cash_payout"}:
        return "completed"
    return "unknown"


def classify_fund_status(event: Mapping[str, Any]) -> str:
    ocr = event.get("ocr") if isinstance(event.get("ocr"), dict) else {}
    event_type = str(event.get("type") or "")
    declared = clean_text(ocr.get("status_class")).casefold()
    declared_confidence = clean_text(ocr.get("status_class_confidence")).casefold()
    if declared and declared not in STATUS_CLASSES:
        return "unknown"
    if declared and declared_confidence not in {"", "high"}:
        return "unknown"
    if declared in STATUS_CLASSES:
        return declared
    return classify_status(event_type, ocr.get("status_text"))


def load_normalized(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    require(isinstance(data, dict), "normalized top level must be an object")
    require(data.get("contract_version") == NORMALIZED_CONTRACT, "unsupported normalized contract_version")
    require(data.get("timezone") == "Asia/Bangkok", "normalized timezone must be Asia/Bangkok")
    require(
        isinstance(data.get("source_fingerprint"), str) and data["source_fingerprint"].startswith("sha256:"),
        "invalid normalized source_fingerprint",
    )
    groups = data.get("groups")
    require(isinstance(groups, list), "normalized groups must be a list")
    group_keys: set[str] = set()
    message_ids: set[str] = set()
    media_ids: set[str] = set()
    for group_index, group in enumerate(groups):
        require(isinstance(group, dict), f"groups[{group_index}] must be an object")
        group_key = group.get("group_key")
        require(isinstance(group_key, str) and group_key, f"groups[{group_index}].group_key is required")
        require(group_key not in group_keys, f"duplicate group_key: {group_key}")
        group_keys.add(group_key)
        messages = group.get("messages")
        require(isinstance(messages, list), f"{group_key}: messages must be a list")
        last_sequence = 0
        for message_index, message in enumerate(messages):
            require(isinstance(message, dict), f"{group_key}: messages[{message_index}] must be an object")
            message_id = message.get("message_id")
            timestamp = message.get("timestamp")
            require(isinstance(message_id, str) and message_id, f"{group_key}: missing message_id")
            require(message_id not in message_ids, f"duplicate message_id: {message_id}")
            message_ids.add(message_id)
            require(
                isinstance(timestamp, str) and timestamp.endswith("+07:00"),
                f"{message_id}: timestamp must be Bangkok ISO-8601",
            )
            sequence = message.get("source_sequence")
            require(
                isinstance(sequence, int) and sequence > last_sequence,
                f"{group_key}: source_sequence must be strictly increasing",
            )
            last_sequence = sequence
            require(message.get("role") in {"内部人员", "客户候选", "未知"}, f"{message_id}: invalid role")
            require(isinstance(message.get("media", []), list), f"{message_id}: media must be a list")
            for media_index, media in enumerate(message.get("media", [])):
                require(isinstance(media, dict), f"{message_id}: media[{media_index}] must be an object")
                media_id = media.get("media_id")
                require(isinstance(media_id, str) and media_id, f"{message_id}: media[{media_index}] needs media_id")
                require(media_id not in media_ids, f"duplicate media_id: {media_id}")
                media_ids.add(media_id)
                availability = media.get("availability")
                require(availability in {"available", "missing"}, f"{media_id}: unsupported availability")
                if availability == "available":
                    require(isinstance(media.get("path"), str) and media["path"], f"{media_id}: available media needs path")
                    blob_hash = media.get("blob_sha256")
                    if blob_hash not in (None, ""):
                        require(
                            isinstance(blob_hash, str) and re.fullmatch(r"[0-9a-f]{64}", blob_hash) is not None,
                            f"{media_id}: blob_sha256 must be lowercase SHA-256 when present",
                        )
                    require(
                        isinstance(media.get("byte_size"), int) and media["byte_size"] >= 0,
                        f"{media_id}: invalid byte_size",
                    )
                else:
                    require(not media.get("path"), f"{media_id}: missing media cannot have a path")
                    require(
                        isinstance(media.get("missing_kind"), str) and media["missing_kind"],
                        f"{media_id}: missing media needs missing_kind",
                    )
    return data


def load_events(events_dir: Path, normalized: Mapping[str, Any]) -> tuple[list[dict[str, Any]], str]:
    messages = {
        message["message_id"]: (group["group_key"], message)
        for group in normalized.get("groups", [])
        for message in group.get("messages", [])
    }
    expected_evidence_media = {
        str(media["media_id"])
        for _, message in messages.values()
        for media in message.get("media", [])
        if isinstance(media, dict) and media.get("media_id") and media_is_evidence(media)
    }
    classified_media: set[str] = set()
    events: list[dict[str, Any]] = []
    raw_hashes: list[str] = []
    for path in sorted(events_dir.glob("*.jsonl"), key=lambda item: item.name.casefold()):
        raw = path.read_bytes()
        raw_hashes.append(f"{path.name}:{sha256_bytes(raw)}")
        local_counts: defaultdict[tuple[str, str], int] = defaultdict(int)
        for line_number, line in enumerate(raw.decode("utf-8-sig").splitlines(), start=1):
            if not line.strip() or line.lstrip().startswith("#"):
                continue
            try:
                event = json.loads(line)
            except json.JSONDecodeError as exc:
                raise ValueError(f"{path}:{line_number}: invalid JSON: {exc.msg}") from exc
            require(isinstance(event, dict), f"{path}:{line_number}: event must be an object")
            event_type = event.get("type")
            message_id = event.get("message_id")
            require(event_type in EVENT_TYPES, f"{path}:{line_number}: unsupported event type {event_type!r}")
            require(message_id in messages, f"{path}:{line_number}: unknown message_id {message_id!r}")
            group_key, source_message = messages[message_id]
            require(event.get("group_key") in (None, group_key), f"{path}:{line_number}: group_key mismatch")
            source_media = {
                str(media.get("media_id")): media
                for media in source_message.get("media", [])
                if isinstance(media, dict) and media.get("media_id")
            }
            event_media_id = str(event.get("media_id") or "")
            if event_type in MEDIA_EVENT_TYPES:
                require(event_media_id in source_media, f"{path}:{line_number}: media_id does not belong to message")
                classified_media.add(event_media_id)
            if event_type in FUND_EVENT_TYPES:
                require(
                    media_is_evidence(source_media[event_media_id]),
                    f"{path}:{line_number}: audio/video/sticker/animation cannot be financial evidence",
                )
                ocr = event.get("ocr")
                require(isinstance(ocr, dict), f"{path}:{line_number}: {event_type} requires ocr")
                normalized_ocr = dict(ocr)
                normalized_ocr["currency"] = normalize_currency(
                    ocr.get("currency"),
                    field=f"{path}:{line_number}: ocr.currency",
                    allow_none=True,
                )
                payee = validate_payee(
                    ocr.get("payee"),
                    field=f"{path}:{line_number}: ocr.payee",
                    cash=event_type in {"cash_payment", "cash_payout"},
                    payee_state=ocr.get("payee_state"),
                )
                normalized_ocr["payee"] = payee
                status_class = clean_text(ocr.get("status_class")).casefold()
                status_confidence = clean_text(ocr.get("status_class_confidence")).casefold()
                amount_completeness = clean_text(ocr.get("amount_completeness")).casefold()
                confidence = clean_text(ocr.get("confidence")).casefold()
                if status_class:
                    require(status_class in STATUS_CLASSES, f"{path}:{line_number}: unsupported OCR status_class")
                    require(
                        status_confidence in CONFIDENCE_LEVELS,
                        f"{path}:{line_number}: status_class requires confidence",
                    )
                elif status_confidence:
                    raise ValueError(f"{path}:{line_number}: status_class_confidence requires status_class")
                if amount_completeness:
                    require(
                        amount_completeness in AMOUNT_COMPLETENESS,
                        f"{path}:{line_number}: unsupported amount_completeness",
                    )
                if confidence:
                    require(confidence in CONFIDENCE_LEVELS, f"{path}:{line_number}: unsupported confidence")
                event = {**event, "ocr": normalized_ocr}
            ordinal_key = (str(message_id), str(event_type))
            local_counts[ordinal_key] += 1
            normalized_event = dict(event)
            normalized_event["group_key"] = group_key
            normalized_event["event_id"] = event.get("event_id") or (
                f"{message_id}#{event_type}#{local_counts[ordinal_key]}"
            )
            events.append(normalized_event)
    missing_classifications = sorted(expected_evidence_media - classified_media)
    require(
        not missing_classifications,
        "unclassified evidence media: " + ", ".join(missing_classifications[:20]),
    )
    event_ids = [event["event_id"] for event in events]
    require(len(event_ids) == len(set(event_ids)), "duplicate event_id values")
    fingerprint = "sha256:" + sha256_bytes("\n".join(raw_hashes).encode("utf-8"))
    return events, fingerprint


def message_and_group_indexes(
    normalized: Mapping[str, Any],
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    groups = {group["group_key"]: group for group in normalized.get("groups", [])}
    messages = {
        message["message_id"]: {**message, "_group_key": group["group_key"]}
        for group in groups.values()
        for message in group.get("messages", [])
    }
    return messages, groups
