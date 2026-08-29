#!/usr/bin/env python3
"""Interface-level regression tests for the group decision workflow."""

from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

import core
import group_decisions
import simple_ledger


GROUP_KEY = "telegram:4909790695"


def media(media_id: str) -> dict[str, object]:
    return {
        "media_id": media_id,
        "kind": "image",
        "availability": "available",
        "path": f"C:/evidence/{media_id.replace(':', '_')}.jpg",
        "blob_sha256": hashlib.sha256(media_id.encode("utf-8")).hexdigest(),
        "byte_size": 100,
    }


def message(
    sequence: int,
    *,
    text: str = "",
    sender_id: str = "customer-1",
    sender_name: str = "客户甲",
    role: str = "客户候选",
    media_id: str | None = None,
    reply_to_message_id: str | None = None,
    group_key: str = GROUP_KEY,
) -> dict[str, object]:
    message_id = f"{group_key}:{sequence}"
    return {
        "message_id": message_id,
        "timestamp": f"2026-08-02T12:{sequence:02d}:00+07:00",
        "source_sequence": sequence,
        "sender_id": sender_id,
        "sender_name": sender_name,
        "role": role,
        "text": text,
        "reply_to_message_id": reply_to_message_id,
        "mentions": [],
        "media": [media(media_id)] if media_id else [],
        "excluded_from_accounting": False,
    }


def normalized(messages: list[dict[str, object]], *, extra_groups: list[dict] | None = None) -> dict:
    groups = [
        {
            "group_key": GROUP_KEY,
            "platform": "Telegram",
            "group_name": "小额测试群",
            "messages": messages,
        }
    ]
    if extra_groups:
        groups.extend(extra_groups)
    return {
        "contract_version": core.NORMALIZED_CONTRACT,
        "timezone": "Asia/Bangkok",
        "source_fingerprint": "sha256:" + "1" * 64,
        "groups": groups,
    }


def ocr(amount: str, currency: str, payee: str) -> dict[str, object]:
    return {
        "amount": amount,
        "currency": currency,
        "payee": payee,
        "status_text": "成功",
        "status_class": "completed",
        "status_class_confidence": "high",
        "amount_completeness": "complete",
        "confidence": "high",
    }


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


class GroupDecisionTests(unittest.TestCase):
    def test_read_group_paginates_every_message_once_and_resolves_cross_page_reply(self) -> None:
        messages = [message(sequence) for sequence in range(1, 13)]
        messages[9]["reply_to_message_id"] = messages[0]["message_id"]
        source = normalized(messages)

        cursor: str | None = None
        collected: list[str] = []
        reply_snapshot: dict | None = None
        while True:
            page = group_decisions.read_group(source, GROUP_KEY, cursor=cursor, limit=4)
            collected.extend(item["message_id"] for item in page["messages"])
            for item in page["messages"]:
                if item["message_id"] == messages[9]["message_id"]:
                    reply_snapshot = item["reply_message"]
            if page["done"]:
                self.assertIsNone(page["next_cursor"])
                break
            cursor = page["next_cursor"]

        self.assertEqual(collected, [item["message_id"] for item in messages])
        self.assertEqual(len(collected), len(set(collected)))
        self.assertIsNotNone(reply_snapshot)
        self.assertEqual(reply_snapshot["message_id"], messages[0]["message_id"])

    def test_cursor_is_bound_to_the_current_group_content(self) -> None:
        source = normalized([message(sequence) for sequence in range(1, 7)])
        first = group_decisions.read_group(source, GROUP_KEY, limit=2)
        source["groups"][0]["messages"][0]["text"] = "source changed"

        with self.assertRaisesRegex(ValueError, "cursor is stale"):
            group_decisions.read_group(
                source,
                GROUP_KEY,
                cursor=first["next_cursor"],
                limit=2,
            )

    def test_prepare_creates_one_group_decision_without_overwriting(self) -> None:
        payment_media = f"{GROUP_KEY}:2#media:0"
        source = normalized([message(1), message(2, media_id=payment_media)])
        with tempfile.TemporaryDirectory(prefix="group-decisions-prepare-") as temporary:
            decisions_dir = Path(temporary)
            first = group_decisions.prepare_decision_files(source, decisions_dir)
            path = Path(first["groups"][0]["decision_file"])
            document = read_json(path)
            self.assertEqual(document["evidence_media_count"], 1)
            self.assertEqual(document["media_decisions"][0]["event_id"], f"{payment_media}#event")
            document["orders"] = [{"keep": True}]
            core.atomic_json(path, document)

            second = group_decisions.prepare_decision_files(source, decisions_dir)
            self.assertFalse(second["groups"][0]["written"])
            self.assertEqual(read_json(path)["orders"], [{"keep": True}])

    def test_fund_evidence_writer_requires_and_preserves_explicit_payee(self) -> None:
        base_ocr = {
            "amount": "10416",
            "currency": "THB",
            "status_text": "Transaction successful",
            "status_class": "completed",
            "status_class_confidence": "high",
            "amount_completeness": "complete",
            "confidence": "high",
        }
        for payee in ("王少秋（**秋）", "206-4-xxx781"):
            with self.subTest(payee=payee):
                decision = group_decisions.fund_evidence_decision(
                    event_type="payout_screenshot",
                    payee=payee,
                    ocr=base_ocr,
                    flow_side="payout",
                )
                self.assertEqual(decision["ocr"]["payee"], payee)

        with self.assertRaisesRegex(ValueError, "payee is required"):
            group_decisions.fund_evidence_decision(
                event_type="payout_screenshot",
                payee=" ",
                ocr=base_ocr,
            )

    def test_missing_media_remains_covered_without_a_model_decision(self) -> None:
        missing_media_id = f"{GROUP_KEY}:2#media:0"
        missing_message = message(2)
        missing_message["media"] = [
            {
                "media_id": missing_media_id,
                "kind": "image",
                "availability": "missing",
                "path": None,
                "blob_sha256": None,
                "byte_size": None,
                "missing_kind": "not_exported",
            }
        ]
        source = normalized([message(1), missing_message])
        with tempfile.TemporaryDirectory(prefix="group-decisions-missing-") as temporary:
            root = Path(temporary)
            decisions_dir = root / "decisions"
            group_decisions.prepare_decision_files(source, decisions_dir)
            report = group_decisions.compile_group_decisions(
                source,
                decisions_dir,
                root / "events",
                root / "plan.json",
            )
            events, _ = core.load_events(root / "events", source)

            self.assertEqual(report["dispositions"]["missing"], 1)
            self.assertEqual(events[0]["type"], "missing_evidence_media")

    def test_reference_media_cannot_carry_fund_fields(self) -> None:
        payment_media = f"{GROUP_KEY}:2#media:0"
        source = normalized([message(1), message(2, media_id=payment_media)])
        with tempfile.TemporaryDirectory(prefix="group-decisions-reference-") as temporary:
            root = Path(temporary)
            decisions_dir = root / "decisions"
            group_decisions.prepare_decision_files(source, decisions_dir)
            path = next(decisions_dir.glob("*.json"))
            document = read_json(path)
            document["media_decisions"][0]["decision"].update(
                {
                    "disposition": "reference",
                    "event_type": "payment_screenshot",
                    "ocr": ocr("100", "CNY", "测试收款方"),
                }
            )
            core.atomic_json(path, document)

            with self.assertRaisesRegex(ValueError, "reference must not contain"):
                group_decisions.compile_group_decisions(
                    source,
                    decisions_dir,
                    root / "events",
                    root / "plan.json",
                )

    def test_distant_quote_and_formula_are_cited_in_the_same_semantic_pass(self) -> None:
        payment_media = f"{GROUP_KEY}:21#media:0"
        payout_media = f"{GROUP_KEY}:33#media:0"
        messages = [message(sequence, text=f"unrelated {sequence}") for sequence in range(1, 34)]
        messages[0]["text"] = "quoted channel rate 4.96"
        messages[7]["text"] = "30000 / 4.96 = 6048"
        messages[20] = message(21, text="customer payment", media_id=payment_media)
        messages[32] = message(
            33,
            text="internal payout",
            sender_id="staff-1",
            sender_name="财务",
            role="内部人员",
            media_id=payout_media,
        )
        source = normalized(messages)

        with tempfile.TemporaryDirectory(prefix="group-decisions-distant-") as temporary:
            root = Path(temporary)
            decisions_dir = root / "decisions"
            group_decisions.prepare_decision_files(source, decisions_dir)
            decision_path = next(decisions_dir.glob("*.json"))
            document = read_json(decision_path)
            by_media = {item["media_id"]: item for item in document["media_decisions"]}
            by_media[payment_media]["decision"] = group_decisions.fund_evidence_decision(
                event_type="payment_screenshot",
                payee="Goddess Space",
                ocr=ocr("6048", "CNY", "Goddess Space"),
                flow_side="payment",
            )
            by_media[payout_media]["decision"] = group_decisions.fund_evidence_decision(
                event_type="payout_screenshot",
                payee="206-4-xxx781",
                ocr=ocr("30000", "THB", "206-4-xxx781"),
                flow_side="payout",
            )
            document["orders"] = [
                {
                    "event_ids": [
                        by_media[payment_media]["event_id"],
                        by_media[payout_media]["event_id"],
                    ],
                    "source_message_ids": [
                        messages[0]["message_id"],
                        messages[7]["message_id"],
                        messages[20]["message_id"],
                        messages[32]["message_id"],
                    ],
                    "customer_id": "customer-1",
                    "customer_nickname": "客户",
                    "direction": "CNY->THB",
                    "rate": "4.96",
                    "expected_payout": "30000",
                }
            ]
            core.atomic_json(decision_path, document)

            report = group_decisions.compile_group_decisions(
                source,
                decisions_dir,
                root / "events",
                root / "plan.json",
            )
            events, events_fingerprint = core.load_events(root / "events", source)
            plan = read_json(root / "plan.json")
            orders, _ = simple_ledger.compile_simple_ledger(
                source,
                events,
                events_fingerprint,
                plan,
            )

            compiled_plan_order = plan["groups"][0]["orders"][0]
            order = orders["groups"][0]["orders"][0]
            self.assertEqual(report["orders"], 1)
            self.assertEqual(compiled_plan_order["source_message_ids"][0], messages[0]["message_id"])
            self.assertEqual(compiled_plan_order["rate"], "4.96")
            self.assertEqual(compiled_plan_order["expected_payout"], "30000")
            self.assertEqual(
                compiled_plan_order["event_sides"],
                {
                    by_media[payment_media]["event_id"]: "payment",
                    by_media[payout_media]["event_id"]: "payout",
                },
            )
            self.assertEqual(order["payment_total"], "6048")
            self.assertEqual(order["actual_payout_total"], "30000")
            self.assertEqual(order["actual_rate"], "4.96")
            self.assertEqual(order["expected_payout"], "30000")
            self.assertEqual(order["review_result"], "")
            self.assertEqual(
                [flow["payee"] for flow in order["flows"]],
                ["Goddess Space", "206-4-xxx781"],
            )

    def test_order_source_citation_must_exist_in_the_same_group(self) -> None:
        payment_media = f"{GROUP_KEY}:2#media:0"
        other_group_key = "telegram:other"
        other_message = message(1, group_key=other_group_key)
        source = normalized(
            [message(1), message(2, media_id=payment_media)],
            extra_groups=[
                {
                    "group_key": other_group_key,
                    "platform": "Telegram",
                    "group_name": "另一个群",
                    "messages": [other_message],
                }
            ],
        )
        with tempfile.TemporaryDirectory(prefix="group-decisions-cross-group-") as temporary:
            root = Path(temporary)
            decisions_dir = root / "decisions"
            report = group_decisions.prepare_decision_files(source, decisions_dir)
            path = Path(report["groups"][0]["decision_file"])
            document = read_json(path)
            item = document["media_decisions"][0]
            item["decision"].update(
                {
                    "disposition": "order_evidence",
                    "event_type": "payment_screenshot",
                    "ocr": ocr("100", "CNY", "测试收款方"),
                }
            )
            document["orders"] = [
                {
                    "event_ids": [item["event_id"]],
                    "source_message_ids": [other_message["message_id"]],
                }
            ]
            core.atomic_json(path, document)

            with self.assertRaisesRegex(ValueError, "missing or cross-group message"):
                group_decisions.compile_group_decisions(
                    source,
                    decisions_dir,
                    root / "events",
                    root / "plan.json",
                )

    def test_order_evidence_must_be_assigned_exactly_once(self) -> None:
        payment_media = f"{GROUP_KEY}:2#media:0"
        source = normalized([message(1), message(2, media_id=payment_media)])
        with tempfile.TemporaryDirectory(prefix="group-decisions-unassigned-") as temporary:
            root = Path(temporary)
            decisions_dir = root / "decisions"
            group_decisions.prepare_decision_files(source, decisions_dir)
            path = next(decisions_dir.glob("*.json"))
            document = read_json(path)
            document["media_decisions"][0]["decision"].update(
                {
                    "disposition": "order_evidence",
                    "event_type": "payment_screenshot",
                    "ocr": ocr("100", "CNY", "测试收款方"),
                }
            )
            core.atomic_json(path, document)

            with self.assertRaisesRegex(ValueError, "not assigned to an order"):
                group_decisions.compile_group_decisions(
                    source,
                    decisions_dir,
                    root / "events",
                    root / "plan.json",
                )

    def test_stale_group_fingerprint_is_rejected_before_outputs_are_written(self) -> None:
        source = normalized([message(1)])
        with tempfile.TemporaryDirectory(prefix="group-decisions-stale-") as temporary:
            root = Path(temporary)
            decisions_dir = root / "decisions"
            group_decisions.prepare_decision_files(source, decisions_dir)
            source["groups"][0]["messages"][0]["text"] = "changed source"

            with self.assertRaisesRegex(ValueError, "group_fingerprint is stale"):
                group_decisions.compile_group_decisions(
                    source,
                    decisions_dir,
                    root / "events",
                    root / "plan.json",
                )
            self.assertFalse((root / "events").exists())
            self.assertFalse((root / "plan.json").exists())


if __name__ == "__main__":
    unittest.main()
