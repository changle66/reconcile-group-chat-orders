from __future__ import annotations

import json
import tempfile
import unittest
from argparse import Namespace
from pathlib import Path

import core
import reconcile


class ReconcileWorkflowTests(unittest.TestCase):
    def _start_fixture(self, root: Path) -> tuple[Path, dict, dict]:
        export = root / "ChatExport"
        photos = export / "photos"
        photos.mkdir(parents=True)
        (photos / "customer.jpg").write_bytes(b"customer-evidence")
        (photos / "staff.jpg").write_bytes(b"staff-evidence")
        raw = {
            "id": "fixture-small",
            "name": "测试小额群",
            "messages": [
                {
                    "id": 1,
                    "type": "message",
                    "date_unixtime": "1780000000",
                    "from": "Alice",
                    "from_id": "user:alice",
                    "text": "付款100元",
                    "photo": "photos/customer.jpg",
                },
                {
                    "id": 2,
                    "type": "message",
                    "date_unixtime": "1780000060",
                    "from": "QQ财务3",
                    "from_id": "user:staff",
                    "text": "回500铢",
                    "photo": "photos/staff.jpg",
                },
            ],
        }
        (export / "result.json").write_text(
            json.dumps(raw, ensure_ascii=False), encoding="utf-8"
        )
        work = root / "run"
        report = reconcile.start_run(
            Namespace(
                inputs=[export],
                work=work,
                contains="小额",
                timezone="Asia/Bangkok",
                roster=Path(__file__).resolve().parent.parent / "config" / "roster.yaml",
                line_backup=[],
                line_self_name="LINE_SELF",
            )
        )
        run = reconcile._load_run(work)
        normalized = reconcile._load_snapshot(work, run)
        self.assertEqual(report["groups"], 1)
        self.assertTrue(
            all(
                media.get("blob_sha256") in (None, "")
                for message in normalized["groups"][0]["messages"]
                for media in message["media"]
            )
        )
        reconcile.review_command(
            Namespace(work=work, review_action="next", group=None, limit=500)
        )
        return work, run, normalized

    def _write_complete_decision(self, work: Path, *, payee: str = "张三") -> Path:
        run = reconcile._load_run(work)
        run_group = run["groups"][0]
        path = reconcile._decision_path(work, run_group)
        decision = reconcile._load_json(path)
        decision["media_decisions"] = {
            "M0001": {
                "classification": "fund",
                "viewed_original": True,
                "entries": [
                    {
                        "amount": "100",
                        "currency": "CNY",
                        "payee": payee,
                        "side": "payment",
                        "result": "completed",
                        "amount_state": "clear",
                    }
                ],
            },
            "M0002": {
                "classification": "fund",
                "viewed_original": True,
                "entries": [
                    {
                        "amount": "500",
                        "currency": "THB",
                        "payee": "206-4-xxx781",
                        "side": "payout",
                        "result": "completed",
                        "amount_state": "clear",
                    }
                ],
            },
        }
        decision["orders"] = [
            {
                "id": "O001",
                "entry_ids": ["M0001.1", "M0002.1"],
                "source_messages": ["S00001", "S00002"],
                "customer_id": "user:alice",
                "customer_nickname": "Alice",
                "direction": "CNY->THB",
                "rate_state": "adopted",
                "rate": "5",
                "expected_payout_state": "explicit",
                "expected_payout": "500",
            }
        ]
        core.atomic_json(path, decision)
        return path

    def test_start_seal_and_finish_publish_only_the_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            work, _, _ = self._start_fixture(root)
            decision_path = self._write_complete_decision(work)
            self.assertEqual(
                reconcile._load_json(decision_path)["contract_version"],
                "group-chat-decision/2.2",
            )
            sealed = reconcile.review_command(
                Namespace(work=work, review_action="seal", group="测试小额群")
            )
            self.assertTrue(sealed["sealed"])
            output = root / "ledger.xlsx"
            result = reconcile.finish_run(
                Namespace(
                    work=work,
                    output=output,
                    template=Path(__file__).resolve().parent.parent / "assets" / "模版.xlsx",
                )
            )
            self.assertTrue(output.is_file())
            self.assertEqual(result["pending_orders"], 0)
            self.assertFalse((work / "events").exists())
            self.assertFalse((work / "orders.json").exists())
            self.assertFalse((work / "simple_plan.json").exists())

    def test_generic_payee_is_rejected_at_group_seal(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            self._write_complete_decision(work, payee="截图所示人民币收款方")
            with self.assertRaisesRegex(ValueError, "generic placeholder"):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_one_image_can_compile_multiple_fund_entries(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, normalized = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["media_decisions"]["M0001"]["entries"].append(
                {
                    "amount": "25",
                    "currency": "USD",
                    "payee": "张三",
                    "side": "payment",
                    "result": "completed",
                    "amount_state": "clear",
                }
            )
            decision["orders"][0]["entry_ids"].insert(1, "M0001.2")
            core.atomic_json(decision_path, decision)
            reconcile.review_command(
                Namespace(work=work, review_action="seal", group="测试小额群")
            )
            sealed = reconcile._load_json(decision_path)
            events, _ = reconcile._compile_decisions_v2(
                normalized,
                {str(normalized["groups"][0]["group_key"]): sealed},
            )
            fund_events = [event for event in events if event["type"] in core.FUND_EVENT_TYPES]
            self.assertEqual(len(fund_events), 3)
            first_media_ids = [
                event["event_id"]
                for event in fund_events
                if event["media_id"].endswith("#media:0")
                and event["message_id"].endswith(":1")
            ]
            self.assertEqual(len(first_media_ids), 2)

    def test_seal_rejects_missing_explicit_fund_side(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["media_decisions"]["M0001"]["entries"][0].pop("side")
            core.atomic_json(decision_path, decision)

            with self.assertRaisesRegex(ValueError, "side is required"):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_seal_rejects_fund_entries_not_explicitly_assigned_to_an_order(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["media_decisions"]["M0001"]["entries"].append(
                {
                    "amount": "25",
                    "currency": "USD",
                    "payee": "张三",
                    "side": "payment",
                    "result": "completed",
                    "amount_state": "clear",
                }
            )
            core.atomic_json(decision_path, decision)

            with self.assertRaisesRegex(ValueError, "must be explicitly assigned to an order"):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_rate_requires_operator_only_when_it_calculates_expected_payout(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["orders"][0].pop("expected_payout")
            decision["orders"][0]["expected_payout_state"] = "calculated_from_rate"
            core.atomic_json(decision_path, decision)
            with self.assertRaisesRegex(ValueError, "rate_operator is required"):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_calculated_expected_state_uses_the_existing_ledger_math(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, normalized = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            order = decision["orders"][0]
            order.pop("expected_payout")
            order["expected_payout_state"] = "calculated_from_rate"
            order["rate_operator"] = "multiply"
            core.atomic_json(decision_path, decision)

            reconcile.review_command(
                Namespace(work=work, review_action="seal", group="测试小额群")
            )
            sealed = reconcile._load_json(decision_path)
            events, plan = reconcile._compile_decisions_v2(
                normalized,
                {str(normalized["groups"][0]["group_key"]): sealed},
            )
            orders, _ = reconcile.simple_ledger.compile_simple_ledger(
                normalized,
                events,
                plan["events_fingerprint"],
                plan,
            )
            self.assertEqual(
                orders["groups"][0]["orders"][0]["expected_payout"],
                "500",
            )

    def test_check_reports_and_seal_rejects_missing_pricing_states(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["orders"][0].pop("rate_state")
            decision["orders"][0].pop("expected_payout_state")
            core.atomic_json(decision_path, decision)

            checked = reconcile.review_command(
                Namespace(work=work, review_action="check", group="测试小额群")
            )
            self.assertEqual(checked["pricing_scopes"], 1)
            self.assertEqual(checked["missing_pricing_states"], 2)
            self.assertFalse(checked["sealed"])

            with self.assertRaisesRegex(
                ValueError,
                r"pricing states are required.*rate_state.*expected_payout_state",
            ):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_explicit_absence_seals_without_guessing_values(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, normalized = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            order = decision["orders"][0]
            order["rate_state"] = "not_stated"
            order.pop("rate")
            order["expected_payout_state"] = "not_stated"
            order.pop("expected_payout")
            core.atomic_json(decision_path, decision)

            reconcile.review_command(
                Namespace(work=work, review_action="seal", group="测试小额群")
            )
            sealed = reconcile._load_json(decision_path)
            _, plan = reconcile._compile_decisions_v2(
                normalized,
                {str(normalized["groups"][0]["group_key"]): sealed},
            )
            compiled = plan["groups"][0]["orders"][0]
            self.assertNotIn("rate_state", compiled)
            self.assertNotIn("expected_payout_state", compiled)
            self.assertNotIn("rate", compiled)
            self.assertNotIn("expected_payout", compiled)

    def test_pricing_state_rejects_a_conflicting_value(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["orders"][0]["rate_state"] = "not_stated"
            core.atomic_json(decision_path, decision)

            with self.assertRaisesRegex(ValueError, "rate must be empty"):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_multileg_order_requires_pricing_states_on_every_leg(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            order = decision["orders"][0]
            order["direction"] = ""
            for field in (
                "rate_state",
                "rate",
                "expected_payout_state",
                "expected_payout",
            ):
                order.pop(field, None)
            order["legs"] = [
                {
                    "leg_id": "thb",
                    "direction": "CNY->THB",
                    "allocation_amount": "60",
                    "rate_state": "adopted",
                    "rate": "5",
                    "expected_payout_state": "explicit",
                    "expected_payout": "300",
                    "payout_entry_ids": ["M0002.1"],
                    "recovery_entry_ids": [],
                },
                {
                    "leg_id": "usd",
                    "direction": "CNY->USD",
                    "allocation_amount": "40",
                    "payout_entry_ids": [],
                    "recovery_entry_ids": [],
                },
            ]
            core.atomic_json(decision_path, decision)

            with self.assertRaisesRegex(
                ValueError,
                r"pricing states are required.*legs\[1\].*rate_state.*expected_payout_state",
            ):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

            decision = reconcile._load_json(decision_path)
            second_leg = decision["orders"][0]["legs"][1]
            second_leg["rate_state"] = "not_stated"
            second_leg["expected_payout_state"] = "not_stated"
            core.atomic_json(decision_path, decision)
            reconcile.review_command(
                Namespace(work=work, review_action="seal", group="测试小额群")
            )
            sealed = reconcile._load_json(decision_path)
            normalized = reconcile._load_snapshot(work, reconcile._load_run(work))
            _, plan = reconcile._compile_decisions_v2(
                normalized,
                {str(normalized["groups"][0]["group_key"]): sealed},
            )
            for leg in plan["groups"][0]["orders"][0]["legs"]:
                self.assertNotIn("rate_state", leg)
                self.assertNotIn("expected_payout_state", leg)

    def test_previous_decision_contract_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, _ = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["contract_version"] = "group-chat-decision/2.1"
            core.atomic_json(decision_path, decision)

            with self.assertRaisesRegex(ValueError, "unsupported decision contract"):
                reconcile.review_command(
                    Namespace(work=work, review_action="seal", group="测试小额群")
                )

    def test_order_note_survives_decision_compile(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, _, normalized = self._start_fixture(Path(temporary))
            decision_path = self._write_complete_decision(work)
            decision = reconcile._load_json(decision_path)
            decision["orders"][0]["note"] = "特殊安排写入订单汇总备注。"
            core.atomic_json(decision_path, decision)
            reconcile.review_command(
                Namespace(work=work, review_action="seal", group="测试小额群")
            )
            sealed = reconcile._load_json(decision_path)
            _, plan = reconcile._compile_decisions_v2(
                normalized,
                {str(normalized["groups"][0]["group_key"]): sealed},
            )
            self.assertEqual(
                plan["groups"][0]["orders"][0]["note"],
                "特殊安排写入订单汇总备注。",
            )

    def test_changed_snapshot_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            work, run, _ = self._start_fixture(Path(temporary))
            snapshot = reconcile._snapshot_path(work)
            snapshot.write_text(snapshot.read_text(encoding="utf-8") + " ", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "snapshot changed"):
                reconcile._load_snapshot(work, run)


if __name__ == "__main__":
    unittest.main()
