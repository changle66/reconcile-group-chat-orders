#!/usr/bin/env python3
"""Run tests partitioned by their relationship to the public skill workflow."""

from __future__ import annotations

import argparse
import importlib
import sys
import unittest
from collections.abc import Iterable


TEST_MODULES = (
    "test_reconcile",
    "test_line_android_miui",
    "test_finance_materials",
    "test_simple_ledger",
)

# These non-ledger tests already exercise the supported start/review/finish
# protocol, the supported 3.1 checkpoint upgrade, or a current conditional
# input adapter. Every simple-ledger test is current unless explicitly listed
# as compatibility-only below.
CURRENT_TEST_IDS = frozenset(
    {
        "test_reconcile.ReconcileWorkflowTests.test_apply_batch_atomically_updates_a_controlled_decision",
        "test_reconcile.ReconcileWorkflowTests.test_controlled_decision_rejects_direct_semantic_overwrite",
        "test_reconcile.ReconcileWorkflowTests.test_invalid_apply_batch_does_not_partially_write_the_decision",
        "test_reconcile.ReconcileWorkflowTests.test_large_group_finish_writes_daily_records_and_rate_grouped_totals",
        "test_reconcile.ReconcileWorkflowTests.test_large_group_multiple_fund_images_compile_to_one_full_order",
        "test_reconcile.ReconcileWorkflowTests.test_large_group_review_uses_full_order_contract",
        "test_reconcile.ReconcileWorkflowTests.test_order_continuity_risk_flags_split_without_mutating_or_blocking",
        "test_reconcile.ReconcileWorkflowTests.test_order_continuity_risk_ignores_correctly_merged_order",
        "test_reconcile.ReconcileWorkflowTests.test_order_continuity_risk_ignores_two_complete_independent_orders",
        "test_reconcile.ReconcileWorkflowTests.test_order_continuity_risk_skips_balance_linked_orders",
        "test_reconcile.ReconcileWorkflowTests.test_page_commit_advances_and_carries_open_order_context",
        "test_reconcile.ReconcileWorkflowTests.test_page_commit_requires_explicit_open_order_state_and_valid_token",
        "test_reconcile.ReconcileWorkflowTests.test_review_next_default_returns_largest_complete_page_within_budget",
        "test_reconcile.ReconcileWorkflowTests.test_review_next_is_read_only_until_page_commit",
        "test_reconcile.ReconcileWorkflowTests.test_start_date_filters_messages_by_bangkok_calendar_day",
        "test_reconcile.ReconcileWorkflowTests.test_start_large_mode_requires_one_accounting_date",
        "test_reconcile.ReconcileWorkflowTests.test_start_large_mode_selects_every_group_except_small_and_finance",
        "test_reconcile.ReconcileWorkflowTests.test_start_rejects_invalid_date_before_creating_work_directory",
        "test_reconcile.ReconcileWorkflowTests.test_start_rejects_invalid_time_window_before_creating_work_directory",
        "test_reconcile.ReconcileWorkflowTests.test_start_seal_and_finish_publish_only_the_workbook",
        "test_reconcile.ReconcileWorkflowTests.test_start_time_range_filters_start_inclusive_end_exclusive",
        "test_reconcile.ReconcileWorkflowTests.test_upgrade_checkpoints_preserves_legacy_decisions_but_resets_read_progress",
        "test_line_android_miui.LineAndroidMiuiTests.test_compressed_backup_is_rejected_clearly",
        "test_line_android_miui.LineAndroidMiuiTests.test_start_discovers_and_normalizes_miui_backup",
        "test_line_android_miui.LineAndroidMiuiTests.test_start_finance_mode_selects_finance_miui_group",
        "test_line_android_miui.LineAndroidMiuiTests.test_start_large_mode_selects_nonexcluded_miui_groups",
    }
)

CURRENT_TEST_PREFIXES = (
    "test_finance_materials.FinanceMaterialsTests.",
    "test_simple_ledger.SimpleLedgerTests.",
)
COMPATIBILITY_TEST_IDS = frozenset(
    {
        "test_reconcile.ReconcileWorkflowTests.test_legacy_large_exchange_contract_still_finishes",
        "test_simple_ledger.SimpleLedgerTests.test_legacy_simple_plan_contract_remains_supported",
    }
)

EXPECTED_ALL_TESTS = 115


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "suite",
        choices=("current", "compatibility", "all", "inventory"),
        help="test partition to run, or inventory to list partition membership",
    )
    return parser.parse_args(argv)


def _flatten(suite: unittest.TestSuite) -> Iterable[unittest.TestCase]:
    for item in suite:
        if isinstance(item, unittest.TestSuite):
            yield from _flatten(item)
        else:
            yield item


def _all_tests() -> dict[str, unittest.TestCase]:
    loader = unittest.defaultTestLoader
    cases: dict[str, unittest.TestCase] = {}
    for module_name in TEST_MODULES:
        module = importlib.import_module(module_name)
        for case in _flatten(loader.loadTestsFromModule(module)):
            test_id = case.id()
            if test_id in cases:
                raise RuntimeError(f"duplicate test id: {test_id}")
            cases[test_id] = case
    if len(cases) != EXPECTED_ALL_TESTS:
        raise RuntimeError(
            f"test inventory changed: expected {EXPECTED_ALL_TESTS}, found {len(cases)}; "
            "classify the change before updating EXPECTED_ALL_TESTS"
        )
    missing = sorted(CURRENT_TEST_IDS - set(cases))
    if missing:
        raise RuntimeError("current test ids no longer exist: " + ", ".join(missing))
    missing_compatibility = sorted(COMPATIBILITY_TEST_IDS - set(cases))
    if missing_compatibility:
        raise RuntimeError(
            "compatibility test ids no longer exist: "
            + ", ".join(missing_compatibility)
        )
    return cases


def _current_test_ids(cases: dict[str, unittest.TestCase]) -> frozenset[str]:
    selected = set(CURRENT_TEST_IDS)
    selected.update(
        test_id
        for test_id in cases
        if test_id not in COMPATIBILITY_TEST_IDS
        and test_id.startswith(CURRENT_TEST_PREFIXES)
    )
    return frozenset(selected)


def _partition(
    cases: dict[str, unittest.TestCase],
    current_test_ids: frozenset[str],
    suite_name: str,
) -> list[unittest.TestCase]:
    if suite_name == "current":
        selected = current_test_ids
    elif suite_name == "compatibility":
        selected = set(cases) - current_test_ids
    else:
        selected = set(cases)
    return [cases[test_id] for test_id in sorted(selected)]


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    try:
        cases = _all_tests()
    except (ImportError, RuntimeError) as exc:
        print(f"Test partition error: {exc}", file=sys.stderr)
        return 2
    current_test_ids = _current_test_ids(cases)
    compatibility_count = len(cases) - len(current_test_ids)
    if args.suite == "inventory":
        print(
            f"current={len(current_test_ids)} compatibility={compatibility_count} "
            f"all={len(cases)}"
        )
        for test_id in sorted(cases):
            partition = "current" if test_id in current_test_ids else "compatibility"
            print(f"{partition}\t{test_id}")
        return 0
    selected = _partition(cases, current_test_ids, args.suite)
    print(
        f"Running {args.suite} skill tests: {len(selected)} "
        f"(current={len(current_test_ids)}, compatibility={compatibility_count})",
        flush=True,
    )
    result = unittest.TextTestRunner(verbosity=2).run(unittest.TestSuite(selected))
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
