#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import datetime as dt
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import daily_wecom_report as core


class RegressionMinimalTest(unittest.TestCase):
    @staticmethod
    def _decode_output(data: bytes) -> str:
        for enc in ("utf-8", "gbk", "cp936"):
            try:
                return data.decode(enc)
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="ignore")

    @staticmethod
    def _write_minimal_line_sheet(writer: pd.ExcelWriter, sheet_name: str, date_text: str = "2026-01-01") -> None:
        raw = pd.DataFrame(
            [
                ["", "批次", "投料日期", "烧结压实", "扣电", "扣电", "扣电"],
                ["", "", "", "", "0.1C充", "0.1C放", "0.1C首效"],
                ["", "", "", "", "≤0.055", "≤0.055", "≥96%"],
                [date_text, "DA2601-001", date_text, 2.31, 160.2, 156.3, 97.2],
            ]
        )
        raw.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

    @staticmethod
    def _write_header_only_sheet(writer: pd.ExcelWriter, sheet_name: str) -> None:
        raw = pd.DataFrame(
            [
                ["批次", "", "成品粒度", "", ""],
                ["", "g/cm3", "粒度D10(um)", "粒度D50(um)", "粒度D90(um)"],
                ["", "≥0.7", "", "1.02±0.16μm", "≤7μm"],
            ]
        )
        raw.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

    def test_parse_spec_support_ascii_comparator(self) -> None:
        le = core.parse_spec_from_colname("Li+含量<=0.05")
        self.assertIsNotNone(le)
        assert le is not None
        self.assertIsNone(le.lower)
        self.assertAlmostEqual(le.upper, 0.05, places=8)

        ge = core.parse_spec_from_colname("0.1C充电>=155")
        self.assertIsNotNone(ge)
        assert ge is not None
        self.assertAlmostEqual(ge.lower, 155.0, places=8)
        self.assertIsNone(ge.upper)

    def test_batch_judge_scale_with_spec(self) -> None:
        col = "Li+含量≤0.05"
        df = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 1, 1)],
                "批次": ["DA2601-001"],
                col: [0.04],
            }
        )
        summary = core._batch_out_of_spec_summary(df, [col], scale=10000, enable_spec=True)
        self.assertIsInstance(summary, dict)
        assert isinstance(summary, dict)
        self.assertEqual(summary["异常批次"], 0)

        stat = core._stat_for_cols(df, [col], scale=10000, enable_spec=True)
        self.assertTrue(stat["有数据"])
        self.assertEqual(stat["判异"]["异常"], False)

    def test_header_build_should_not_spread_spec_token(self) -> None:
        header = pd.DataFrame(
            [
                ["扣电", "", ""],
                ["0.1C充", "0.1C放", "0.1C首效"],
                ["≤0.055", "", "≥96%"],
            ]
        )
        cols = core._make_columns_from_multirow_header(header)
        self.assertIn("≤0.055", cols[0])
        self.assertNotIn("≤0.055", cols[1])
        self.assertIn("≥96%", cols[2])

    def test_unreasonable_spec_should_be_ignored_for_01c(self) -> None:
        col = "扣电_0.1C充_≤0.055"
        df = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 1, 1)],
                "批次": ["DA2601-001"],
                col: [160.2],
            }
        )
        stat = core._stat_for_cols(df, [col], scale=1.0, enable_spec=True)
        self.assertTrue(stat["有数据"])
        self.assertIsNone(stat.get("判异"))
        self.assertIsNone(stat.get("批次判异"))

    def test_list_line_sheets_should_skip_header_only_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel = Path(tmp_dir) / "line_filter.xlsx"
            with pd.ExcelWriter(excel) as writer:
                self._write_minimal_line_sheet(writer, "S18-A线")
                self._write_header_only_sheet(writer, "S18-C线")

            usable, skipped = core.list_line_sheets_with_skipped(str(excel))
            self.assertIn("S18-A线", usable)
            self.assertIn("S18-C线", skipped)

    def test_resolve_report_dates_for_global_and_per_line(self) -> None:
        df_a = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 3, 3)],
                "批次": ["DA2603-001"],
                "烧结压实": [2.31],
            }
        )
        df_b = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 2, 2)],
                "批次": ["DB2602-001"],
                "烧结压实": [2.28],
            }
        )
        per_line = core.resolve_report_dates({"S18-A线": df_a, "S18-B线": df_b}, None, "per-line")
        self.assertEqual(per_line["S18-A线"], dt.date(2026, 3, 3))
        self.assertEqual(per_line["S18-B线"], dt.date(2026, 2, 2))

        global_mode = core.resolve_report_dates({"S18-A线": df_a, "S18-B线": df_b}, None, "global")
        self.assertEqual(global_mode["S18-A线"], dt.date(2026, 3, 3))
        self.assertEqual(global_mode["S18-B线"], dt.date(2026, 3, 3))

    def test_validate_sheet_data_quality_issue_and_autofix(self) -> None:
        col = "扣电_0.1C首效_≥96%"
        df = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 2, 5), dt.date(2026, 2, 6)],
                "批次": ["DA2602-010", "DA2602-011"],
                col: [9781, 97.5],
            }
        )

        check_only = core.validate_sheet_data(df, auto_fix=False)
        self.assertEqual(len(check_only["issues"]), 1)
        self.assertEqual(check_only["fixed_count"], 0)
        self.assertEqual(check_only["df"].loc[0, col], 9781)

        fixed = core.validate_sheet_data(df, auto_fix=True)
        self.assertEqual(len(fixed["issues"]), 1)
        self.assertEqual(fixed["fixed_count"], 1)
        self.assertAlmostEqual(float(fixed["df"].loc[0, col]), 97.81, places=2)

    def test_unconfigured_metric_state(self) -> None:
        df = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 3, 3)],
                "批次": ["DA2603-001"],
                "烧结压实": [2.31],
                "粉碎压实": [2.37],
                "扣电_0.1C充_≥155": [160.1],
                "扣电_0.1C放_≥150": [156.0],
                "扣电_0.1C首效_≥96": [97.5],
            }
        )
        metrics = core.extract_metrics(df, dt.date(2026, 3, 3), lookback_days=7, trend_days=7, enable_spec=True)
        self.assertEqual(metrics["成品"]["成品压实"]["状态"], core.STAT_UNCONFIGURED)
        self.assertEqual(metrics["成品"]["比表(麦克比表)"]["状态"], core.STAT_UNCONFIGURED)

    def test_stale_state_should_mark_when_lag_exceeds_threshold(self) -> None:
        df = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 2, 26), dt.date(2026, 3, 4)],
                "批次": ["DA2602-001", "DA2603-001"],
                "烧结压实": [2.31, 2.33],
                "扣电_0.1C首效_≥96%": [97.2, None],
            }
        )
        stale_cfg = core.StaleThresholdConfig(process_days=2, product_days=3, electrochem_days=5)
        metrics = core.extract_metrics(
            df,
            dt.date(2026, 3, 4),
            lookback_days=10,
            trend_days=7,
            enable_spec=True,
            line_label="S18-A线",
            model="S18",
            stale_thresholds=stale_cfg,
        )
        eff = metrics["成品"]["首效"]
        self.assertEqual(eff["状态"], core.STAT_STALE)
        self.assertEqual(eff["lag_days"], 6)

    def test_spec_registry_should_override_col_spec(self) -> None:
        df = pd.DataFrame(
            {
                "投料日期": [dt.date(2026, 3, 3)],
                "批次": ["DA2603-001"],
                "烧结压实": [2.31],
                "粉碎压实": [2.37],
                "成品压实_>=2.37": [2.40],
                "扣电_0.1C首效_≥96": [97.5],
            }
        )
        rules = [core.SpecRule(metric="成品压实", model="S18", spec=core.Spec(lower=2.50, upper=None, text="≥2.50"))]
        metrics = core.extract_metrics(
            df,
            dt.date(2026, 3, 3),
            lookback_days=0,
            trend_days=0,
            enable_spec=True,
            line_label="S18-A线",
            model="S18",
            spec_registry=rules,
        )
        stat = metrics["成品"]["成品压实"]
        self.assertTrue(stat["有数据"])
        self.assertTrue(stat["判异"]["异常"])

    def test_spec_health_should_mark_suspected_and_denoise(self) -> None:
        dates = [dt.date(2026, 2, 1) + dt.timedelta(days=i) for i in range(14)]
        values = [2.40] * 10 + [2.56, 2.57, 2.58, 2.59]
        df = pd.DataFrame(
            {
                "投料日期": dates,
                "批次": [f"DB2602-{i+1:03d}" for i in range(14)],
                "烧结压实": [2.31] * 14,
                "粉碎压实": [2.37] * 14,
                "成品压实_>=2.53": values,
                "扣电_0.1C首效_≥96.3": [97.0] * 14,
            }
        )
        health_cfg = core.SpecHealthConfig(enabled=True, window_days=14, abnormal_ratio_threshold=0.4, min_consecutive_days=5)
        metrics = core.extract_metrics(
            df,
            dt.date(2026, 2, 14),
            lookback_days=0,
            trend_days=0,
            enable_spec=True,
            line_label="S006-B线",
            model="S006",
            spec_health=health_cfg,
        )
        stat = metrics["成品"]["成品压实"]
        self.assertTrue(stat["spec_health"]["suspected"])
        self.assertFalse(core._stat_is_abnormal(stat))

    def test_load_spec_registry_from_default_file(self) -> None:
        rules = core.load_spec_registry(str(REPO_ROOT / "config" / "spec_registry.yaml"))
        self.assertGreaterEqual(len(rules), 1)
        self.assertIsInstance(rules[0], core.SpecRule)

    def test_cli_model_filter_no_match_should_fail(self) -> None:
        cli = REPO_ROOT / "daily_wecom_report.py"
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel = Path(tmp_dir) / "only_s006.xlsx"
            with pd.ExcelWriter(excel) as writer:
                self._write_minimal_line_sheet(writer, "S006-B线")
            proc = subprocess.run(
                [sys.executable, str(cli), "--excel", str(excel), "--model", "S18"],
                cwd=REPO_ROOT,
                capture_output=True,
            )

        self.assertNotEqual(proc.returncode, 0)
        output = f"{self._decode_output(proc.stdout)}\n{self._decode_output(proc.stderr)}"
        self.assertIn("未找到型号为 S18 的候选线别", output)


if __name__ == "__main__":
    unittest.main()
