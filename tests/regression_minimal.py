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
    def _write_minimal_excel(path: Path, sheet_name: str) -> None:
        raw = pd.DataFrame(
            [
                ["", "\u6279\u6b21", "\u6295\u6599\u65e5\u671f", "\u70e7\u7ed3\u538b\u5b9e"],
                ["2026-01-01", "DA2601-001", "2026-01-01", 1.23],
            ]
        )
        with pd.ExcelWriter(path) as writer:
            raw.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

    def test_parse_spec_support_ascii_comparator(self) -> None:
        le = core.parse_spec_from_colname("Li+\u542b\u91cf<=0.05")
        self.assertIsNotNone(le)
        assert le is not None
        self.assertIsNone(le.lower)
        self.assertAlmostEqual(le.upper, 0.05, places=8)

        ge = core.parse_spec_from_colname("0.1C\u5145\u7535>=155")
        self.assertIsNotNone(ge)
        assert ge is not None
        self.assertAlmostEqual(ge.lower, 155.0, places=8)
        self.assertIsNone(ge.upper)

    def test_batch_judge_scale_with_spec(self) -> None:
        col = "Li+\u542b\u91cf\u22640.05"
        df = pd.DataFrame(
            {
                "\u6295\u6599\u65e5\u671f": [dt.date(2026, 1, 1)],
                "\u6279\u6b21": ["DA2601-001"],
                col: [0.04],
            }
        )

        summary = core._batch_out_of_spec_summary(df, [col], scale=10000, enable_spec=True)
        self.assertIsInstance(summary, dict)
        assert isinstance(summary, dict)
        self.assertEqual(summary["\u5f02\u5e38\u6279\u6b21"], 0)

        stat = core._stat_for_cols(df, [col], scale=10000, enable_spec=True)
        self.assertTrue(stat["\u6709\u6570\u636e"])
        self.assertEqual(stat["\u5224\u5f02"]["\u5f02\u5e38"], False)

    def test_cli_model_filter_no_match_should_fail(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        cli = repo_root / "daily_wecom_report.py"

        with tempfile.TemporaryDirectory() as tmp_dir:
            excel = Path(tmp_dir) / "only_s006.xlsx"
            self._write_minimal_excel(excel, "S006-B\u7ebf")
            proc = subprocess.run(
                [sys.executable, str(cli), "--excel", str(excel), "--model", "S18"],
                cwd=repo_root,
                capture_output=True,
            )

        self.assertNotEqual(proc.returncode, 0)
        output = f"{self._decode_output(proc.stdout)}\n{self._decode_output(proc.stderr)}"
        self.assertIn("\u672a\u627e\u5230\u578b\u53f7\u4e3a S18 \u7684\u5019\u9009\u7ebf\u522b", output)


if __name__ == "__main__":
    unittest.main()
