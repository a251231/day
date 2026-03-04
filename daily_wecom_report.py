#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import datetime as dt
import glob
import re
import sys
from dataclasses import dataclass
from typing import Any, Iterable, Optional

import numpy as np
import pandas as pd

LINE_A_SHEET = "S18-A线"
LINE_B_SHEET = "S006-B线"
LINE_A_LABEL = "S18-A线"
LINE_B_LABEL = "S006-B线"


@dataclass(frozen=True)
class ProductSpecProfile:
    model: str
    enable_spec: bool


PRODUCT_SPEC_PROFILES: dict[str, ProductSpecProfile] = {
    "S18": ProductSpecProfile(model="S18", enable_spec=True),
    "S006": ProductSpecProfile(model="S006", enable_spec=True),
}
DEFAULT_PROFILE = ProductSpecProfile(model="UNKNOWN", enable_spec=False)
LINE_A_SPEC_ENABLED = True
LINE_B_SPEC_ENABLED = True


def _ensure_utf8_stdout() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _norm_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    text = str(value).strip()
    if text == "nan":
        return ""
    return text


def _ffill_right(values: Iterable[Any]) -> list[str]:
    out: list[str] = []
    last = ""
    for v in values:
        s = _norm_cell(v)
        if s:
            last = s
        out.append(last)
    return out


def _dedupe(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for name in names:
        count = seen.get(name, 0) + 1
        seen[name] = count
        out.append(name if count == 1 else f"{name}__{count}")
    return out


def _sanitize_col(name: str) -> str:
    text = re.sub(r"\s+", "", str(name)).strip()
    return (
        text.replace("＋", "+")
        .replace("﹢", "+")
        .replace("。", ".")
        .replace("．", ".")
        .replace("｡", ".")
        .replace("，", ",")
        .replace("（", "(")
        .replace("）", ")")
        .replace("～", "~")
        .replace("—", "-")
        .replace("−", "-")
        .replace("－", "-")
    )


def detect_model_from_sheet(sheet_name: str) -> str:
    text = _sanitize_col(sheet_name).upper()
    if "S18" in text:
        return "S18"
    if "S006" in text:
        return "S006"
    return "UNKNOWN"


def get_profile_for_sheet(sheet_name: str) -> ProductSpecProfile:
    return PRODUCT_SPEC_PROFILES.get(detect_model_from_sheet(sheet_name), DEFAULT_PROFILE)


def list_workbook_sheets(path: str) -> list[str]:
    with pd.ExcelFile(path, engine="calamine") as book:
        return [str(n).strip() for n in book.sheet_names if str(n).strip()]


def list_line_sheets(path: str) -> list[str]:
    sheets = list_workbook_sheets(path)
    line_like = [s for s in sheets if "线" in s]
    if not line_like:
        return sheets
    preferred = [s for s in line_like if detect_model_from_sheet(s) != "UNKNOWN"]
    return preferred if preferred else line_like


_BATCH_RE = re.compile(r"^(?:D?[AB]\d{4}-\d{3}|[A-Za-z]+\d{3,8}-\d{2,4})$")


def _col_contains(col: str, token: str) -> bool:
    def _normalize_for_match(text: str) -> str:
        s = _sanitize_col(text).lower().replace("μ", "u").replace("µ", "u")
        return re.sub(r"[_\(\)\[\]\{\}]", "", s)

    norm_col = _normalize_for_match(col)
    norm_token = _normalize_for_match(token)
    return (token in str(col)) or (norm_token in norm_col)


def _find_row_contains(df: pd.DataFrame, keyword: str, max_rows: int = 80) -> Optional[int]:
    for i in range(min(len(df), max_rows)):
        row = df.iloc[i].astype(str)
        if row.str.contains(keyword, na=False).any():
            return i
    return None


def _is_date_like(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (dt.date, dt.datetime, pd.Timestamp)):
        return True
    if isinstance(value, float) and np.isnan(value):
        return False
    s = _norm_cell(value)
    if not s:
        return False
    # 纯短数字更可能是工艺参数，不当作日期
    if re.fullmatch(r"\d{1,4}", s):
        return False
    # Excel 序列日期常见区间（约 1954~2091）
    if re.fullmatch(r"\d+(?:\.\d+)?", s):
        try:
            num = float(s)
            return 20000 <= num <= 70000
        except Exception:
            return False
    parsed = pd.to_datetime(s, errors="coerce")
    return not pd.isna(parsed)


def _find_data_start(df: pd.DataFrame, header_start: int = 0, max_rows: int = 240) -> Optional[int]:
    scan_start = max(0, int(header_start))
    scan_end = min(len(df), max_rows)

    # 1) 优先按批次格式命中（兼容 DA2601-001 / S006001-001 等）
    for i in range(scan_start, scan_end):
        row = df.iloc[i].astype(str)
        if row.str.match(_BATCH_RE, na=False).any():
            return i

    # 2) 回退：首列是日期，次列是非空批次文本
    for i in range(scan_start, scan_end):
        if df.shape[1] < 2:
            break
        row = df.iloc[i]
        c0 = row.iloc[0]
        c1 = _norm_cell(row.iloc[1])
        if not c1:
            continue
        # 排除明显表头关键词
        if any(k in c1 for k in ("批次", "线别", "窑炉温度", "D10", "D50", "D90", "均值", "时间")):
            continue
        if _is_date_like(c0):
            return i
    return None


def _make_columns_from_multirow_header(header_df: pd.DataFrame) -> list[str]:
    filled_rows = [_ffill_right(header_df.iloc[i].tolist()) for i in range(len(header_df))]
    cols: list[str] = []
    for j in range(header_df.shape[1]):
        parts: list[str] = []
        for i in range(len(filled_rows)):
            s = _norm_cell(filled_rows[i][j])
            if s and (not parts or s != parts[-1]):
                parts.append(s)
        cols.append(_sanitize_col("_".join(parts)) if parts else f"COL{j}")
    return _dedupe(cols)


def _read_excel_raw(path: str, sheet_name: str) -> pd.DataFrame:
    # 使用 calamine 读 Excel：不依赖样式，兼容性更强（对 WPS/复杂样式表更稳）
    return pd.read_excel(path, sheet_name=sheet_name, header=None, engine="calamine")


def load_sheet_table(path: str, sheet_name: str) -> pd.DataFrame:
    raw = _read_excel_raw(path, sheet_name=sheet_name)

    header_start = _find_row_contains(raw, "批次") or 0
    data_start = _find_data_start(raw, header_start=header_start)
    if data_start is None or data_start <= header_start:
        raise ValueError(f"无法定位数据起始行：{path} / {sheet_name}")

    header_df = raw.iloc[header_start:data_start, :]
    cols = _make_columns_from_multirow_header(header_df)

    df = raw.iloc[data_start:, :].copy()
    df.columns = cols
    df = df.dropna(how="all")

    if "投料日期" in df.columns:
        # 兼容“日期单元格合并/只在首行填写”的情况：先转 datetime 再向下填充，再转 date
        s = pd.to_datetime(df["投料日期"], errors="coerce")
        s = s.ffill()
        df["投料日期"] = s.dt.date
    if "批次" in df.columns:
        s = df["批次"].astype(str).str.strip().replace({"": np.nan, "nan": np.nan})
        df["批次"] = s.ffill()
    if "是否为验证批次" in df.columns:
        s = df["是否为验证批次"].astype(str).str.strip().replace({"": np.nan, "nan": np.nan})
        df["是否为验证批次"] = s.ffill()
    return df


def _parse_num(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, float) and np.isnan(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    parts = text.split("/")
    nums: list[float] = []
    for part in parts:
        cleaned = (
            part.replace("≤", "")
            .replace("≥", "")
            .replace("＋", "+")
            .replace("。", ".")
            .replace("．", ".")
            .replace("%", "")
            .strip()
        )
        cleaned = re.sub(r"(?<=\d),(?=\d)", ".", cleaned)
        m = re.search(r"[-+]?\d+(?:\.\d+)?", cleaned)
        if m:
            nums.append(float(m.group(0)))
    if not nums:
        return None
    return float(sum(nums) / len(nums))


def _to_num_series(s: pd.Series) -> pd.Series:
    return s.apply(_parse_num)


def _flatten_numeric_values(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    if not cols:
        return pd.Series(dtype=float)
    series_list = [_to_num_series(df[c]) for c in cols if c in df.columns]
    if not series_list:
        return pd.Series(dtype=float)
    values = pd.concat(series_list, axis=0).dropna()
    return values


@dataclass(frozen=True)
class RangeStat:
    min: float
    max: float
    mean: float
    n: int


def _range_stat(values: pd.Series) -> Optional[RangeStat]:
    values = values.dropna()
    if values.empty:
        return None
    return RangeStat(
        min=float(values.min()),
        max=float(values.max()),
        mean=float(values.mean()),
        n=int(values.shape[0]),
    )


@dataclass(frozen=True)
class Spec:
    lower: Optional[float]
    upper: Optional[float]
    text: str


_NUM = r"[-+]?\d+(?:[.,]\d+)?"


def _to_float(text: str) -> float:
    return float(text.replace(",", "."))


def parse_spec_from_colname(col: str) -> Optional[Spec]:
    # 从列名里解析常见内控口径：≥x、≤x、a~b、a-b、x±y
    # 注意：这里只用于“是否超内控”的粗判定；真正口径建议从“内控/规格表”统一维护。
    text = _sanitize_col(col).replace("～", "~").replace("—", "-").replace("−", "-")
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)

    m = re.search(rf"(?:≥|>=)\s*({_NUM})", text)
    if m:
        return Spec(lower=_to_float(m.group(1)), upper=None, text=f"≥{m.group(1)}")
    m = re.search(rf"(?:≤|<=)\s*({_NUM})", text)
    if m:
        return Spec(lower=None, upper=_to_float(m.group(1)), text=f"≤{m.group(1)}")
    m = re.search(rf"({_NUM})\s*[~-]\s*({_NUM})", text)
    if m:
        a = _to_float(m.group(1))
        b = _to_float(m.group(2))
        lo, hi = (a, b) if a <= b else (b, a)
        return Spec(lower=lo, upper=hi, text=f"{lo}-{hi}")
    m = re.search(rf"({_NUM})\s*±\s*({_NUM})", text)
    if m:
        x = _to_float(m.group(1))
        y = _to_float(m.group(2))
        return Spec(lower=x - y, upper=x + y, text=f"{x}±{y}")
    return None


def judge_out_of_spec(values: pd.Series, spec: Optional[Spec]) -> Optional[dict[str, Any]]:
    if spec is None:
        return None
    values = values.dropna()
    if values.empty:
        return None
    out_low = spec.lower is not None and (values < spec.lower).any()
    out_high = spec.upper is not None and (values > spec.upper).any()
    if not (out_low or out_high):
        return {"异常": False, "口径": spec.text}
    return {
        "异常": True,
        "口径": spec.text,
        "低于下限": bool(out_low),
        "高于上限": bool(out_high),
    }


def _has_any_metric_data(df: pd.DataFrame, date_: dt.date) -> bool:
    if "投料日期" not in df.columns:
        return False
    day = df[df["投料日期"] == date_]
    if day.empty:
        return False

    patterns = [
        "成品压实",
        "粉末电阻",
        "碳含量",
        "麦克比表",
        "Li+含量",
        "0.1C充",
        "0.1C放",
        "0.1C首效",
        "3.2V容量占比",
        "烧结压实",
        "粉碎压实",
    ]
    for pat in patterns:
        cols = [c for c in df.columns if _col_contains(c, pat)]
        values = _flatten_numeric_values(day, cols)
        if not values.empty:
            return True
    return False


def pick_date_from_dfs(dfs: Iterable[pd.DataFrame], date_arg: Optional[str]) -> dt.date:
    if date_arg:
        return pd.to_datetime(date_arg).date()

    dates: set[dt.date] = set()
    valid_dfs: list[pd.DataFrame] = [df for df in dfs if isinstance(df, pd.DataFrame)]
    for df in valid_dfs:
        if "投料日期" in df.columns:
            dates |= set([d for d in df["投料日期"].dropna().tolist() if isinstance(d, dt.date)])
    if not dates:
        raise ValueError("无法从Excel中推断最新日期，请使用 --date 指定。")

    # 优先选择“最新且有关键指标数据”的日期（避免最新投料日期只有制程前段录入，成品/扣电还未出数）
    for d in sorted(dates, reverse=True):
        if any(_has_any_metric_data(df, d) for df in valid_dfs):
            return d

    # 兜底：取最大日期
    return max(dates)


def pick_date(df_a: pd.DataFrame, df_b: pd.DataFrame, date_arg: Optional[str]) -> dt.date:
    return pick_date_from_dfs([df_a, df_b], date_arg)


def _stat_for_cols(day: pd.DataFrame, cols: list[str], scale: float = 1.0, enable_spec: bool = True) -> dict[str, Any]:
    values = _flatten_numeric_values(day, cols)
    if scale != 1.0:
        values = values * scale
    st = _range_stat(values)
    if st is None:
        return {"有数据": False}
    batch_judge = _batch_out_of_spec_summary(day, cols, scale=scale, enable_spec=enable_spec)
    spec = parse_spec_from_colname(cols[0]) if (cols and enable_spec) else None
    if spec is not None and scale != 1.0:
        spec = Spec(
            lower=None if spec.lower is None else spec.lower * scale,
            upper=None if spec.upper is None else spec.upper * scale,
            text=spec.text,
        )
    judge = judge_out_of_spec(values, spec)
    if isinstance(batch_judge, dict):
        abnormal_batches = int(batch_judge.get("异常批次", 0))
        judge = {"异常": abnormal_batches > 0, "口径": "按列规格(按批次汇总)"}
    return {
        "有数据": True,
        "min": st.min,
        "max": st.max,
        "mean": st.mean,
        "n": st.n,
        "判异": judge,
        "批次判异": batch_judge,
    }


_BATCH_PARSE_RE = re.compile(r"^(?P<prefix>[A-Za-z0-9-]+?)(?P<yymm>\d{4})-(?P<seq>\d{3})$")


def _batch_sort_key(batch: str) -> tuple[Any, ...]:
    b = (batch or "").strip()
    m = _BATCH_PARSE_RE.match(b)
    if not m:
        return (b, 0, 0)
    return (m.group("prefix"), int(m.group("yymm")), int(m.group("seq")))


def _split_consecutive(nums: list[int]) -> list[tuple[int, int]]:
    if not nums:
        return []
    out: list[tuple[int, int]] = []
    start = nums[0]
    prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        out.append((start, prev))
        start = n
        prev = n
    out.append((start, prev))
    return out


def _format_seq_span(start: int, end: int) -> str:
    if start == end:
        return f"{start:03d}"
    if end == start + 1:
        return f"{start:03d}-{end:03d}"
    return f"{start:03d}~{end:03d}"


def _summarize_batches(batches: list[str]) -> str:
    batches = [b.strip() for b in batches if isinstance(b, str) and b.strip() and b.strip().lower() != "nan"]
    if not batches:
        return ""
    batches = sorted(list(dict.fromkeys(batches)), key=_batch_sort_key)
    if len(batches) == 1:
        return batches[0]

    parsed: list[tuple[str, int, int]] = []
    for b in batches:
        m = _BATCH_PARSE_RE.match(b)
        if not m:
            # 存在无法解析的批次号时，退回到逐个列举，避免误判成连续区间
            return "、".join(batches)
        parsed.append((m.group("prefix"), int(m.group("yymm")), int(m.group("seq"))))

    # 同一前缀+年月内：
    # - 连续：DA2602-103-104 / S006-DB2602-074~079
    # - 不连续：DA2602-103、105
    result_parts: list[str] = []
    keys: list[tuple[str, int]] = []
    grouped: dict[tuple[str, int], list[int]] = {}
    for prefix, yymm, seq in parsed:
        key = (prefix, yymm)
        if key not in grouped:
            grouped[key] = []
            keys.append(key)
        grouped[key].append(seq)

    for prefix, yymm in keys:
        seqs = sorted(set(grouped[(prefix, yymm)]))
        spans = _split_consecutive(seqs)
        if not spans:
            continue

        first_start, first_end = spans[0]
        head = f"{prefix}{yymm:04d}-{_format_seq_span(first_start, first_end)}"
        if len(spans) == 1:
            result_parts.append(head)
            continue

        tail = [_format_seq_span(s, e) for s, e in spans[1:]]
        result_parts.append("、".join([head] + tail))

    return "、".join(result_parts)


def _batch_summary_for_day(day: pd.DataFrame, cols: list[str]) -> str:
    if day.empty or "批次" not in day.columns:
        return ""
    metric_cols = [c for c in cols if c in day.columns]
    if not metric_cols:
        return ""

    has_value = pd.Series(False, index=day.index)
    for c in metric_cols:
        s = _to_num_series(day[c])
        has_value = has_value | s.notna()

    if not has_value.any():
        return ""

    batches = day.loc[has_value, "批次"].dropna().astype(str).tolist()
    return _summarize_batches(batches)


def _batch_out_of_spec_summary(
    day: pd.DataFrame, cols: list[str], scale: float = 1.0, enable_spec: bool = True
) -> Optional[dict[str, int]]:
    if day.empty or "批次" not in day.columns or not enable_spec:
        return None
    metric_cols = [c for c in cols if c in day.columns]
    if not metric_cols:
        return None

    batch_series = day["批次"].astype(str).str.strip()
    valid_batch_mask = (batch_series != "") & (batch_series.str.lower() != "nan")

    has_value = pd.Series(False, index=day.index)
    for c in metric_cols:
        has_value = has_value | _to_num_series(day[c]).notna()
    total_batches = set(batch_series[has_value & valid_batch_mask].tolist())

    abnormal_batches: set[str] = set()
    low_batches: set[str] = set()
    high_batches: set[str] = set()
    spec_cols = 0

    for c in metric_cols:
        spec = parse_spec_from_colname(c)
        if spec is None:
            continue
        spec_cols += 1
        if scale != 1.0:
            spec = Spec(
                lower=None if spec.lower is None else spec.lower * scale,
                upper=None if spec.upper is None else spec.upper * scale,
                text=spec.text,
            )

        values = _to_num_series(day[c])
        if scale != 1.0:
            values = values * scale
        valid_value_mask = values.notna() & valid_batch_mask
        if not valid_value_mask.any():
            continue

        values = values[valid_value_mask]
        value_batches = batch_series[valid_value_mask]

        if spec.lower is not None:
            low_mask = values < spec.lower
            if low_mask.any():
                hit = set(value_batches[low_mask].tolist())
                low_batches |= hit
                abnormal_batches |= hit
        if spec.upper is not None:
            high_mask = values > spec.upper
            if high_mask.any():
                hit = set(value_batches[high_mask].tolist())
                high_batches |= hit
                abnormal_batches |= hit

    if spec_cols <= 0:
        return None

    return {
        "总批次": int(len(total_batches)),
        "异常批次": int(len(abnormal_batches)),
        "低于下限批次": int(len(low_batches)),
        "高于上限批次": int(len(high_batches)),
        "口径列数": int(spec_cols),
    }


def _latest_stat_within_days(
    df: pd.DataFrame,
    cols: list[str],
    report_date: dt.date,
    lookback_days: int,
    scale: float = 1.0,
    enable_spec: bool = True,
) -> dict[str, Any]:
    if "投料日期" not in df.columns:
        return {"有数据": False}

    for delta in range(0, max(0, lookback_days) + 1):
        d = report_date - dt.timedelta(days=delta)
        day = df[df["投料日期"] == d]
        st = _stat_for_cols(day, cols, scale=scale, enable_spec=enable_spec)
        if st.get("有数据"):
            st["来源日期"] = d
            st["是否当日"] = (delta == 0)
            st["来源批次摘要"] = _batch_summary_for_day(day, cols)
            return st
    return {"有数据": False}


_LI_PASS_MAX = 500
_TREND_ANOMALY_RATIO = 1.3
_LI_TREND_ANOMALY_ABS = 80
_TREND_MAX_LOOKBACK_DAYS = 60
_TREND_SIGNIFICANT_REL = 0.005
_TREND_SIGNIFICANT_REL_LI = 0.02
_TREND_SIGNIFICANT_ABS_LI = 10.0


def _trend_for_cols(
    df: pd.DataFrame,
    cols: list[str],
    report_date: dt.date,
    trend_points: int,
    max_lookback_days: int,
    scale: float = 1.0,
    anomaly_ratio: float = _TREND_ANOMALY_RATIO,
    anomaly_abs: Optional[float] = None,
    enable_spec: bool = True,
) -> dict[str, Any]:
    if "投料日期" not in df.columns or not cols or trend_points <= 0:
        return {"有数据": False, "窗口": trend_points, "点数": 0}

    points: list[tuple[dt.date, dict[str, Any]]] = []
    for delta in range(0, max_lookback_days + 1):
        d = report_date - dt.timedelta(days=delta)
        day = df[df["投料日期"] == d]
        st = _stat_for_cols(day, cols, scale=scale, enable_spec=enable_spec)
        if st.get("有数据"):
            points.append((d, st))
            if len(points) >= trend_points:
                break

    if not points:
        return {"有数据": False, "窗口": trend_points, "点数": 0}

    points = list(reversed(points))
    means = [p[1]["mean"] for p in points]
    overall_mean = float(sum(means) / len(means)) if means else float("nan")
    anomalies: list[dict[str, Any]] = []
    if overall_mean and not np.isnan(overall_mean):
        lower = overall_mean / anomaly_ratio if anomaly_ratio > 0 else overall_mean
        upper = overall_mean * anomaly_ratio
        for d, st in points:
            mean = float(st["mean"])
            if mean < lower or mean > upper or (anomaly_abs is not None and abs(mean - overall_mean) >= anomaly_abs):
                anomalies.append({"日期": d, "mean": mean})

    direction = "—"
    if means[-1] > means[0] + 1e-9:
        direction = "↑"
    elif means[-1] < means[0] - 1e-9:
        direction = "↓"

    return {
        "有数据": True,
        "窗口": trend_points,
        "点数": len(points),
        "日期": [p[0] for p in points],
        "均值": means,
        "方向": direction,
        "异常": anomalies,
    }


def extract_metrics(
    df: pd.DataFrame,
    report_date: dt.date,
    lookback_days: int,
    trend_days: int = 7,
    enable_spec: bool = True,
) -> dict[str, Any]:
    if "投料日期" in df.columns:
        day = df[df["投料日期"] == report_date].copy()
    else:
        day = df.head(0).copy()

    # 注意：很多表会出现“均值未填，但B1-1/B1-2/B1-3已填”的情况，因此这里不只取“均值”
    sinter_cols = [c for c in df.columns if "烧结压实" in c]
    crush_cols = [c for c in df.columns if "粉碎压实" in c]

    prod_density_cols = [c for c in df.columns if "成品压实" in c]
    powder_res_cols = [c for c in df.columns if _col_contains(c, "粉末电阻")]
    li_cols = [c for c in df.columns if _col_contains(c, "Li+含量")]
    carbon_cols = [c for c in df.columns if c.startswith("碳含量")]
    if not carbon_cols:
        carbon_cols = [c for c in df.columns if _col_contains(c, "碳含量") and not _col_contains(c, "粉碎碳含量")]
    bet_cols = [c for c in df.columns if "麦克比表" in c]

    charge_cols = [c for c in df.columns if _col_contains(c, "0.1C充")]
    discharge_cols = [c for c in df.columns if _col_contains(c, "0.1C放")]
    eff_cols = [c for c in df.columns if _col_contains(c, "0.1C首效")]

    plat_cols = [c for c in df.columns if _col_contains(c, "3.2V容量占比")]
    if not plat_cols:
        plat_cols = [c for c in df.columns if _col_contains(c, "3.2V平台效率")]
    if not plat_cols:
        plat_cols = [c for c in df.columns if _col_contains(c, "2.95V容量占比")]

    trend_lookback_days = max(lookback_days, trend_days * 5, 30)
    trend_lookback_days = min(trend_lookback_days, _TREND_MAX_LOOKBACK_DAYS)

    def _trend(cols: list[str], scale: float = 1.0, anomaly_abs: Optional[float] = None) -> dict[str, Any]:
        return _trend_for_cols(
            df,
            cols,
            report_date,
            trend_days,
            trend_lookback_days,
            scale=scale,
            anomaly_abs=anomaly_abs,
            enable_spec=enable_spec,
        )

    li_trend = _trend(li_cols, scale=10000, anomaly_abs=_LI_TREND_ANOMALY_ABS)
    sinter_trend = _trend(sinter_cols)
    crush_trend = _trend(crush_cols)
    prod_density_trend = _trend(prod_density_cols)
    powder_res_trend = _trend(powder_res_cols)
    carbon_trend = _trend(carbon_cols)
    bet_trend = _trend(bet_cols)
    charge_trend = _trend(charge_cols)
    discharge_trend = _trend(discharge_cols)
    eff_trend = _trend(eff_cols)
    plat_trend = _trend(plat_cols)

    return {
        "制程": {
            "烧结压实": _latest_stat_within_days(df, sinter_cols, report_date, lookback_days, enable_spec=enable_spec),
            "粉碎压实": _latest_stat_within_days(df, crush_cols, report_date, lookback_days, enable_spec=enable_spec),
            "烧结压实趋势": sinter_trend,
            "粉碎压实趋势": crush_trend,
        },
        "成品": {
            "成品压实": _latest_stat_within_days(
                df, prod_density_cols, report_date, lookback_days, enable_spec=enable_spec
            ),
            "0.1C充电": _latest_stat_within_days(df, charge_cols, report_date, lookback_days, enable_spec=enable_spec),
            "0.1C放电": _latest_stat_within_days(
                df, discharge_cols, report_date, lookback_days, enable_spec=enable_spec
            ),
            "首效": _latest_stat_within_days(df, eff_cols, report_date, lookback_days, enable_spec=enable_spec),
            "平台效率": _latest_stat_within_days(df, plat_cols, report_date, lookback_days, enable_spec=enable_spec),
            "残碱(Li+)": _latest_stat_within_days(
                df, li_cols, report_date, lookback_days, scale=10000, enable_spec=enable_spec
            ),
            "碳含量": _latest_stat_within_days(df, carbon_cols, report_date, lookback_days, enable_spec=enable_spec),
            "粉阻(粉末电阻)": _latest_stat_within_days(
                df, powder_res_cols, report_date, lookback_days, enable_spec=enable_spec
            ),
            "比表(麦克比表)": _latest_stat_within_days(df, bet_cols, report_date, lookback_days, enable_spec=enable_spec),
            "残碱(Li+)趋势": li_trend,
            "成品压实趋势": prod_density_trend,
            "0.1C充电趋势": charge_trend,
            "0.1C放电趋势": discharge_trend,
            "首效趋势": eff_trend,
            "平台效率趋势": plat_trend,
            "碳含量趋势": carbon_trend,
            "粉阻(粉末电阻)趋势": powder_res_trend,
            "比表(麦克比表)趋势": bet_trend,
        },
        "当日行数": int(day.shape[0]),
    }


def _merge_stats(a: dict[str, Any], b: dict[str, Any]) -> dict[str, Any]:
    if not a.get("有数据") and not b.get("有数据"):
        return {"有数据": False}

    mins: list[float] = []
    maxs: list[float] = []
    means: list[float] = []
    ns: list[int] = []
    for st in (a, b):
        if st.get("有数据"):
            mins.append(float(st["min"]))
            maxs.append(float(st["max"]))
            means.append(float(st["mean"]))
            ns.append(int(st["n"]))

    total_n = sum(ns) if ns else 0
    mean = sum(m * n for m, n in zip(means, ns)) / total_n if total_n else float("nan")
    return {
        "有数据": True,
        "min": float(min(mins)),
        "max": float(max(maxs)),
        "mean": float(mean),
        "n": int(total_n),
    }


def _fmt_range(stat: dict[str, Any], decimals: int = 3) -> str:
    if not stat.get("有数据"):
        return "未找到有效数据"
    mn = float(stat["min"])
    mx = float(stat["max"])
    if mn == mx:
        return f"{mn:.{decimals}f}"
    return f"{mn:.{decimals}f}-{mx:.{decimals}f}"


def _fmt_status(stat: dict[str, Any], force: bool = False) -> str:
    judge = stat.get("判异") if isinstance(stat, dict) else None
    if not isinstance(judge, dict):
        if force and stat.get("有数据"):
            return "（未设定口径）"
        return ""
    abnormal = judge.get("异常")
    batch_judge = stat.get("批次判异") if isinstance(stat, dict) else None
    if abnormal is True and isinstance(batch_judge, dict):
        total = int(batch_judge.get("总批次", 0))
        abnormal_cnt = int(batch_judge.get("异常批次", 0))
        if total > 0 and abnormal_cnt > 0:
            low_cnt = int(batch_judge.get("低于下限批次", 0))
            high_cnt = int(batch_judge.get("高于上限批次", 0))
            direction_parts: list[str] = []
            if low_cnt > 0:
                direction_parts.append("低于下限")
            if high_cnt > 0:
                direction_parts.append("高于上限")
            direction = f"，{'/'.join(direction_parts)}" if direction_parts else ""
            return f"（异常，{abnormal_cnt}/{total}批次超规{direction}）"
    if abnormal is True:
        return "（异常）"
    if abnormal is False:
        return "（正常）"
    return ""


def _fmt_source_date(stat: dict[str, Any]) -> str:
    d = stat.get("来源日期")
    if isinstance(d, dt.date) and not stat.get("是否当日", False):
        date_str = d.strftime("%Y.%m.%d")
        batch_summary = stat.get("来源批次摘要")
        if isinstance(batch_summary, str) and batch_summary:
            return f"（{date_str}投料批次{batch_summary}）"
        return f"（{date_str}投料批次未知）"
    return ""


def _with_unit(value: str, unit: str) -> str:
    return value if value == "未找到有效数据" else f"{value}{unit}"


def _fmt_metric(stat: dict[str, Any], decimals: int, force_status: bool = False) -> str:
    return _fmt_range(stat, decimals) + _fmt_status(stat, force=force_status) + _fmt_source_date(stat)


def _fmt_li_status(stat: dict[str, Any]) -> str:
    if not stat.get("有数据"):
        return ""
    return "（合格）" if float(stat["max"]) < _LI_PASS_MAX else "（超标）"


def _trend_sparkline(values: list[float]) -> str:
    if not values:
        return ""
    levels = "▁▂▃▄▅▆▇█"
    vmin = min(values)
    vmax = max(values)
    if vmin == vmax:
        return levels[0] * len(values)
    span = vmax - vmin
    chars: list[str] = []
    for v in values:
        idx = int(round((v - vmin) / span * (len(levels) - 1)))
        idx = max(0, min(len(levels) - 1, idx))
        chars.append(levels[idx])
    return "".join(chars)


def _fmt_trend_dates(trend: dict[str, Any]) -> str:
    dates = trend.get("日期") or []
    if not dates:
        return ""
    sorted_dates = sorted([d for d in dates if isinstance(d, dt.date)])
    if not sorted_dates:
        return ""
    is_continuous = all(
        (sorted_dates[i] - sorted_dates[i - 1]).days == 1 for i in range(1, len(sorted_dates))
    )
    if is_continuous:
        if len(sorted_dates) == 1:
            return sorted_dates[0].strftime("%m.%d")
        return f"{sorted_dates[0].strftime('%m.%d')}-{sorted_dates[-1].strftime('%m.%d')}"
    return "/".join(d.strftime("%m.%d") for d in sorted_dates)


def _trend_significant(trend: dict[str, Any], rel_threshold: float, abs_threshold: Optional[float]) -> bool:
    if not trend.get("有数据"):
        return False
    means = trend.get("均值") or []
    if len(means) < 2:
        return False
    mn = min(means)
    mx = max(means)
    amplitude = mx - mn
    if abs_threshold is not None and amplitude >= abs_threshold:
        return True
    mean = sum(means) / len(means) if means else 0.0
    if mean and amplitude / mean >= rel_threshold:
        return True
    return False


def _fmt_trend_brief(trend: dict[str, Any], decimals: int, unit: str = "") -> str:
    if not trend.get("有数据"):
        return "无数据"
    means = trend.get("均值") or []
    if not means:
        return "无数据"
    fmt = lambda v: f"{v:.{decimals}f}"
    mn = min(means)
    mx = max(means)
    direction = trend.get("方向", "—")
    if direction == "↑":
        word = "微升"
    elif direction == "↓":
        word = "微降"
    else:
        word = "稳定"
    range_str = f"{fmt(mn)}~{fmt(mx)}"
    if unit:
        range_str = f"{range_str}{unit}"
    dates = _fmt_trend_dates(trend)
    suffix = f" [{dates}]" if dates else ""
    return f"{word}（{range_str}）{suffix}"


def _fmt_trend_layered(
    trend: dict[str, Any], decimals: int, unit: str, rel_threshold: float, abs_threshold: Optional[float]
) -> str:
    if _trend_significant(trend, rel_threshold, abs_threshold):
        return _fmt_trend(trend, decimals, unit)
    return _fmt_trend_brief(trend, decimals, unit)


def _fmt_trend(trend: dict[str, Any], decimals: int, unit: str = "") -> str:
    if not trend.get("有数据"):
        return "无数据"
    means = trend.get("均值") or []
    if not means:
        return "无数据"
    fmt = lambda v: f"{v:.{decimals}f}"
    seq = ",".join(fmt(v) for v in means)
    spark = _trend_sparkline(means)
    base = f"{seq} / {spark}（{trend.get('方向', '—')}）"
    if unit:
        base = f"{base}{unit}"
    dates = _fmt_trend_dates(trend)
    if dates:
        base = f"{base}（日期{dates}）"
    anomalies = trend.get("异常") or []
    if not anomalies:
        return base
    first = anomalies[0]
    date_str = first["日期"].strftime("%Y.%m.%d") if isinstance(first.get("日期"), dt.date) else "未知日期"
    hint = f"异常：{date_str} {fmt(first['mean'])}"
    if unit:
        hint = f"{hint}{unit}"
    if len(anomalies) > 1:
        hint += f" 等{len(anomalies)}天"
    return f"{base}，{hint}"


def build_wecom_text(report_date: dt.date, a: dict[str, Any], b: dict[str, Any]) -> str:
    date_str = report_date.strftime("%Y.%m.%d")

    a_sinter_stat = a["制程"]["烧结压实"]
    b_sinter_stat = b["制程"]["烧结压实"]
    a_sinter = _fmt_metric(a_sinter_stat, 3, force_status=True)
    b_sinter = _fmt_metric(b_sinter_stat, 3, force_status=False)
    ab_sinter = _fmt_range(_merge_stats(a_sinter_stat, b_sinter_stat), 3)

    a_crush_stat = a["制程"]["粉碎压实"]
    b_crush_stat = b["制程"]["粉碎压实"]
    a_crush = _fmt_metric(a_crush_stat, 3, force_status=True)
    b_crush = _fmt_metric(b_crush_stat, 3, force_status=False)
    ab_crush = _fmt_range(_merge_stats(a_crush_stat, b_crush_stat), 3)
    a_sinter_trend = _fmt_trend_layered(a["制程"]["烧结压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    b_sinter_trend = _fmt_trend_layered(b["制程"]["烧结压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    a_crush_trend = _fmt_trend_layered(a["制程"]["粉碎压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    b_crush_trend = _fmt_trend_layered(b["制程"]["粉碎压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)

    a_prod_density_stat = a["成品"]["成品压实"]
    b_prod_density_stat = b["成品"]["成品压实"]
    ab_prod_density = _fmt_range(_merge_stats(a_prod_density_stat, b_prod_density_stat), 3)

    a_charge = _fmt_metric(a["成品"]["0.1C充电"], 1, force_status=True)
    a_discharge = _fmt_metric(a["成品"]["0.1C放电"], 1, force_status=True)
    a_eff = _fmt_metric(a["成品"]["首效"], 2, force_status=True)
    a_plat = _fmt_metric(a["成品"]["平台效率"], 1, force_status=True)

    b_charge = _fmt_metric(b["成品"]["0.1C充电"], 1, force_status=False)
    b_discharge = _fmt_metric(b["成品"]["0.1C放电"], 1, force_status=False)
    b_eff = _fmt_metric(b["成品"]["首效"], 2, force_status=False)
    b_plat = _fmt_metric(b["成品"]["平台效率"], 1, force_status=False)

    ab_alkali_stat = _merge_stats(a["成品"]["残碱(Li+)"], b["成品"]["残碱(Li+)"])
    ab_alkali = _fmt_range(ab_alkali_stat, 0)
    if LINE_B_SPEC_ENABLED:
        ab_alkali_status = _fmt_li_status(ab_alkali_stat)
    else:
        ab_alkali_status = "（规格口径待更新）" if ab_alkali_stat.get("有数据") else ""
    a_carbon = _fmt_metric(a["成品"]["碳含量"], 2, force_status=True)
    b_carbon = _fmt_metric(b["成品"]["碳含量"], 2, force_status=False)
    a_powder_r = _fmt_metric(a["成品"]["粉阻(粉末电阻)"], 1, force_status=True)
    b_powder_r = _fmt_metric(b["成品"]["粉阻(粉末电阻)"], 1, force_status=False)
    a_bet = _fmt_metric(a["成品"]["比表(麦克比表)"], 1, force_status=True)
    b_bet = _fmt_metric(b["成品"]["比表(麦克比表)"], 1, force_status=False)
    a_prod_trend = _fmt_trend_layered(a["成品"]["成品压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    b_prod_trend = _fmt_trend_layered(b["成品"]["成品压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    a_charge_trend = _fmt_trend_layered(a["成品"]["0.1C充电趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    b_charge_trend = _fmt_trend_layered(b["成品"]["0.1C充电趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    a_discharge_trend = _fmt_trend_layered(a["成品"]["0.1C放电趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    b_discharge_trend = _fmt_trend_layered(b["成品"]["0.1C放电趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    a_eff_trend = _fmt_trend_layered(a["成品"]["首效趋势"], 2, "", _TREND_SIGNIFICANT_REL, None)
    b_eff_trend = _fmt_trend_layered(b["成品"]["首效趋势"], 2, "", _TREND_SIGNIFICANT_REL, None)
    a_plat_trend = _fmt_trend_layered(a["成品"]["平台效率趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    b_plat_trend = _fmt_trend_layered(b["成品"]["平台效率趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    a_li_trend = _fmt_trend_layered(
        a["成品"]["残碱(Li+)趋势"], 0, "ppm", _TREND_SIGNIFICANT_REL_LI, _TREND_SIGNIFICANT_ABS_LI
    )
    b_li_trend = _fmt_trend_layered(
        b["成品"]["残碱(Li+)趋势"], 0, "ppm", _TREND_SIGNIFICANT_REL_LI, _TREND_SIGNIFICANT_ABS_LI
    )
    a_carbon_trend = _fmt_trend_layered(a["成品"]["碳含量趋势"], 2, "", _TREND_SIGNIFICANT_REL, None)
    b_carbon_trend = _fmt_trend_layered(b["成品"]["碳含量趋势"], 2, "", _TREND_SIGNIFICANT_REL, None)
    a_powder_trend = _fmt_trend_layered(a["成品"]["粉阻(粉末电阻)趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    b_powder_trend = _fmt_trend_layered(b["成品"]["粉阻(粉末电阻)趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    a_bet_trend = _fmt_trend_layered(a["成品"]["比表(麦克比表)趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    b_bet_trend = _fmt_trend_layered(b["成品"]["比表(麦克比表)趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)

    # 输出结构按你给的 1~6 段落口径；未提供的数据先保留占位，便于你后续把“第二张表”接进来。
    lines: list[str] = []
    lines.append(f"{date_str}数据表更新：")
    if not LINE_B_SPEC_ENABLED:
        lines.append(f"注：{LINE_B_LABEL}已更换产品，规格口径待更新，当前仅展示数据区间，不做超规判定。")
    lines.append("1、原料bom：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append("2、配方：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append(
        f"3、制程：烧结压实({LINE_A_LABEL}+{LINE_B_LABEL}) {ab_sinter}；{LINE_A_LABEL} {a_sinter}；{LINE_B_LABEL} {b_sinter}。"
        f"粉碎压实({LINE_A_LABEL}+{LINE_B_LABEL}) {ab_crush}；{LINE_A_LABEL} {a_crush}；{LINE_B_LABEL} {b_crush}。"
    )
    trend_points = max(a["制程"]["烧结压实趋势"].get("点数", 0), b["制程"]["烧结压实趋势"].get("点数", 0))
    trend_window = max(a["制程"]["烧结压实趋势"].get("窗口", 0), b["制程"]["烧结压实趋势"].get("窗口", 0))
    if trend_points:
        label = f"近{trend_points}日数据均值"
        if trend_window and trend_points < trend_window:
            label = f"{label}（不足{trend_window}日数据）"
        lines.append(f"  制程趋势（{label}）：")
        lines.append(f"    烧结压实 {LINE_A_LABEL} {a_sinter_trend}；{LINE_B_LABEL} {b_sinter_trend}。")
        lines.append(f"    粉碎压实 {LINE_A_LABEL} {a_crush_trend}；{LINE_B_LABEL} {b_crush_trend}。")
    lines.append("4、成品：")
    lines.append(f"  ①{LINE_A_LABEL}+{LINE_B_LABEL}成品压实：{ab_prod_density}。")
    lines.append(f"  ②0.1C充电：{LINE_A_LABEL} {a_charge}；{LINE_B_LABEL} {b_charge}。")
    lines.append(f"0.1C放电：{LINE_A_LABEL} {a_discharge}；{LINE_B_LABEL} {b_discharge}。")
    lines.append(f"首效：{LINE_A_LABEL} {a_eff}；{LINE_B_LABEL} {b_eff}。")
    lines.append(f"平台效率：{LINE_A_LABEL} {a_plat}；{LINE_B_LABEL} {b_plat}。")
    lines.append(f"  ③残碱(Li+)：{LINE_A_LABEL}+{LINE_B_LABEL} {_with_unit(ab_alkali, 'ppm')}{ab_alkali_status}。")
    lines.append(f"  ④碳含量：{LINE_A_LABEL} {a_carbon}；{LINE_B_LABEL} {b_carbon}。")
    lines.append(f"  ⑤粉阻(粉末电阻)：{LINE_A_LABEL} {a_powder_r}；{LINE_B_LABEL} {b_powder_r}。")
    lines.append(f"  ⑥比表(麦克比表)：{LINE_A_LABEL} {a_bet}；{LINE_B_LABEL} {b_bet}。")
    trend_points = max(a["成品"]["残碱(Li+)趋势"].get("点数", 0), b["成品"]["残碱(Li+)趋势"].get("点数", 0))
    trend_window = max(a["成品"]["残碱(Li+)趋势"].get("窗口", 0), b["成品"]["残碱(Li+)趋势"].get("窗口", 0))
    if trend_points:
        label = f"近{trend_points}日数据均值"
        if trend_window and trend_points < trend_window:
            label = f"{label}（不足{trend_window}日数据）"
        lines.append(f"  成品趋势（{label}）：")
        lines.append(f"    成品压实 {LINE_A_LABEL} {a_prod_trend}；{LINE_B_LABEL} {b_prod_trend}。")
        lines.append(f"    0.1C充电 {LINE_A_LABEL} {a_charge_trend}；{LINE_B_LABEL} {b_charge_trend}。")
        lines.append(f"    0.1C放电 {LINE_A_LABEL} {a_discharge_trend}；{LINE_B_LABEL} {b_discharge_trend}。")
        lines.append(f"    首效 {LINE_A_LABEL} {a_eff_trend}；{LINE_B_LABEL} {b_eff_trend}。")
        lines.append(f"    平台效率 {LINE_A_LABEL} {a_plat_trend}；{LINE_B_LABEL} {b_plat_trend}。")
        lines.append(f"    残碱(Li+) {LINE_A_LABEL} {a_li_trend}；{LINE_B_LABEL} {b_li_trend}。")
        lines.append(f"    碳含量 {LINE_A_LABEL} {a_carbon_trend}；{LINE_B_LABEL} {b_carbon_trend}。")
        lines.append(f"    粉阻(粉末电阻) {LINE_A_LABEL} {a_powder_trend}；{LINE_B_LABEL} {b_powder_trend}。")
        lines.append(f"    比表(麦克比表) {LINE_A_LABEL} {a_bet_trend}；{LINE_B_LABEL} {b_bet_trend}。")
    lines.append("5、下一步计划：本次Excel未包含（可从模板/手工输入/第二张表接入）")
    lines.append("6、工艺验证：本次Excel未包含（待接入第二张表/Sheet）")
    return "\n".join(lines)


def build_wecom_text_single(
    report_date: dt.date, metrics: dict[str, Any], line_label: str, model: str, enable_spec: bool = True
) -> str:
    date_str = report_date.strftime("%Y.%m.%d")

    process = metrics["制程"]
    product = metrics["成品"]

    sinter = _fmt_metric(process["烧结压实"], 3, force_status=True)
    crush = _fmt_metric(process["粉碎压实"], 3, force_status=True)

    prod_density = _fmt_metric(product["成品压实"], 3, force_status=True)
    charge = _fmt_metric(product["0.1C充电"], 1, force_status=True)
    discharge = _fmt_metric(product["0.1C放电"], 1, force_status=True)
    eff = _fmt_metric(product["首效"], 2, force_status=True)
    plat = _fmt_metric(product["平台效率"], 1, force_status=True)

    li_stat = product["残碱(Li+)"]
    li = _with_unit(_fmt_range(li_stat, 0), "ppm") + _fmt_status(li_stat, force=True) + _fmt_source_date(li_stat)
    carbon = _fmt_metric(product["碳含量"], 2, force_status=True)
    powder_r = _fmt_metric(product["粉阻(粉末电阻)"], 1, force_status=True)
    bet = _fmt_metric(product["比表(麦克比表)"], 1, force_status=True)

    sinter_trend = _fmt_trend_layered(process["烧结压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    crush_trend = _fmt_trend_layered(process["粉碎压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    prod_trend = _fmt_trend_layered(product["成品压实趋势"], 3, "", _TREND_SIGNIFICANT_REL, None)
    charge_trend = _fmt_trend_layered(product["0.1C充电趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    discharge_trend = _fmt_trend_layered(product["0.1C放电趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    eff_trend = _fmt_trend_layered(product["首效趋势"], 2, "", _TREND_SIGNIFICANT_REL, None)
    plat_trend = _fmt_trend_layered(product["平台效率趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    li_trend = _fmt_trend_layered(product["残碱(Li+)趋势"], 0, "ppm", _TREND_SIGNIFICANT_REL_LI, _TREND_SIGNIFICANT_ABS_LI)
    carbon_trend = _fmt_trend_layered(product["碳含量趋势"], 2, "", _TREND_SIGNIFICANT_REL, None)
    powder_trend = _fmt_trend_layered(product["粉阻(粉末电阻)趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)
    bet_trend = _fmt_trend_layered(product["比表(麦克比表)趋势"], 1, "", _TREND_SIGNIFICANT_REL, None)

    lines: list[str] = []
    lines.append(f"{date_str}数据表更新（{line_label} / {model}）：")
    if not enable_spec:
        lines.append("注：当前线别未匹配到产品规格口径，仅展示数据区间，不做超规判定。")
    lines.append("1、原料bom：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append("2、配方：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append(f"3、制程：烧结压实 {sinter}；粉碎压实 {crush}。")

    process_trend_points = max(process["烧结压实趋势"].get("点数", 0), process["粉碎压实趋势"].get("点数", 0))
    process_trend_window = max(process["烧结压实趋势"].get("窗口", 0), process["粉碎压实趋势"].get("窗口", 0))
    if process_trend_points:
        label = f"近{process_trend_points}日数据均值"
        if process_trend_window and process_trend_points < process_trend_window:
            label = f"{label}（不足{process_trend_window}日数据）"
        lines.append(f"  制程趋势（{label}）：")
        lines.append(f"    烧结压实 {sinter_trend}。")
        lines.append(f"    粉碎压实 {crush_trend}。")

    lines.append("4、成品：")
    lines.append(f"  ①成品压实：{prod_density}。")
    lines.append(f"  ②0.1C充电：{charge}。")
    lines.append(f"0.1C放电：{discharge}。")
    lines.append(f"首效：{eff}。")
    lines.append(f"平台效率：{plat}。")
    lines.append(f"  ③残碱(Li+)：{li}。")
    lines.append(f"  ④碳含量：{carbon}。")
    lines.append(f"  ⑤粉阻(粉末电阻)：{powder_r}。")
    lines.append(f"  ⑥比表(麦克比表)：{bet}。")

    product_trend_points = max(
        product["成品压实趋势"].get("点数", 0),
        product["0.1C充电趋势"].get("点数", 0),
        product["0.1C放电趋势"].get("点数", 0),
        product["首效趋势"].get("点数", 0),
        product["平台效率趋势"].get("点数", 0),
        product["残碱(Li+)趋势"].get("点数", 0),
        product["碳含量趋势"].get("点数", 0),
        product["粉阻(粉末电阻)趋势"].get("点数", 0),
        product["比表(麦克比表)趋势"].get("点数", 0),
    )
    product_trend_window = max(
        product["成品压实趋势"].get("窗口", 0),
        product["0.1C充电趋势"].get("窗口", 0),
        product["0.1C放电趋势"].get("窗口", 0),
        product["首效趋势"].get("窗口", 0),
        product["平台效率趋势"].get("窗口", 0),
        product["残碱(Li+)趋势"].get("窗口", 0),
        product["碳含量趋势"].get("窗口", 0),
        product["粉阻(粉末电阻)趋势"].get("窗口", 0),
        product["比表(麦克比表)趋势"].get("窗口", 0),
    )
    if product_trend_points:
        label = f"近{product_trend_points}日数据均值"
        if product_trend_window and product_trend_points < product_trend_window:
            label = f"{label}（不足{product_trend_window}日数据）"
        lines.append(f"  成品趋势（{label}）：")
        lines.append(f"    成品压实 {prod_trend}。")
        lines.append(f"    0.1C充电 {charge_trend}。")
        lines.append(f"    0.1C放电 {discharge_trend}。")
        lines.append(f"    首效 {eff_trend}。")
        lines.append(f"    平台效率 {plat_trend}。")
        lines.append(f"    残碱(Li+) {li_trend}。")
        lines.append(f"    碳含量 {carbon_trend}。")
        lines.append(f"    粉阻(粉末电阻) {powder_trend}。")
        lines.append(f"    比表(麦克比表) {bet_trend}。")

    lines.append("5、下一步计划：本次Excel未包含（可从模板/手工输入/第二张表接入）")
    lines.append("6、工艺验证：本次Excel未包含（待接入第二张表/Sheet）")
    return "\n".join(lines)


def build_wecom_text_multi(line_reports: list[dict[str, Any]]) -> str:
    if not line_reports:
        return ""

    valid_reports = [
        r
        for r in line_reports
        if isinstance(r, dict) and isinstance(r.get("metrics"), dict) and isinstance(r.get("line_label"), str)
    ]
    if not valid_reports:
        return ""

    line_labels = [str(r["line_label"]) for r in valid_reports]
    line_metrics = {str(r["line_label"]): r["metrics"] for r in valid_reports}
    line_enable_spec = {str(r["line_label"]): bool(r.get("enable_spec", True)) for r in valid_reports}

    report_dates = [r.get("report_date") for r in valid_reports if isinstance(r.get("report_date"), dt.date)]
    date_set = sorted({d for d in report_dates if isinstance(d, dt.date)})
    if len(date_set) == 1:
        date_str = date_set[0].strftime("%Y.%m.%d")
    else:
        date_str = "/".join(d.strftime("%Y.%m.%d") for d in date_set) if date_set else dt.date.today().strftime("%Y.%m.%d")

    def _metric_parts(section: str, key: str, decimals: int) -> str:
        parts: list[str] = []
        for label in line_labels:
            st = line_metrics[label][section][key]
            parts.append(f"{label} {_fmt_metric(st, decimals, force_status=line_enable_spec[label])}")
        return "；".join(parts)

    def _trend_parts(section: str, key: str, decimals: int, unit: str, rel: float, abs_v: Optional[float]) -> str:
        parts: list[str] = []
        for label in line_labels:
            trend = line_metrics[label][section][key]
            parts.append(f"{label} {_fmt_trend_layered(trend, decimals, unit, rel, abs_v)}")
        return "；".join(parts)

    def _max_points(section: str, key: str) -> int:
        return max((int(line_metrics[label][section][key].get("点数", 0)) for label in line_labels), default=0)

    def _max_window(section: str, key: str) -> int:
        return max((int(line_metrics[label][section][key].get("窗口", 0)) for label in line_labels), default=0)

    lines: list[str] = []
    lines.append(f"{date_str}数据表更新（{ '、'.join(line_labels) }）：")
    no_spec_lines = [label for label in line_labels if not line_enable_spec[label]]
    if no_spec_lines:
        lines.append(f"注：{ '、'.join(no_spec_lines) }规格口径待更新，当前仅展示数据区间，不做超规判定。")
    lines.append("1、原料bom：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append("2、配方：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append(
        "3、制程："
        + f"烧结压实 {_metric_parts('制程', '烧结压实', 3)}。"
        + f"粉碎压实 {_metric_parts('制程', '粉碎压实', 3)}。"
    )

    process_points = max(
        _max_points("制程", "烧结压实趋势"),
        _max_points("制程", "粉碎压实趋势"),
    )
    process_window = max(
        _max_window("制程", "烧结压实趋势"),
        _max_window("制程", "粉碎压实趋势"),
    )
    if process_points:
        label = f"近{process_points}日数据均值"
        if process_window and process_points < process_window:
            label = f"{label}（不足{process_window}日数据）"
        lines.append(f"  制程趋势（{label}）：")
        lines.append(
            "    烧结压实 "
            + _trend_parts("制程", "烧结压实趋势", 3, "", _TREND_SIGNIFICANT_REL, None)
            + "。"
        )
        lines.append(
            "    粉碎压实 "
            + _trend_parts("制程", "粉碎压实趋势", 3, "", _TREND_SIGNIFICANT_REL, None)
            + "。"
        )

    lines.append("4、成品：")
    lines.append(f"  ①成品压实：{_metric_parts('成品', '成品压实', 3)}。")
    lines.append(f"  ②0.1C充电：{_metric_parts('成品', '0.1C充电', 1)}。")
    lines.append(f"0.1C放电：{_metric_parts('成品', '0.1C放电', 1)}。")
    lines.append(f"首效：{_metric_parts('成品', '首效', 2)}。")
    lines.append(f"平台效率：{_metric_parts('成品', '平台效率', 1)}。")
    lines.append(f"  ③残碱(Li+)：{_metric_parts('成品', '残碱(Li+)', 0)}。")
    lines.append(f"  ④碳含量：{_metric_parts('成品', '碳含量', 2)}。")
    lines.append(f"  ⑤粉阻(粉末电阻)：{_metric_parts('成品', '粉阻(粉末电阻)', 1)}。")
    lines.append(f"  ⑥比表(麦克比表)：{_metric_parts('成品', '比表(麦克比表)', 1)}。")

    product_points = max(
        _max_points("成品", "成品压实趋势"),
        _max_points("成品", "0.1C充电趋势"),
        _max_points("成品", "0.1C放电趋势"),
        _max_points("成品", "首效趋势"),
        _max_points("成品", "平台效率趋势"),
        _max_points("成品", "残碱(Li+)趋势"),
        _max_points("成品", "碳含量趋势"),
        _max_points("成品", "粉阻(粉末电阻)趋势"),
        _max_points("成品", "比表(麦克比表)趋势"),
    )
    product_window = max(
        _max_window("成品", "成品压实趋势"),
        _max_window("成品", "0.1C充电趋势"),
        _max_window("成品", "0.1C放电趋势"),
        _max_window("成品", "首效趋势"),
        _max_window("成品", "平台效率趋势"),
        _max_window("成品", "残碱(Li+)趋势"),
        _max_window("成品", "碳含量趋势"),
        _max_window("成品", "粉阻(粉末电阻)趋势"),
        _max_window("成品", "比表(麦克比表)趋势"),
    )
    if product_points:
        label = f"近{product_points}日数据均值"
        if product_window and product_points < product_window:
            label = f"{label}（不足{product_window}日数据）"
        lines.append(f"  成品趋势（{label}）：")
        lines.append("    成品压实 " + _trend_parts("成品", "成品压实趋势", 3, "", _TREND_SIGNIFICANT_REL, None) + "。")
        lines.append("    0.1C充电 " + _trend_parts("成品", "0.1C充电趋势", 1, "", _TREND_SIGNIFICANT_REL, None) + "。")
        lines.append("    0.1C放电 " + _trend_parts("成品", "0.1C放电趋势", 1, "", _TREND_SIGNIFICANT_REL, None) + "。")
        lines.append("    首效 " + _trend_parts("成品", "首效趋势", 2, "", _TREND_SIGNIFICANT_REL, None) + "。")
        lines.append("    平台效率 " + _trend_parts("成品", "平台效率趋势", 1, "", _TREND_SIGNIFICANT_REL, None) + "。")
        lines.append(
            "    残碱(Li+) "
            + _trend_parts("成品", "残碱(Li+)趋势", 0, "ppm", _TREND_SIGNIFICANT_REL_LI, _TREND_SIGNIFICANT_ABS_LI)
            + "。"
        )
        lines.append("    碳含量 " + _trend_parts("成品", "碳含量趋势", 2, "", _TREND_SIGNIFICANT_REL, None) + "。")
        lines.append(
            "    粉阻(粉末电阻) "
            + _trend_parts("成品", "粉阻(粉末电阻)趋势", 1, "", _TREND_SIGNIFICANT_REL, None)
            + "。"
        )
        lines.append(
            "    比表(麦克比表) "
            + _trend_parts("成品", "比表(麦克比表)趋势", 1, "", _TREND_SIGNIFICANT_REL, None)
            + "。"
        )

    lines.append("5、下一步计划：本次Excel未包含（可从模板/手工输入/第二张表接入）")
    lines.append("6、工艺验证：本次Excel未包含（待接入第二张表/Sheet）")
    return "\n".join(lines)


_LEADER_ABNORMAL_KEYS: list[tuple[str, str]] = [
    ("制程", "烧结压实"),
    ("制程", "粉碎压实"),
    ("成品", "成品压实"),
    ("成品", "0.1C充电"),
    ("成品", "0.1C放电"),
    ("成品", "首效"),
    ("成品", "平台效率"),
    ("成品", "残碱(Li+)"),
    ("成品", "碳含量"),
    ("成品", "粉阻(粉末电阻)"),
    ("成品", "比表(麦克比表)"),
]

_LEADER_KEY_METRICS: list[tuple[str, str, int, str]] = [
    ("制程", "烧结压实", 3, ""),
    ("制程", "粉碎压实", 3, ""),
    ("成品", "成品压实", 3, ""),
    ("成品", "0.1C放电", 1, ""),
    ("成品", "首效", 2, ""),
    ("成品", "残碱(Li+)", 0, "ppm"),
]


def _stat_is_abnormal(stat: Any) -> bool:
    if not isinstance(stat, dict):
        return False
    judge = stat.get("判异")
    return isinstance(judge, dict) and bool(judge.get("异常") is True)


def _stat_source_batch_text(stat: Any) -> str:
    if not isinstance(stat, dict):
        return "投料批次未知"
    source_date = stat.get("来源日期")
    batch_summary = stat.get("来源批次摘要")
    if isinstance(source_date, dt.date):
        date_str = source_date.strftime("%Y.%m.%d")
        if isinstance(batch_summary, str) and batch_summary:
            return f"{date_str}投料批次{batch_summary}"
        return f"{date_str}投料批次未知"
    if isinstance(batch_summary, str) and batch_summary:
        return f"投料批次{batch_summary}"
    return "投料批次未知"


def _leader_key_metric_line(line_label: str, metrics: dict[str, Any]) -> str:
    parts: list[str] = []
    for section, key, decimals, unit in _LEADER_KEY_METRICS:
        st = metrics.get(section, {}).get(key)
        if not isinstance(st, dict):
            continue
        value = _fmt_range(st, decimals)
        if unit:
            value = _with_unit(value, unit)
        parts.append(f"{key}{value}")
    if not parts:
        return f"{line_label} 无有效数据"
    return f"{line_label} " + "，".join(parts)


def build_wecom_text_leader(line_reports: list[dict[str, Any]]) -> str:
    valid_reports = [
        r
        for r in line_reports
        if isinstance(r, dict) and isinstance(r.get("metrics"), dict) and isinstance(r.get("line_label"), str)
    ]
    if not valid_reports:
        return ""

    report_dates = [r.get("report_date") for r in valid_reports if isinstance(r.get("report_date"), dt.date)]
    date_set = sorted({d for d in report_dates if isinstance(d, dt.date)})
    if len(date_set) == 1:
        date_str = date_set[0].strftime("%Y.%m.%d")
    else:
        date_str = "/".join(d.strftime("%Y.%m.%d") for d in date_set) if date_set else dt.date.today().strftime("%Y.%m.%d")

    abnormal_items: list[str] = []
    key_metric_lines: list[str] = []
    for r in valid_reports:
        line_label = str(r["line_label"])
        metrics = r["metrics"]
        key_metric_lines.append(_leader_key_metric_line(line_label, metrics))
        for section, key in _LEADER_ABNORMAL_KEYS:
            st = metrics.get(section, {}).get(key)
            if _stat_is_abnormal(st):
                abnormal_items.append(f"{line_label} {key}（{_stat_source_batch_text(st)}）")

    abnormal_items = list(dict.fromkeys(abnormal_items))
    if abnormal_items:
        conclusion = f"关注（{len(abnormal_items)}项异常）"
        abnormal_text = "；".join(abnormal_items)
    else:
        conclusion = "正常"
        abnormal_text = "无"

    key_metric_text = "；".join([x for x in key_metric_lines if x]) if key_metric_lines else "无"
    lines = [
        f"1、今日结论（{date_str}）：{conclusion}",
        f"2、异常项清单：{abnormal_text}",
        f"3、关键指标区间：{key_metric_text}",
    ]
    return "\n".join(lines)


def main() -> int:
    _ensure_utf8_stdout()

    parser = argparse.ArgumentParser(description="从Excel生成企业微信日报文本（制程/成品）。")
    parser.add_argument("--excel", default=None, help="Excel路径；默认匹配当前目录下 *.xlsx")
    parser.add_argument("--sheet", default=None, help="指定要生成日报的工作表名（如 S18-B线、S006-B线）")
    parser.add_argument("--model", default=None, choices=["S18", "S006"], help="按产品型号筛选工作表")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作簿中的候选线别并退出")
    parser.add_argument("--date", default=None, help="日期：YYYY-MM-DD；默认取表内最新投料日期")
    parser.add_argument("--lookback-days", type=int, default=7, help="指标取数向前回溯天数（避免当日某列为空显示“未录入”）")
    parser.add_argument("--trend-days", type=int, default=7, help="趋势窗口（近N次有数）；设为0可关闭")
    parser.add_argument("--disable-trend", action="store_true", help="关闭趋势窗口分析")
    parser.add_argument("--out", default=None, help="输出到文件（UTF-8）；不填则打印到控制台")
    args = parser.parse_args()

    path = args.excel
    if not path:
        matches = sorted(glob.glob("*.xlsx"))
        if not matches:
            print("未找到Excel：请使用 --excel 指定路径", file=sys.stderr)
            return 2
        path = matches[0]

    all_sheets = list_workbook_sheets(path)
    line_sheets = list_line_sheets(path)

    if args.list_sheets:
        if not line_sheets:
            print("未识别到可用线别工作表。")
            return 0
        print("候选线别工作表：")
        for s in line_sheets:
            model = detect_model_from_sheet(s)
            print(f"- {s} (model={model})")
        return 0

    if not line_sheets:
        raise ValueError(f"未识别到可用线别工作表：{path}")

    target_sheet: Optional[str] = None
    if args.sheet:
        if args.sheet not in all_sheets:
            raise ValueError(f"工作表不存在：{args.sheet}，可选：{', '.join(all_sheets)}")
        if args.model:
            sheet_model = detect_model_from_sheet(args.sheet)
            if sheet_model != "UNKNOWN" and sheet_model != args.model:
                raise ValueError(
                    f"工作表 {args.sheet} 的型号为 {sheet_model}，与 --model {args.model} 不一致"
                )
        target_sheet = args.sheet
    else:
        candidates = line_sheets
        if args.model:
            candidates = [s for s in candidates if detect_model_from_sheet(s) == args.model]
            if not candidates:
                available = ", ".join([f"{s}({detect_model_from_sheet(s)})" for s in line_sheets])
                raise ValueError(f"未找到型号为 {args.model} 的候选线别；现有：{available}")
        for s in candidates:
            try:
                _ = load_sheet_table(path, s)
                target_sheet = s
                break
            except Exception:
                continue
        if target_sheet is None:
            raise ValueError(f"候选线别均无法读取：{', '.join(candidates)}")

    assert target_sheet is not None

    df = load_sheet_table(path, target_sheet)

    trend_days = 0 if args.disable_trend else max(0, int(args.trend_days))
    detected_model = detect_model_from_sheet(target_sheet)
    model = detected_model if detected_model != "UNKNOWN" else (args.model or detected_model)
    profile = PRODUCT_SPEC_PROFILES.get(model, get_profile_for_sheet(target_sheet))

    report_date = pick_date_from_dfs([df], args.date)
    metrics = extract_metrics(df, report_date, args.lookback_days, trend_days, enable_spec=profile.enable_spec)
    detail_text = build_wecom_text_single(
        report_date=report_date,
        metrics=metrics,
        line_label=target_sheet,
        model=model,
        enable_spec=profile.enable_spec,
    )
    line_report = {
        "line_label": target_sheet,
        "report_date": report_date,
        "model": model,
        "enable_spec": profile.enable_spec,
        "metrics": metrics,
    }
    leader_text = build_wecom_text_leader([line_report])
    text = leader_text if not detail_text else (leader_text + "\n\n【工程版】\n" + detail_text)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
            f.write("\n")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
