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
    return re.sub(r"\s+", "", name).strip()


_BATCH_RE = re.compile(r"^(?:D?A|D?B)\d{4}-\d{3}$")


def _col_contains(col: str, token: str) -> bool:
    norm_col = _sanitize_col(col).replace("＋", "+")
    norm_token = _sanitize_col(token).replace("＋", "+")
    return (token in col) or (norm_token in norm_col)


def _find_row_contains(df: pd.DataFrame, keyword: str, max_rows: int = 80) -> Optional[int]:
    for i in range(min(len(df), max_rows)):
        row = df.iloc[i].astype(str)
        if row.str.contains(keyword, na=False).any():
            return i
    return None


def _find_data_start(df: pd.DataFrame, max_rows: int = 120) -> Optional[int]:
    for i in range(min(len(df), max_rows)):
        row = df.iloc[i].astype(str)
        if row.str.match(_BATCH_RE, na=False).any():
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
    data_start = _find_data_start(raw)
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
            .replace("%", "")
            .strip()
        )
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


_NUM = r"[-+]?\d+(?:\.\d+)?"


def parse_spec_from_colname(col: str) -> Optional[Spec]:
    # 从列名里解析常见内控口径：≥x、≤x、a~b、a-b、x±y
    # 注意：这里只用于“是否超内控”的粗判定；真正口径建议从“内控/规格表”统一维护。
    text = _sanitize_col(col).replace("～", "~").replace("—", "-").replace("−", "-")

    m = re.search(rf"≥\s*({_NUM})", text)
    if m:
        return Spec(lower=float(m.group(1)), upper=None, text=f"≥{m.group(1)}")
    m = re.search(rf"≤\s*({_NUM})", text)
    if m:
        return Spec(lower=None, upper=float(m.group(1)), text=f"≤{m.group(1)}")
    m = re.search(rf"({_NUM})\s*[~-]\s*({_NUM})", text)
    if m:
        a = float(m.group(1))
        b = float(m.group(2))
        lo, hi = (a, b) if a <= b else (b, a)
        return Spec(lower=lo, upper=hi, text=f"{lo}-{hi}")
    m = re.search(rf"({_NUM})\s*±\s*({_NUM})", text)
    if m:
        x = float(m.group(1))
        y = float(m.group(2))
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
        "扣电_0.1C充",
        "扣电_0.1C放",
        "扣电_0.1C首效",
        "烧结压实",
        "粉碎压实",
    ]
    for pat in patterns:
        if "+" in pat or "＋" in pat:
            cols = [c for c in df.columns if _col_contains(c, pat)]
        else:
            cols = [c for c in df.columns if (pat in c) or c.startswith(pat)]
        values = _flatten_numeric_values(day, cols)
        if not values.empty:
            return True
    return False


def pick_date(df_a: pd.DataFrame, df_b: pd.DataFrame, date_arg: Optional[str]) -> dt.date:
    if date_arg:
        return pd.to_datetime(date_arg).date()

    dates: set[dt.date] = set()
    for df in (df_a, df_b):
        if "投料日期" in df.columns:
            dates |= set([d for d in df["投料日期"].dropna().tolist() if isinstance(d, dt.date)])
    if not dates:
        raise ValueError("无法从Excel中推断最新日期，请使用 --date 指定。")

    # 优先选择“最新且有关键指标数据”的日期（避免最新投料日期只有制程前段录入，成品/扣电还未出数）
    for d in sorted(dates, reverse=True):
        if _has_any_metric_data(df_a, d) or _has_any_metric_data(df_b, d):
            return d

    # 兜底：取最大日期
    return max(dates)


def _stat_for_cols(day: pd.DataFrame, cols: list[str], scale: float = 1.0) -> dict[str, Any]:
    values = _flatten_numeric_values(day, cols)
    if scale != 1.0:
        values = values * scale
    st = _range_stat(values)
    if st is None:
        return {"有数据": False}
    spec = parse_spec_from_colname(cols[0]) if cols else None
    if spec is not None and scale != 1.0:
        spec = Spec(
            lower=None if spec.lower is None else spec.lower * scale,
            upper=None if spec.upper is None else spec.upper * scale,
            text=spec.text,
        )
    judge = judge_out_of_spec(values, spec)
    return {
        "有数据": True,
        "min": st.min,
        "max": st.max,
        "mean": st.mean,
        "n": st.n,
        "判异": judge,
    }


_BATCH_PARSE_RE = re.compile(r"^(?P<prefix>[A-Za-z]+)(?P<yymm>\d{4})-(?P<seq>\d{3})$")


def _batch_sort_key(batch: str) -> tuple[Any, ...]:
    b = (batch or "").strip()
    m = _BATCH_PARSE_RE.match(b)
    if not m:
        return (b, 0, 0)
    return (m.group("prefix"), int(m.group("yymm")), int(m.group("seq")))


def _summarize_batches(batches: list[str]) -> str:
    batches = [b.strip() for b in batches if isinstance(b, str) and b.strip() and b.strip().lower() != "nan"]
    if not batches:
        return ""
    batches = sorted(list(dict.fromkeys(batches)), key=_batch_sort_key)
    if len(batches) == 1:
        return batches[0]
    return f"{batches[0]}~{batches[-1]}"


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


def _latest_stat_within_days(
    df: pd.DataFrame, cols: list[str], report_date: dt.date, lookback_days: int, scale: float = 1.0
) -> dict[str, Any]:
    if "投料日期" not in df.columns:
        return {"有数据": False}

    for delta in range(0, max(0, lookback_days) + 1):
        d = report_date - dt.timedelta(days=delta)
        day = df[df["投料日期"] == d]
        st = _stat_for_cols(day, cols, scale=scale)
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


def _trend_for_cols(
    df: pd.DataFrame,
    cols: list[str],
    report_date: dt.date,
    trend_points: int,
    max_lookback_days: int,
    scale: float = 1.0,
    anomaly_ratio: float = _TREND_ANOMALY_RATIO,
    anomaly_abs: Optional[float] = None,
) -> dict[str, Any]:
    if "投料日期" not in df.columns or not cols or trend_points <= 0:
        return {"有数据": False, "窗口": trend_points, "点数": 0}

    points: list[tuple[dt.date, dict[str, Any]]] = []
    for delta in range(0, max_lookback_days + 1):
        d = report_date - dt.timedelta(days=delta)
        day = df[df["投料日期"] == d]
        st = _stat_for_cols(day, cols, scale=scale)
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


def extract_metrics(df: pd.DataFrame, report_date: dt.date, lookback_days: int, trend_days: int = 7) -> dict[str, Any]:
    if "投料日期" in df.columns:
        day = df[df["投料日期"] == report_date].copy()
    else:
        day = df.head(0).copy()

    # 注意：很多表会出现“均值未填，但B1-1/B1-2/B1-3已填”的情况，因此这里不只取“均值”
    sinter_cols = [c for c in df.columns if "烧结压实" in c]
    crush_cols = [c for c in df.columns if "粉碎压实" in c]

    prod_density_cols = [c for c in df.columns if "成品压实" in c]
    powder_res_cols = [c for c in df.columns if "粉末电阻" in c]
    li_cols = [c for c in df.columns if _col_contains(c, "Li+含量")]
    carbon_cols = [c for c in df.columns if c.startswith("碳含量")]
    bet_cols = [c for c in df.columns if "麦克比表" in c]

    charge_cols = [c for c in df.columns if c.startswith("扣电_0.1C充")]
    discharge_cols = [c for c in df.columns if c.startswith("扣电_0.1C放")]
    eff_cols = [c for c in df.columns if c.startswith("扣电_0.1C首效")]
    plat_cols = [c for c in df.columns if c.startswith("扣电_3.2V平台效率")]

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
            "烧结压实": _latest_stat_within_days(df, sinter_cols, report_date, lookback_days),
            "粉碎压实": _latest_stat_within_days(df, crush_cols, report_date, lookback_days),
            "烧结压实趋势": sinter_trend,
            "粉碎压实趋势": crush_trend,
        },
        "成品": {
            "成品压实": _latest_stat_within_days(df, prod_density_cols, report_date, lookback_days),
            "0.1C充电": _latest_stat_within_days(df, charge_cols, report_date, lookback_days),
            "0.1C放电": _latest_stat_within_days(df, discharge_cols, report_date, lookback_days),
            "首效": _latest_stat_within_days(df, eff_cols, report_date, lookback_days),
            "平台效率": _latest_stat_within_days(df, plat_cols, report_date, lookback_days),
            "残碱(Li+)": _latest_stat_within_days(df, li_cols, report_date, lookback_days, scale=10000),
            "碳含量": _latest_stat_within_days(df, carbon_cols, report_date, lookback_days),
            "粉阻(粉末电阻)": _latest_stat_within_days(df, powder_res_cols, report_date, lookback_days),
            "比表(麦克比表)": _latest_stat_within_days(df, bet_cols, report_date, lookback_days),
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


def _fmt_status(stat: dict[str, Any]) -> str:
    judge = stat.get("判异") if isinstance(stat, dict) else None
    if not isinstance(judge, dict):
        return ""
    abnormal = judge.get("异常")
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


def _fmt_metric(stat: dict[str, Any], decimals: int) -> str:
    return _fmt_range(stat, decimals) + _fmt_status(stat) + _fmt_source_date(stat)


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
    a_sinter = _fmt_metric(a_sinter_stat, 3)
    b_sinter = _fmt_metric(b_sinter_stat, 3)
    ab_sinter = _fmt_range(_merge_stats(a_sinter_stat, b_sinter_stat), 3)

    a_crush_stat = a["制程"]["粉碎压实"]
    b_crush_stat = b["制程"]["粉碎压实"]
    a_crush = _fmt_metric(a_crush_stat, 3)
    b_crush = _fmt_metric(b_crush_stat, 3)
    ab_crush = _fmt_range(_merge_stats(a_crush_stat, b_crush_stat), 3)
    a_sinter_trend = _fmt_trend(a["制程"]["烧结压实趋势"], 3)
    b_sinter_trend = _fmt_trend(b["制程"]["烧结压实趋势"], 3)
    a_crush_trend = _fmt_trend(a["制程"]["粉碎压实趋势"], 3)
    b_crush_trend = _fmt_trend(b["制程"]["粉碎压实趋势"], 3)

    a_prod_density_stat = a["成品"]["成品压实"]
    b_prod_density_stat = b["成品"]["成品压实"]
    ab_prod_density = _fmt_range(_merge_stats(a_prod_density_stat, b_prod_density_stat), 3)

    a_charge = _fmt_metric(a["成品"]["0.1C充电"], 1)
    a_discharge = _fmt_metric(a["成品"]["0.1C放电"], 1)
    a_eff = _fmt_metric(a["成品"]["首效"], 2)
    a_plat = _fmt_metric(a["成品"]["平台效率"], 1)

    b_charge = _fmt_metric(b["成品"]["0.1C充电"], 1)
    b_discharge = _fmt_metric(b["成品"]["0.1C放电"], 1)
    b_eff = _fmt_metric(b["成品"]["首效"], 2)
    b_plat = _fmt_metric(b["成品"]["平台效率"], 1)

    ab_alkali_stat = _merge_stats(a["成品"]["残碱(Li+)"], b["成品"]["残碱(Li+)"])
    ab_alkali = _fmt_range(ab_alkali_stat, 0)
    ab_alkali_status = _fmt_li_status(ab_alkali_stat)
    a_carbon = _fmt_metric(a["成品"]["碳含量"], 2)
    b_carbon = _fmt_metric(b["成品"]["碳含量"], 2)
    a_powder_r = _fmt_metric(a["成品"]["粉阻(粉末电阻)"], 1)
    b_powder_r = _fmt_metric(b["成品"]["粉阻(粉末电阻)"], 1)
    a_bet = _fmt_metric(a["成品"]["比表(麦克比表)"], 1)
    b_bet = _fmt_metric(b["成品"]["比表(麦克比表)"], 1)
    a_prod_trend = _fmt_trend(a["成品"]["成品压实趋势"], 3)
    b_prod_trend = _fmt_trend(b["成品"]["成品压实趋势"], 3)
    a_charge_trend = _fmt_trend(a["成品"]["0.1C充电趋势"], 1)
    b_charge_trend = _fmt_trend(b["成品"]["0.1C充电趋势"], 1)
    a_discharge_trend = _fmt_trend(a["成品"]["0.1C放电趋势"], 1)
    b_discharge_trend = _fmt_trend(b["成品"]["0.1C放电趋势"], 1)
    a_eff_trend = _fmt_trend(a["成品"]["首效趋势"], 2)
    b_eff_trend = _fmt_trend(b["成品"]["首效趋势"], 2)
    a_plat_trend = _fmt_trend(a["成品"]["平台效率趋势"], 1)
    b_plat_trend = _fmt_trend(b["成品"]["平台效率趋势"], 1)
    a_li_trend = _fmt_trend(a["成品"]["残碱(Li+)趋势"], 0, unit="ppm")
    b_li_trend = _fmt_trend(b["成品"]["残碱(Li+)趋势"], 0, unit="ppm")
    a_carbon_trend = _fmt_trend(a["成品"]["碳含量趋势"], 2)
    b_carbon_trend = _fmt_trend(b["成品"]["碳含量趋势"], 2)
    a_powder_trend = _fmt_trend(a["成品"]["粉阻(粉末电阻)趋势"], 1)
    b_powder_trend = _fmt_trend(b["成品"]["粉阻(粉末电阻)趋势"], 1)
    a_bet_trend = _fmt_trend(a["成品"]["比表(麦克比表)趋势"], 1)
    b_bet_trend = _fmt_trend(b["成品"]["比表(麦克比表)趋势"], 1)

    # 输出结构按你给的 1~6 段落口径；未提供的数据先保留占位，便于你后续把“第二张表”接进来。
    lines: list[str] = []
    lines.append(f"{date_str}数据表更新：")
    lines.append("1、原料bom：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append("2、配方：本次Excel未包含（待接入第二张表/Sheet）")
    lines.append(
        f"3、制程：烧结压实(AB) {ab_sinter}；A线 {a_sinter}；B线 {b_sinter}。"
        f"粉碎压实(AB) {ab_crush}；A线 {a_crush}；B线 {b_crush}。"
    )
    trend_points = max(a["制程"]["烧结压实趋势"].get("点数", 0), b["制程"]["烧结压实趋势"].get("点数", 0))
    trend_window = max(a["制程"]["烧结压实趋势"].get("窗口", 0), b["制程"]["烧结压实趋势"].get("窗口", 0))
    if trend_points:
        label = f"近{trend_points}个有数日均值"
        if trend_window and trend_points < trend_window:
            label = f"{label}（不足{trend_window}个有数日）"
        lines.append(f"  制程趋势（{label}）：")
        lines.append(f"    烧结压实 A线 {a_sinter_trend}；B线 {b_sinter_trend}。")
        lines.append(f"    粉碎压实 A线 {a_crush_trend}；B线 {b_crush_trend}。")
    lines.append("4、成品：")
    lines.append(f"  ①AB线成品压实：{ab_prod_density}。")
    lines.append(f"  ②A线0.1C充电：{a_charge}；0.1C放电：{a_discharge}；首效：{a_eff}；平台效率：{a_plat}。")
    lines.append(f"  ③B线0.1C充电：{b_charge}；0.1C放电：{b_discharge}；首效：{b_eff}；平台效率：{b_plat}。")
    lines.append(f"  ④AB线残碱(Li+)：{_with_unit(ab_alkali, 'ppm')}{ab_alkali_status}。")
    lines.append(f"  ⑤碳含量：A线 {a_carbon}；B线 {b_carbon}。")
    lines.append(f"  ⑥粉阻(粉末电阻)：A线 {a_powder_r}；B线 {b_powder_r}。")
    lines.append(f"  ⑦比表(麦克比表)：A线 {a_bet}；B线 {b_bet}。")
    trend_points = max(a["成品"]["残碱(Li+)趋势"].get("点数", 0), b["成品"]["残碱(Li+)趋势"].get("点数", 0))
    trend_window = max(a["成品"]["残碱(Li+)趋势"].get("窗口", 0), b["成品"]["残碱(Li+)趋势"].get("窗口", 0))
    if trend_points:
        label = f"近{trend_points}个有数日均值"
        if trend_window and trend_points < trend_window:
            label = f"{label}（不足{trend_window}个有数日）"
        lines.append(f"  成品趋势（{label}）：")
        lines.append(f"    成品压实 A线 {a_prod_trend}；B线 {b_prod_trend}。")
        lines.append(f"    0.1C充电 A线 {a_charge_trend}；B线 {b_charge_trend}。")
        lines.append(f"    0.1C放电 A线 {a_discharge_trend}；B线 {b_discharge_trend}。")
        lines.append(f"    首效 A线 {a_eff_trend}；B线 {b_eff_trend}。")
        lines.append(f"    平台效率 A线 {a_plat_trend}；B线 {b_plat_trend}。")
        lines.append(f"    残碱(Li+) A线 {a_li_trend}；B线 {b_li_trend}。")
        lines.append(f"    碳含量 A线 {a_carbon_trend}；B线 {b_carbon_trend}。")
        lines.append(f"    粉阻(粉末电阻) A线 {a_powder_trend}；B线 {b_powder_trend}。")
        lines.append(f"    比表(麦克比表) A线 {a_bet_trend}；B线 {b_bet_trend}。")
    lines.append("5、下一步计划：本次Excel未包含（可从模板/手工输入/第二张表接入）")
    lines.append("6、工艺验证：本次Excel未包含（待接入第二张表/Sheet）")
    return "\n".join(lines)


def main() -> int:
    _ensure_utf8_stdout()

    parser = argparse.ArgumentParser(description="从Excel生成企业微信日报文本（制程/成品）。")
    parser.add_argument("--excel", default=None, help="Excel路径；默认匹配当前目录下 2026*.xlsx")
    parser.add_argument("--date", default=None, help="日期：YYYY-MM-DD；默认取表内最新投料日期")
    parser.add_argument("--lookback-days", type=int, default=7, help="指标取数向前回溯天数（避免当日某列为空显示“未录入”）")
    parser.add_argument("--trend-days", type=int, default=7, help="趋势窗口（近N次有数）")
    parser.add_argument("--out", default=None, help="输出到文件（UTF-8）；不填则打印到控制台")
    args = parser.parse_args()

    path = args.excel
    if not path:
        matches = glob.glob("2026*.xlsx")
        if not matches:
            print("未找到Excel：请使用 --excel 指定路径", file=sys.stderr)
            return 2
        path = matches[0]

    df_a = load_sheet_table(path, "A线")
    df_b = load_sheet_table(path, "B线")

    report_date = pick_date(df_a, df_b, args.date)
    a = extract_metrics(df_a, report_date, args.lookback_days, args.trend_days)
    b = extract_metrics(df_b, report_date, args.lookback_days, args.trend_days)
    text = build_wecom_text(report_date, a, b)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
            f.write("\n")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
