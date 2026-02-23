#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import datetime as dt
import glob
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Optional

import pandas as pd

import daily_wecom_report as core

DEFAULT_SHEETS = ["S18-A线", "S18-B线", "S006-B线", "S006-A线"]


def _extract_metrics_compat(
    df: pd.DataFrame, report_date: dt.date, lookback_days: int, trend_days: int, enable_spec: bool
) -> dict:
    try:
        return core.extract_metrics(df, report_date, lookback_days, trend_days, enable_spec=enable_spec)
    except TypeError as e:
        if "enable_spec" not in str(e):
            raise
        return core.extract_metrics(df, report_date, lookback_days, trend_days)


def _detect_model(sheet_name: str) -> str:
    if hasattr(core, "detect_model_from_sheet"):
        return str(core.detect_model_from_sheet(sheet_name))
    upper = sheet_name.upper()
    if "S18" in upper:
        return "S18"
    if "S006" in upper:
        return "S006"
    return "UNKNOWN"


def _list_line_sheets_compat(excel_path: str) -> list[str]:
    if hasattr(core, "list_line_sheets"):
        try:
            out = list(core.list_line_sheets(excel_path))
            if out:
                return [str(x) for x in out]
        except Exception:
            pass
    if hasattr(core, "list_workbook_sheets"):
        try:
            all_sheets = [str(x) for x in core.list_workbook_sheets(excel_path)]
            line_like = [s for s in all_sheets if "线" in s]
            if line_like:
                return line_like
            return all_sheets
        except Exception:
            pass
    fallback: list[str] = []
    for key in ("LINE_A_SHEET", "LINE_B_SHEET"):
        val = getattr(core, key, None)
        if isinstance(val, str) and val and val not in fallback:
            fallback.append(val)
    return fallback if fallback else DEFAULT_SHEETS


def _profile_enable_spec_compat(sheet_name: str) -> bool:
    if hasattr(core, "get_profile_for_sheet"):
        try:
            profile = core.get_profile_for_sheet(sheet_name)
            return bool(getattr(profile, "enable_spec", True))
        except Exception:
            pass
    model = _detect_model(sheet_name)
    if model in ("S18", "S006"):
        return True
    return False


def _pick_date_single_compat(df: pd.DataFrame, date_arg: Optional[str]) -> dt.date:
    if hasattr(core, "pick_date_from_dfs"):
        return core.pick_date_from_dfs([df], date_arg)
    return core.pick_date(df, df, date_arg)


def _build_text_single_compat(
    report_date: dt.date, metrics: dict, line_label: str, model: str, enable_spec: bool
) -> str:
    if hasattr(core, "build_wecom_text_single"):
        return core.build_wecom_text_single(
            report_date=report_date,
            metrics=metrics,
            line_label=line_label,
            model=model,
            enable_spec=enable_spec,
        )
    # 兼容旧核心：降级为“单线复制到A/B”拼接
    return core.build_wecom_text(report_date, metrics, metrics)


def _build_text_multi_compat(line_reports: list[dict[str, Any]]) -> str:
    if hasattr(core, "build_wecom_text_multi"):
        return core.build_wecom_text_multi(line_reports)
    blocks: list[str] = []
    for r in line_reports:
        blocks.append(
            _build_text_single_compat(
                report_date=r["report_date"],
                metrics=r["metrics"],
                line_label=r["line_label"],
                model=r.get("model", "UNKNOWN"),
                enable_spec=bool(r.get("enable_spec", True)),
            )
        )
    return "\n\n".join(blocks)


def _build_text_leader_compat(line_reports: list[dict[str, Any]]) -> str:
    if hasattr(core, "build_wecom_text_leader"):
        try:
            return core.build_wecom_text_leader(line_reports)
        except Exception:
            pass

    report_dates = [r.get("report_date") for r in line_reports if isinstance(r.get("report_date"), dt.date)]
    date_set = sorted({d for d in report_dates if isinstance(d, dt.date)})
    if len(date_set) == 1:
        date_str = date_set[0].strftime("%Y.%m.%d")
    else:
        date_str = "/".join(d.strftime("%Y.%m.%d") for d in date_set) if date_set else dt.date.today().strftime("%Y.%m.%d")

    labels = [str(r.get("line_label", "")) for r in line_reports if isinstance(r, dict)]
    line_text = "、".join([x for x in labels if x]) or "未知线别"
    return (
        f"1、今日结论（{date_str}）：已生成\n"
        f"2、异常项清单：请查看工程版\n"
        f"3、关键指标区间：{line_text}（请查看工程版）"
    )


def _find_local_excels() -> list[str]:
    paths = sorted(glob.glob("*.xlsx"))
    # 优先把你这种命名放前面
    preferred = [p for p in paths if "制程" in p and "成品" in p]
    others = [p for p in paths if p not in preferred]
    return preferred + others


def _save_upload_to_temp(upload) -> str:
    suffix = Path(upload.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
        f.write(upload.getbuffer())
        return f.name


def _cleanup_temp_file(path: Optional[str]) -> None:
    if not path:
        return
    try:
        Path(path).unlink()
    except FileNotFoundError:
        pass
    except OSError:
        pass


def _upload_signature(upload) -> tuple[str, int]:
    size = getattr(upload, "size", None)
    if size is None:
        size = len(upload.getbuffer())
    return (upload.name, int(size))


def _date_to_str(d: Optional[pd.Timestamp | object]) -> Optional[str]:
    if d is None:
        return None
    try:
        return pd.to_datetime(d).date().isoformat()
    except Exception:
        return None


def _running_under_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner.script_run_context import get_script_run_ctx

        return get_script_run_ctx() is not None
    except Exception:
        return False


def _launch_with_streamlit() -> int:
    env = dict(os.environ)
    env["WE_COM_REPORT_UI_LAUNCHER"] = "1"
    address = env.get("WE_COM_REPORT_UI_BIND", "0.0.0.0")
    port = env.get("WE_COM_REPORT_UI_PORT")
    cmd = [sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()), f"--server.address={address}"]
    if port:
        cmd.append(f"--server.port={port}")
    return subprocess.call(cmd, env=env)


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="企业微信日报生成器", layout="wide")

    st.title("企业微信日报生成器（制程/成品）")
    st.caption("支持两行/三行拼接表头；按指标回溯最近有数的投料批次并标注。")

    @st.cache_data(show_spinner=False)
    def _load_dates_and_recommended(
        excel_path: str, mtime_ns: int, sheet_names: tuple[str, ...]
    ) -> tuple[list[dt.date], Optional[dt.date]]:
        dfs: list[pd.DataFrame] = []
        dates: set[dt.date] = set()
        for sheet_name in sheet_names:
            try:
                df = core.load_sheet_table(excel_path, sheet_name)
            except Exception:
                continue
            dfs.append(df)
            if "投料日期" in df.columns:
                for d in df["投料日期"].dropna().tolist():
                    if isinstance(d, dt.date):
                        dates.add(d)

        recommended: Optional[dt.date] = None
        if dates and dfs:
            try:
                if hasattr(core, "pick_date_from_dfs"):
                    recommended = core.pick_date_from_dfs(dfs, None)
                elif len(dfs) >= 2:
                    recommended = core.pick_date(dfs[0], dfs[1], None)
                else:
                    recommended = _pick_date_single_compat(dfs[0], None)
            except Exception:
                recommended = sorted(dates, reverse=True)[0]
        return sorted(dates, reverse=True), recommended

    with st.sidebar:
        st.header("输入")

        source_mode = st.radio("Excel来源", ["选择本地文件", "上传文件"], horizontal=False)

        excel_path: Optional[str] = None
        if source_mode == "选择本地文件":
            options = _find_local_excels()
            if not options:
                st.warning("当前目录未发现 .xlsx 文件，请改用“上传文件”。")
            else:
                excel_path = st.selectbox("选择Excel", options, index=0)
        else:
            upload = st.file_uploader("上传Excel（.xlsx）", type=["xlsx"])
            if upload is not None:
                temp_path = st.session_state.get("temp_upload_path")
                temp_sig = st.session_state.get("temp_upload_sig")
                current_sig = _upload_signature(upload)
                if temp_sig != current_sig or not temp_path or not Path(temp_path).exists():
                    excel_path = _save_upload_to_temp(upload)
                    _cleanup_temp_file(temp_path)
                    st.session_state["temp_upload_path"] = excel_path
                    st.session_state["temp_upload_sig"] = current_sig
                    st.success(f"已接收：{upload.name}")
                else:
                    excel_path = temp_path
            else:
                temp_path = st.session_state.get("temp_upload_path")
                if temp_path and Path(temp_path).exists():
                    excel_path = temp_path

        selected_sheets: list[str] = []
        selected_profiles: dict[str, tuple[str, bool]] = {}
        line_options: list[str] = []
        if excel_path is not None and Path(excel_path).exists():
            try:
                line_options = _list_line_sheets_compat(excel_path)
            except Exception as e:
                st.warning(f"读取工作表列表失败：{e}")
            if line_options:
                selected_sheets = st.multiselect("生成线别（可多选）", line_options, default=[line_options[0]])
                if selected_sheets:
                    for s in selected_sheets:
                        model = _detect_model(s)
                        enable_spec = _profile_enable_spec_compat(s)
                        selected_profiles[s] = (model, enable_spec)
                    summary = "；".join(
                        [f"{s}({selected_profiles[s][0]}/规格{'开' if selected_profiles[s][1] else '关'})" for s in selected_sheets]
                    )
                    st.caption(f"已选线别：{summary}")
                else:
                    st.warning("请至少选择一条线别。")
            else:
                st.warning("未识别到可生成日报的线别工作表。")

        date_arg: Optional[str] = None
        available_dates: list[dt.date] = []
        recommended_date: Optional[dt.date] = None
        if excel_path is not None and Path(excel_path).exists() and selected_sheets:
            try:
                mtime_ns = Path(excel_path).stat().st_mtime_ns
                available_dates, recommended_date = _load_dates_and_recommended(excel_path, mtime_ns, tuple(selected_sheets))
            except Exception as e:
                st.warning(f"日期表读取失败：{e}")

        date_mode = st.radio("日期选择", ["日期表（默认今天）", "自动（推荐）", "手动日历"], horizontal=False)
        if date_mode == "日期表（默认今天）":
            if available_dates:
                today = dt.date.today()
                default_date = today if today in available_dates else available_dates[0]
                idx = available_dates.index(default_date) if default_date in available_dates else 0
                selected = st.selectbox(
                    "选择投料日期",
                    available_dates,
                    index=idx,
                    format_func=lambda d: d.strftime("%Y-%m-%d"),
                )
                date_arg = selected.isoformat()
            else:
                st.info("未能从Excel提取到投料日期列表，请改用“手动日历”。")
        elif date_mode == "自动（推荐）":
            if recommended_date is None:
                st.info("未能自动推荐日期，请改用“日期表/手动日历”。")
            else:
                st.caption(f"已自动选择：{recommended_date.strftime('%Y-%m-%d')}")
                date_arg = recommended_date.isoformat()
        else:
            d = st.date_input("选择日期", value=dt.date.today())
            date_arg = _date_to_str(d)

        lookback_days = int(st.slider("回溯天数（用于避免当日某列为空）", min_value=0, max_value=30, value=7, step=1))
        trend_enabled = bool(st.checkbox("启用趋势窗口", value=True))
        if trend_enabled:
            trend_days = int(st.slider("趋势窗口（近N次有数）", min_value=3, max_value=30, value=7, step=1))
        else:
            trend_days = 0

        generate = st.button(
            "生成日报",
            type="primary",
            use_container_width=True,
            disabled=(excel_path is None or not selected_sheets),
        )

    if not generate:
        st.info("在左侧选择/上传 Excel，然后点击“生成日报”。")
        return

    assert excel_path is not None
    assert selected_sheets

    try:
        line_reports: list[dict[str, Any]] = []
        line_dates: dict[str, dt.date] = {}
        line_errors: list[str] = []

        for sheet_name in selected_sheets:
            try:
                model, enable_spec = selected_profiles.get(
                    sheet_name, (_detect_model(sheet_name), _profile_enable_spec_compat(sheet_name))
                )
                df_line = core.load_sheet_table(excel_path, sheet_name)
                report_date = _pick_date_single_compat(df_line, date_arg)
                metrics = _extract_metrics_compat(df_line, report_date, lookback_days, trend_days, enable_spec)
                line_dates[sheet_name] = report_date
                line_reports.append(
                    {
                        "line_label": sheet_name,
                        "report_date": report_date,
                        "model": model,
                        "enable_spec": enable_spec,
                        "metrics": metrics,
                    }
                )
            except Exception as e:
                line_errors.append(f"{sheet_name}: {e}")

        if not line_reports:
            raise ValueError("所选线别全部生成失败：" + "；".join(line_errors))
        if len(line_reports) == 1:
            r = line_reports[0]
            detail_text = _build_text_single_compat(
                report_date=r["report_date"],
                metrics=r["metrics"],
                line_label=r["line_label"],
                model=r.get("model", "UNKNOWN"),
                enable_spec=bool(r.get("enable_spec", True)),
            )
        else:
            detail_text = _build_text_multi_compat(line_reports)
        leader_text = _build_text_leader_compat(line_reports)
        combined_text = leader_text if not detail_text else (leader_text + "\n\n【工程版】\n" + detail_text)

        if line_errors:
            st.warning("部分线别生成失败：" + "；".join(line_errors))

    except Exception as e:
        st.error(f"生成失败：{e}")
        st.exception(e)
        return

    col1, col2 = st.columns([2, 1], gap="large")
    with col1:
        st.subheader("领导版（3行摘要）")
        st.text_area("可直接复制给领导", value=leader_text, height=140)
        with st.expander("工程版（详细段落）", expanded=False):
            st.text_area("工程版内容", value=detail_text, height=420)
        unique_dates = sorted({d.strftime("%Y%m%d") for d in line_dates.values()})
        date_tag = unique_dates[0] if len(unique_dates) == 1 else "multi"
        file_name = f"企业微信日报_{date_tag}_{len(line_reports)}线_分层版.txt"
        st.download_button("下载分层版txt", data=combined_text.encode("utf-8"), file_name=file_name)

    with col2:
        st.subheader("参数")
        trend_value = trend_days if trend_enabled else "已关闭"
        date_value: str | list[str]
        if len(set(line_dates.values())) == 1:
            date_value = next(iter(line_dates.values())).isoformat()
        else:
            date_value = [f"{k}:{v.isoformat()}" for k, v in line_dates.items()]
        st.write(
            {
                "线别": selected_sheets,
                "日期": date_value,
                "回溯天数": lookback_days,
                "趋势窗口": trend_value,
            }
        )
        st.subheader("提示")
        st.write("如果某指标当日未出数，会自动向前回溯到最近一次有数的“投料批次”，并在括号中标注。")


if __name__ == "__main__":
    # 你很可能是用 `python wecom_report_ui.py` 直接运行的，这会出现
    # “missing ScriptRunContext” 并且页面打不开。这里自动帮你切到正确的启动方式：
    # `python -m streamlit run wecom_report_ui.py`
    if os.environ.get("WE_COM_REPORT_UI_LAUNCHER") != "1" and not _running_under_streamlit():
        raise SystemExit(_launch_with_streamlit())
    main()
