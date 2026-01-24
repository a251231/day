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
from typing import Optional

import pandas as pd

import daily_wecom_report as core


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
    def _load_dates_and_recommended(excel_path: str, mtime_ns: int) -> tuple[list[dt.date], Optional[dt.date]]:
        df_a = core.load_sheet_table(excel_path, "A线")
        df_b = core.load_sheet_table(excel_path, "B线")
        dates: set[dt.date] = set()
        for df in (df_a, df_b):
            if "投料日期" in df.columns:
                for d in df["投料日期"].dropna().tolist():
                    if isinstance(d, dt.date):
                        dates.add(d)
        recommended = core.pick_date(df_a, df_b, None) if dates else None
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

        date_arg: Optional[str] = None
        available_dates: list[dt.date] = []
        recommended_date: Optional[dt.date] = None
        if excel_path is not None and Path(excel_path).exists():
            try:
                mtime_ns = Path(excel_path).stat().st_mtime_ns
                available_dates, recommended_date = _load_dates_and_recommended(excel_path, mtime_ns)
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
        trend_days = int(st.slider("趋势窗口（近N次有数）", min_value=3, max_value=30, value=7, step=1))

        generate = st.button("生成日报", type="primary", use_container_width=True, disabled=(excel_path is None))

    if not generate:
        st.info("在左侧选择/上传 Excel，然后点击“生成日报”。")
        return

    assert excel_path is not None

    try:
        df_a = core.load_sheet_table(excel_path, "A线")
        df_b = core.load_sheet_table(excel_path, "B线")

        report_date = core.pick_date(df_a, df_b, date_arg)
        metrics_a = core.extract_metrics(df_a, report_date, lookback_days, trend_days)
        metrics_b = core.extract_metrics(df_b, report_date, lookback_days, trend_days)
        text = core.build_wecom_text(report_date, metrics_a, metrics_b)

    except Exception as e:
        st.error(f"生成失败：{e}")
        st.exception(e)
        return

    col1, col2 = st.columns([2, 1], gap="large")
    with col1:
        st.subheader("企业微信消息文本")
        st.text_area("可直接复制粘贴到企业微信", value=text, height=420)
        st.download_button("下载为txt", data=text.encode("utf-8"), file_name=f"企业微信日报_{report_date.strftime('%Y%m%d')}.txt")

    with col2:
        st.subheader("参数")
        st.write({"日期": report_date.isoformat(), "回溯天数": lookback_days, "趋势窗口": trend_days})
        st.subheader("提示")
        st.write("如果某指标当日未出数，会自动向前回溯到最近一次有数的“投料批次”，并在括号中标注。")


if __name__ == "__main__":
    # 你很可能是用 `python wecom_report_ui.py` 直接运行的，这会出现
    # “missing ScriptRunContext” 并且页面打不开。这里自动帮你切到正确的启动方式：
    # `python -m streamlit run wecom_report_ui.py`
    if os.environ.get("WE_COM_REPORT_UI_LAUNCHER") != "1" and not _running_under_streamlit():
        raise SystemExit(_launch_with_streamlit())
    main()
