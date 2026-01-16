#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import sys
from pathlib import Path


def _resource_path(rel: str) -> str:
    # PyInstaller onefile/onedir 兼容：打包后资源会解到 sys._MEIPASS
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return str(Path(base) / rel)
    return str(Path(__file__).resolve().parent.parent / rel)


def main() -> int:
    # 让UI脚本认为自己是被 Streamlit 启动的，避免自动“二次拉起”
    os.environ["WE_COM_REPORT_UI_LAUNCHER"] = "1"

    ui_script = _resource_path("wecom_report_ui.py")
    if not Path(ui_script).exists():
        print(f"找不到UI脚本：{ui_script}", file=sys.stderr)
        return 2

    # 等价于：python -m streamlit run wecom_report_ui.py
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        ui_script,
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
    ]
    return int(stcli.main())


if __name__ == "__main__":
    raise SystemExit(main())
