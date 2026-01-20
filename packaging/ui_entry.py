#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import sys
from pathlib import Path
import socket


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

    address = os.environ.get("WE_COM_REPORT_UI_BIND", "0.0.0.0")

    # 避免 8501 端口被占用导致双击“闪退”：在 8501~8600 里选一个可用端口
    port = None
    for candidate in range(8501, 8601):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind((address, candidate))
            except OSError:
                continue
            port = candidate
            break
    if port is None:
        print("未找到可用端口（8501~8600），请关闭占用端口的程序后重试。", file=sys.stderr)
        return 3

    # 等价于：python -m streamlit run wecom_report_ui.py
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        ui_script,
        "--global.developmentMode=false",
        f"--server.port={port}",
        f"--server.address={address}",
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
    ]
    return int(stcli.main())


if __name__ == "__main__":
    raise SystemExit(main())
