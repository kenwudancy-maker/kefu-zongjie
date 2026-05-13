# -*- coding: utf-8 -*-
"""Windows 可执行文件入口：启动 Streamlit 运行 app.py。"""
from __future__ import annotations

import os
import sys
import webbrowser
from pathlib import Path
from threading import Timer


def _bundle_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def main() -> None:
    bundle = _bundle_dir()
    app_py = bundle / "app.py"
    if not app_py.is_file():
        print(f"缺少 app.py：{app_py}", file=sys.stderr)
        sys.exit(1)

    # 数据目录、密钥文件写在 exe 同目录，避免 onefile 解压目录只读
    if getattr(sys, "frozen", False):
        os.chdir(str(Path(sys.executable).resolve().parent))

    port = os.environ.get("STREAMLIT_SERVER_PORT", "8720")
    url = f"http://127.0.0.1:{port}"

    def _open_browser() -> None:
        try:
            webbrowser.open(url)
        except OSError:
            pass

    Timer(2.0, _open_browser).start()

    sys.argv = [
        "streamlit",
        "run",
        str(app_py),
        "--server.port",
        port,
        "--server.address",
        "127.0.0.1",
        "--server.headless",
        "true",
        "--browser.gatherUsageStats",
        "false",
        "--global.developmentMode",
        "false",
    ]
    from streamlit.web.cli import main as st_main

    st_main()


if __name__ == "__main__":
    main()
