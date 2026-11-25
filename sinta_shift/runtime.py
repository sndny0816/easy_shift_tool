"""再配置されたPython配布のTcl/Tkスクリプトを実環境から検出する。"""
import os
from pathlib import Path
import sys


def prepare_tk_runtime() -> None:
    # PyInstallerの専用runtime hookが設定するバンドル内のパスを優先する。
    if getattr(sys, "frozen", False):
        return
    import _tkinter
    for variable, version, filename, prefix in (
        ("TCL_LIBRARY", _tkinter.TCL_VERSION, "init.tcl", "tcl"),
        ("TK_LIBRARY", _tkinter.TK_VERSION, "tk.tcl", "tk"),
    ):
        if os.environ.get(variable):
            continue
        for parent in (Path(sys.base_prefix) / "lib", Path(sys.base_prefix) / "tcl"):
            candidate = parent / f"{prefix}{version}"
            if (candidate / filename).is_file():
                os.environ[variable] = str(candidate)
                break
