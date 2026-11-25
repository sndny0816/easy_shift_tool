# -*- mode: python ; coding: utf-8 -*-
"""対象OSで実行する、GUI付きディレクトリ形式のビルド設定。"""
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all


project_dir = Path(SPECPATH)
# uvなどの再配置されたPythonでも、Analysisの分離プロセスがTcl/Tkを
# 検出できるようにする。配布物の起動時はPyInstallerのruntime hookを使う。
sys.path.insert(0, str(project_dir))
from sinta_shift.runtime import prepare_tk_runtime
from sinta_shift import __version__

prepare_tk_runtime()

datas = []
binaries = []
hiddenimports = []
for package in ("ortools", "pandas", "openpyxl"):
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas.extend(package_datas)
    binaries.extend(package_binaries)
    hiddenimports.extend(package_hiddenimports)

a = Analysis(
    [str(project_dir / "launch.py")],
    pathex=[str(project_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "IPython", "jupyter", "matplotlib"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="SintaShift",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="SintaShift",
)
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="Sinta Shift.app",
        icon=None,
        bundle_identifier="jp.sinta.shift",
        info_plist={
            "CFBundleDisplayName": "Sinta 勤務表作成",
            "CFBundleShortVersionString": __version__,
            "NSHighResolutionCapable": True,
        },
    )
