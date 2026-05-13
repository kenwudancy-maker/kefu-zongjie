# -*- mode: python ; coding: utf-8 -*-
# Windows：在项目目录打开 cmd，执行  build_windows.bat
# 生成目录 dist\客服总结\  内 客服总结.exe（需整文件夹一起拷贝分发）

from PyInstaller.utils.hooks import collect_all

datas = [
    ("app.py", "."),
    (".streamlit/config.toml", ".streamlit"),
]
binaries = []
hiddenimports = []

for pkg in ("streamlit", "altair"):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

hiddenimports += [
    "streamlit.web.cli",
    "streamlit.runtime.scriptrunner.magic_funcs",
    "openpyxl",
    "pandas",
    "httpx",
    "pyarrow",
]

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pytest",
        "pyarrow.tests",
        "pandas.tests",
        "IPython",
        "jupyter",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="客服总结",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
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
    upx_exclude=[],
    name="客服总结",
)
