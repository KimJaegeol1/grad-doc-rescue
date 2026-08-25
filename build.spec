# -*- mode: python ; coding: utf-8 -*-
"""
build.spec — 보조도구 exe 하나.

검수도구와 다른 점
    · exe 가 하나다 (검수도구는 셋)
    · **검수기준 엑셀을 번들에 안 넣는다.**  exe 옆에 둔다 —
      사람이 열어 고치는 물건이라 프로그램파일/ 안에 숨기면 안 된다
    · secrets/ 도 번들에 안 넣는다
"""
import os

ROOT = os.path.abspath(os.getcwd())
INNER = "프로그램파일"          # paths.py 의 INNER 와 같아야 한다

datas = []
binaries = []

# 동적으로 불러 정적 분석이 놓치는 것 — 여기 적어야 묶인다
hiddenimports = [
    "paths", "preflight",
    "core.spec", "core.book", "core.prompt", "core.ask", "core.calc",
    "core.report", "core.flow",
    "engines.read", "engines.ocr", "engines.write",
    "ui.app", "ui.theme", "ui.runner", "ui.widgets",
    "openpyxl", "openpyxl.cell._writer",
    "openai", "dotenv",
    "lxml", "lxml.etree", "lxml._elementpath",
    "pypdf",
]

from PyInstaller.utils.hooks import collect_all   # noqa: E402


def _있나(m):
    try:
        __import__(m)
        return True
    except Exception:
        return False


for pkg in ("openpyxl", "openai", "dotenv", "lxml", "pypdf", "hwp5",
            "google.cloud.documentai_v1", "tkinterdnd2", "certifi"):
    if not _있나(pkg.split(".")[0]):
        continue
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

excludes = [
    "matplotlib", "numpy", "pandas", "scipy", "PIL",
    "PyQt5", "PySide2", "PySide6", "notebook", "IPython",
    "pytest", "setuptools", "pip", "wheel", "xlwings",
]

a = Analysis(["run.py"], pathex=[ROOT], binaries=binaries, datas=datas,
             hiddenimports=hiddenimports, hookspath=[], runtime_hooks=[],
             excludes=excludes, noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True,
          name="보조도구", console=False, icon=None,
          contents_directory=INNER)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False,
               name="보조도구")
