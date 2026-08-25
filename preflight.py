# -*- coding: utf-8 -*-
"""
preflight — 시작 전 점검
================================================================
필요한 라이브러리가 없으면 파이썬은 영어 오류를 뱉고 창도 안 뜬다.
그 전에 붙잡아, 무엇을 어떻게 깔면 되는지 한국어로 알려 준다.

없어도 되는 것(tkinterdnd2 등)은 막지 않는다 — 알리기만 한다.

검수도구와 다른 점: **openai 가 필수다.**  이 도구는 시트 하나를 통째로
모형에게 맡기는 것이 전부라, 열쇠가 없으면 할 수 있는 일이 없다.
(검수도구는 0·1단계가 모형 없이도 돌았다.)
"""

from __future__ import annotations

import importlib
import sys

# (모듈이름, 무엇에 쓰나, 없으면 못 도나)
NEEDS = [
    ("openpyxl", "엑셀을 읽고 씁니다 — 검수기준·성과관리현황 둘 다", True),
    ("openai", "모형에게 시트를 통째로 맡깁니다 — 이게 없으면 검사가 안 됩니다", True),
    ("dotenv", "secrets/.env 에서 열쇠를 읽습니다", False),
    ("lxml", "hwp·hwpx·docx 를 읽습니다", False),
    ("pypdf", "PDF 를 읽습니다", False),
    ("hwp5", "hwp 를 읽습니다 · pip install pyhwp", False),
    ("google.cloud.documentai_v1", "그림 PDF 를 글자로 바꿉니다 (스캔본이 있을 때만)", False),
    ("tkinterdnd2", "끌어다 놓기", False),
]


def check():
    """반환: (없어서 못 도는 것, 없어도 되는데 없는 것)"""
    hard, soft = [], []
    for mod, why, required in NEEDS:
        try:
            importlib.import_module(mod)
        except Exception:                                # noqa: BLE001
            (hard if required else soft).append((mod, why))
    return hard, soft


def _message(hard, soft) -> str:
    lines = ["프로그램을 시작하려면 아래를 먼저 설치해야 합니다.", ""]
    for mod, why in hard:
        lines.append(f"  · {mod}  —  {why}")
    lines += ["", "명령 프롬프트(cmd)를 열고 이렇게 입력하세요:", ""]
    lines.append("    pip install " + " ".join(m for m, _ in hard))
    if soft:
        lines += ["", "함께 깔아 두면 좋은 것:"]
        for mod, why in soft:
            lines.append(f"  · {mod}  —  {why}")
        lines.append("")
        # 모듈 이름과 설치 이름이 다른 것들을 바로잡는다
        pipname = {"dotenv": "python-dotenv", "hwp5": "pyhwp",
                   "google.cloud.documentai_v1": '"google-cloud-documentai>=2.29.0"'}
        lines.append("    pip install "
                     + " ".join(pipname.get(m, m) for m, _ in soft))
    lines += ["", f"(지금 쓰는 파이썬: {sys.executable})"]
    return "\n".join(lines)


def guard() -> bool:
    """못 도는 게 있으면 창으로 알리고 False. 괜찮으면 True."""
    hard, soft = check()
    if not hard:
        return True
    text = _message(hard, soft)
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("먼저 설치가 필요합니다", text)
        root.destroy()
    except Exception:                                    # noqa: BLE001
        print(text)
    return False
