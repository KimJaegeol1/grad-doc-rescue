# -*- coding: utf-8 -*-
"""
ui.widgets — 화면 부품 셋
================================================================
DropZone   끌어다 놓는 자리
PathRow    라벨 + 경로칸 + [선택]
LogView    진행 기록

검수도구 `gui/widgets.py` 에서 이 셋만 오려 왔다 — 한 글자도 안 고쳤다.
안 가져온 것: OptionBar · MatchTable · StepPanel.
  단계가 하나뿐이라 규격을 위젯으로 펴 줄 일이 없고(OptionBar·StepPanel),
  사람에게 짝을 고쳐 달라고 묻지 않기로 했다(MatchTable).
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, ttk

from . import theme as T

try:
    from tkinterdnd2 import DND_FILES
    DND = True
except Exception:                                        # noqa: BLE001
    DND_FILES = None
    DND = False

NO_FOLDER = "(폴더 없음)"      # 재지정 드롭다운의 '매칭 안 함' 항목

def _clean_drop(data: str) -> list[str]:
    out, buf, brace = [], "", False
    for ch in data:
        if ch == "{":
            brace = True
        elif ch == "}":
            brace = False
            out.append(buf); buf = ""
        elif ch == " " and not brace:
            if buf:
                out.append(buf); buf = ""
        else:
            buf += ch
    if buf:
        out.append(buf)
    return [p for p in out if p]


# ══════════════════════════════════════════════════════════════
# 드롭존


# ══════════════════════════════════════════════════════════════
# 드롭존
# ══════════════════════════════════════════════════════════════
class DropZone(tk.Label):
    def __init__(self, parent, title, on_path, pad=(0, 0)):
        super().__init__(parent, height=3, bd=2, relief="ridge", cursor="hand2",
                         font=T.FONT_B, justify="center")
        self.title = title
        self.on_path = on_path
        self.pack(side="left", fill="both", expand=True, padx=pad)
        self.bind("<Button-1>", lambda e: self.on_path(None))
        self.reset()
        if DND:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._drop)
            except Exception:                            # noqa: BLE001
                pass

    def _drop(self, ev):
        paths = _clean_drop(ev.data)
        if paths:
            self.on_path(paths[0])

    def reset(self):
        hint = "여기에 끌어다 놓기" if DND else "클릭해서 고르기"
        self.config(text=f"{self.title}\n{hint}", bg=T.ACCENT, fg=T.ACC_FG)

    def ok(self, name):
        self.config(text=f"{self.title}\n{name}", bg=T.OK_BG, fg="#1B7F3B")

    def bad(self, why):
        self.config(text=f"{self.title}\n{why}", bg=T.BAD_BG, fg="#C62828")


# ══════════════════════════════════════════════════════════════
# 경로 한 줄
# ══════════════════════════════════════════════════════════════


class PathRow(tk.Frame):
    def __init__(self, parent, spec, on_change=None):
        super().__init__(parent)
        self.pack(fill="x", pady=2)
        self.spec = spec
        self.var = tk.StringVar(value=spec.get("value", ""))
        if on_change:
            self.var.trace_add("write", lambda *_: on_change())

        tk.Label(self, text=spec["label"], font=T.FONT, width=12,
                 anchor="w").pack(side="left")
        self.entry = tk.Entry(self, textvariable=self.var, font=T.FONT)
        self.entry.pack(side="left", fill="x", expand=True)
        kind = {"file": "파일", "folder": "폴더", "save": "저장"}[spec["kind"]]
        tk.Button(self, text=f"{kind} 선택", font=T.FONT,
                  command=self.pick).pack(side="left", padx=(6, 0))

        if DND:
            try:
                self.entry.drop_target_register(DND_FILES)
                self.entry.dnd_bind("<<Drop>>", self._drop)
            except Exception:                            # noqa: BLE001
                pass

    def _drop(self, ev):
        paths = _clean_drop(ev.data)
        if paths:
            self.var.set(paths[0])

    def pick(self):
        k, ft = self.spec["kind"], self.spec.get("filter") or [("모든 파일", "*.*")]
        cur = self.var.get().strip()
        init = os.path.dirname(cur) if cur else None
        if k == "file":
            p = filedialog.askopenfilename(title=self.spec["label"], filetypes=ft,
                                           initialdir=init)
        elif k == "folder":
            p = filedialog.askdirectory(title=self.spec["label"], initialdir=init)
        else:
            p = filedialog.asksaveasfilename(title=self.spec["label"], filetypes=ft,
                                             initialdir=init,
                                             defaultextension=ft[0][1].lstrip("*"))
        if p:
            self.var.set(p)

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, v: str):
        self.var.set(v)


# ══════════════════════════════════════════════════════════════
# 옵션 줄
# ══════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════
# 진행 기록
# ══════════════════════════════════════════════════════════════
class LogView(tk.LabelFrame):
    def __init__(self, parent):
        super().__init__(parent, text=" 진행 기록 ", font=T.FONT_B, padx=6, pady=4)
        sb = tk.Scrollbar(self); sb.pack(side="right", fill="y")
        self.txt = tk.Text(self, font=T.FONT, wrap="word", bg=T.PANEL, relief="flat",
                           yscrollcommand=sb.set, state="disabled")
        self.txt.pack(fill="both", expand=True)
        sb.config(command=self.txt.yview)
        for tag, (col, bold, _) in T.LEVELS.items():
            self.txt.tag_config(tag, foreground=col,
                                font=(T.FAMILY, 9, "bold") if bold else T.FONT)

    def write(self, msg, level="info"):
        self.txt.config(state="normal")
        self.txt.insert("end", T.prefix(level) + str(msg) + "\n", level)
        self.txt.see("end")
        self.txt.config(state="disabled")

    def clear(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")


# ══════════════════════════════════════════════════════════════
# 단계판 — 입력 + (선택)매칭 + 실행
# ══════════════════════════════════════════════════════════════
