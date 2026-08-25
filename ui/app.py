# -*- coding: utf-8 -*-
"""
ui.app — 화면 한 장.
================================================================
넣을 것이 셋, 누를 것이 하나다.

    성과관리현황 엑셀   검사 칸(주황)이 들어간 것
    증빙 폴더         폴더 안 폴더까지 알아서 뒤진다
    무엇을 검사할까요   ★ 고정 목록이 아니라 **그 엑셀에서 읽어** 채운다
    [검사 시작]

중간에 멈춰 묻지 않는다
─────────────────────
전에는 "어느 서류를 봤는지 확인해 주세요" 하고 멈췄다.  사람이 파일 이름만
보고 맞는지 알 수 없으니 쓸모가 없었다.  지금은 끝까지 가고, 봐야 할 것은
추출 내역의 「봐야 할 것」 에 모아 적는다.

판정 로직은 한 줄도 없다.  `core.flow` 만 부른다.
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import theme as T
from .runner import Runner
from .widgets import DropZone

제목 = "성과관리현황 검사 보조도구"


def _열기(p):
    p = str(p)
    try:
        if sys.platform.startswith("win"):
            os.startfile(p)                              # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except Exception:                                    # noqa: BLE001
        pass


class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="white")
        self.root = root
        self.pack(fill="both", expand=True)
        self.책 = ""
        self.폴더 = ""
        self.시트들: list = []
        self.결과: dict = {}
        self.runner = Runner(root, self._받음)
        self._짓기()

    # ── 화면 ──────────────────────────────────────────────
    def _짓기(self):
        tk.Label(self, text=제목, font=T.FONT_L, bg="white").pack(pady=(14, 2))
        tk.Label(self, text="파일 이름에 번호가 없어도 서류 안을 읽어 검사 칸을 채웁니다",
                 font=T.FONT, fg=T.MUTED, bg="white").pack(pady=(0, 10))

        존 = tk.Frame(self, bg="white")
        존.pack(fill="x", padx=16)
        self.존1 = DropZone(존, "① 성과관리현황 엑셀\n(검사 칸이 든 것)",
                            self._엑셀, pad=(0, 6))
        self.존2 = DropZone(존, "② 증빙 폴더\n(폴더 안 폴더까지 봅니다)",
                            self._폴더, pad=(6, 0))

        고름 = tk.Frame(self, bg="white")
        고름.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(고름, text="③ 무엇을 검사할까요", font=T.FONT_B,
                 bg="white").pack(side="left")
        self.시트칸 = ttk.Combobox(고름, state="disabled", font=T.FONT, width=42)
        self.시트칸.pack(side="left", padx=8)
        self.시트말 = tk.Label(고름, text="엑셀을 먼저 넣어 주세요",
                               font=T.FONT, fg=T.FAINT, bg="white")
        self.시트말.pack(side="left")

        self.열쇠말 = tk.Label(self, text="", font=T.FONT_B, fg="#C00000",
                               bg="white", wraplength=640, justify="left")
        self.열쇠말.pack(fill="x", padx=16)

        단 = tk.Frame(self, bg="white")
        단.pack(fill="x", padx=16, pady=10)
        self.시작 = tk.Button(단, text="검사 시작", font=T.FONT_L, width=14,
                              bg=T.ACCENT, fg=T.ACC_FG, relief="flat",
                              command=self._시작)
        self.시작.pack(side="left")
        self.중지 = tk.Button(단, text="중지", font=T.FONT, width=8, state="disabled",
                              command=self.runner.stop)
        self.중지.pack(side="left", padx=6)
        self.열기 = tk.Button(단, text="결과 폴더 열기", font=T.FONT, width=14,
                              state="disabled",
                              command=lambda: _열기(Path(
                                  self.결과.get("검사결과", ".")).parent))
        self.열기.pack(side="right")

        self.말 = tk.Label(self, text="", font=T.FONT, bg="white", anchor="w")
        self.말.pack(fill="x", padx=16)
        self.바 = ttk.Progressbar(self, maximum=1000)
        self.바.pack(fill="x", padx=16, pady=(2, 8))

        self.펼침 = tk.BooleanVar(value=False)
        tk.Checkbutton(self, text="자세히", variable=self.펼침, bg="white",
                       font=T.FONT, command=self._펼치기).pack(anchor="w", padx=16)
        self.기록칸 = tk.Frame(self, bg="white")
        self.기록 = tk.Text(self.기록칸, font=T.FONT, height=14, wrap="word",
                            bg=T.PANEL, relief="flat", state="disabled")
        sb = tk.Scrollbar(self.기록칸, command=self.기록.yview)
        self.기록.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.기록.pack(fill="both", expand=True)
        for tag, (col, bold, _u) in T.LEVELS.items():
            self.기록.tag_config(tag, foreground=col,
                                 font=(T.FAMILY, 9, "bold") if bold else T.FONT)

        tk.Label(self, text="원본은 건드리지 않습니다.  검사 칸은 사본에 쓰고, "
                            "쓰기 전에 백업을 한 벌 더 뜹니다.",
                 font=T.FONT_S, fg=T.FAINT, bg="white").pack(pady=(6, 10))
        self._열쇠보기()

    def _펼치기(self):
        if self.펼침.get():
            self.기록칸.pack(fill="both", expand=True, padx=16, pady=(0, 6))
        else:
            self.기록칸.pack_forget()

    # ── 넣기 ──────────────────────────────────────────────
    def _엑셀(self, p=None):
        p = p or filedialog.askopenfilename(
            title="성과관리현황 엑셀", filetypes=[("엑셀", "*.xlsx *.xlsm")])
        if not p:
            return
        self.책 = p
        self.존1.ok(Path(p).name)
        self._시트채우기()

    def _폴더(self, p=None):
        p = p or filedialog.askdirectory(title="증빙 폴더")
        if not p:
            return
        self.폴더 = p
        self.존2.ok(Path(p).name)

    def _시트채우기(self):
        """★ 고정 목록이 아니라 그 엑셀에서 읽는다.

        없는 시트를 고를 수 있으면 골라 봐야 "시트가 없습니다" 로 끝난다.
        """
        self.시트칸.configure(state="disabled", values=[])
        self.시트말.configure(text="엑셀을 읽는 중…", fg=T.FAINT)
        self.update_idletasks()

        def 일():
            try:
                from core import flow
                본 = flow.시트고르기(self.책)
            except Exception as e:                       # noqa: BLE001
                본, 탈 = [], f"{type(e).__name__}: {e}"
            else:
                탈 = ""
            self.root.after(0, lambda: self._시트받음(본, 탈))

        threading.Thread(target=일, daemon=True).start()

    def _시트받음(self, 본, 탈):
        self.시트들 = 본
        if 탈:
            self.시트말.configure(text=f"엑셀을 읽지 못했습니다 — {탈[:60]}",
                                  fg="#C00000")
            return
        쓸것 = [s for s in 본 if s["검사칸"] and s["줄수"]]
        if not 쓸것:
            self.시트말.configure(
                text="검사 칸(주황)이 든 시트를 찾지 못했습니다 — "
                     "검수도구로 [준비하기] 를 먼저 하세요", fg="#C00000")
            return
        보임 = [f"{s['본']}   ({s['줄수']}줄)" for s in 쓸것]
        self.시트칸.configure(state="readonly", values=보임)
        self.시트칸.current(0)
        self._고른것 = {b: s["본"] for b, s in zip(보임, 쓸것)}
        self.시트말.configure(text=f"{len(쓸것)}개 중에서 고르세요", fg=T.MUTED)

    def _열쇠보기(self):
        try:
            from core import ask
            ask.열쇠읽기()
            ok, why = ask.available()
        except Exception as e:                           # noqa: BLE001
            ok, why = False, str(e)
        self.열쇠말.configure(
            text="" if ok else f"⚠ {why}\n   이 도구는 열쇠가 없으면 검사할 수 없습니다.")

    # ── 돌리기 ────────────────────────────────────────────
    def _시작(self):
        if not self.책 or not self.폴더:
            messagebox.showinfo("아직", "엑셀과 증빙 폴더를 둘 다 넣어 주세요.")
            return
        고름 = self.시트칸.get()
        시트 = getattr(self, "_고른것", {}).get(고름, "")
        if not 시트:
            messagebox.showinfo("아직", "검사할 시트를 골라 주세요.")
            return
        self.기록.configure(state="normal")
        self.기록.delete("1.0", "end")
        self.기록.configure(state="disabled")
        self.결과 = {}
        self.시작.configure(state="disabled")
        self.중지.configure(state="normal")
        self.열기.configure(state="disabled")
        self.말.configure(text="서류를 열고 있어요…", fg=T.MUTED)
        self.바.configure(value=0)

        from core import flow
        self.runner.start("검사", flow.검사, self.책, 시트, self.폴더)

    # ── 소식 ──────────────────────────────────────────────
    def _받음(self, m):
        종류 = m[0]
        if 종류 == "log":
            self._적기(m[1], m[2] if len(m) > 2 else "info")
        elif 종류 == "prog":
            i, n = m[1], m[2]
            self.바.configure(value=int(1000 * i / max(n, 1)))
        elif 종류 == "end":
            self._끝(m[2])
        elif 종류 == "stopped":
            self._멈춤()
        elif 종류 == "fail":
            # ui.runner 는 제 Cancelled 만 '중지' 로 안다.  core 는 제 것을 던지므로
            # (엔진이 화면을 몰라야 한다) 여기서 한 번 더 가린다.
            if self.runner.cancel.is_set():
                self._멈춤()
            else:
                self._탈(m[2], m[3] if len(m) > 3 else "")

    def _적기(self, 글, 수준="info"):
        self.기록.configure(state="normal")
        self.기록.insert("end", T.prefix(수준) + str(글) + "\n", 수준)
        self.기록.see("end")
        self.기록.configure(state="disabled")

    def _끝(self, res):
        self.결과 = res or {}
        self.시작.configure(state="normal")
        self.중지.configure(state="disabled")
        self.열기.configure(state="normal")
        self.바.configure(value=1000)
        c = self.결과.get("집계") or {}
        볼것 = [s for s in (self.결과.get("신호") or [])
                if s.get("수준") in ("err", "warn")]
        말 = (f"✓ 끝났어요.   O {c.get('O', 0)} · X {c.get('X', 0)} · "
              f"확인 불가 {c.get('확인 불가', 0)}")
        if 볼것:
            말 += f"    ❗ 봐야 할 것 {len(볼것)}가지 — 추출 내역의 「개요」 를 보세요"
        self.말.configure(text=말, fg="#1B7F3B" if not 볼것 else "#B06000")
        for k in ("검사결과", "추출내역"):
            if self.결과.get(k):
                self._적기(f"{k} — {Path(self.결과[k]).name}", "done")

    def _멈춤(self):
        self.시작.configure(state="normal")
        self.중지.configure(state="disabled")
        self.바.configure(value=0)
        self.말.configure(text="중지했습니다.  다시 [검사 시작] 을 누르면 "
                               "읽어 둔 서류는 그대로 씁니다.", fg="#B06000")
        self._적기("중지했습니다", "warn")

    def _탈(self, 왜, 자세히=""):
        self.시작.configure(state="normal")
        self.중지.configure(state="disabled")
        self.바.configure(value=0)
        self.말.configure(text="멈췄습니다.", fg="#C00000")
        self._적기(str(왜), "err")
        for ln in str(자세히).strip().splitlines()[-8:]:
            self._적기("    " + ln, "skip")
        if not self.펼침.get():
            self.펼침.set(True)
            self._펼치기()
        messagebox.showerror("멈췄습니다", str(왜)[:600])


def main() -> int:
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:                                    # noqa: BLE001
        root = tk.Tk()
    root.title(제목)
    root.geometry("760x640")
    root.configure(bg="white")
    App(root)
    root.mainloop()
    return 0
