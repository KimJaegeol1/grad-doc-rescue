# -*- coding: utf-8 -*-
"""
gui.runner — 작업 실행기 (스레드 + 큐 + 취소)
================================================================
화면은 멈추면 안 되고, 단계는 tkinter 를 몰라야 한다.  그 사이를 잇는 부품.

단계에게 넘겨 주는 것은 셋뿐이다.
    log(글, 수준)      진행 기록 한 줄
    prog(한 것, 전부)  진행바
    cancel             threading.Event — 단계가 스스로 들여다보고 협조한다

단계가 화면 부품을 직접 만지는 일은 없다.  전부 큐를 거친다.
"""

from __future__ import annotations

import queue
import threading
import traceback


class Cancelled(Exception):
    """사람이 [중지] 를 눌렀다."""


class Runner:
    def __init__(self, root, on_msg):
        """
        root   : tkinter 루트 (after 를 쓰기 위해서만)
        on_msg : (종류, 내용...) 을 받는 화면 쪽 처리기
        """
        self.root = root
        self.on_msg = on_msg
        self.q: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self._thread: threading.Thread | None = None
        self.root.after(80, self._pump)

    # ── 상태 ──────────────────────────────────────────────
    @property
    def busy(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    # ── 실행 ──────────────────────────────────────────────
    def start(self, tag: str, fn, *args, **kw) -> bool:
        """fn(*args, log=, prog=, cancel=, **kw) 을 딴 실 위에서 부른다.

        끝나면 ("end", tag, 결과) 또는 ("fail", tag, 사유, 자세히) 가 큐로 온다.
        """
        if self.busy:
            return False
        self.cancel.clear()

        def _log(msg, level="info"):
            self.q.put(("log", str(msg), level))

        def _prog(i, n):
            self.q.put(("prog", i, n))

        def _body():
            try:
                res = fn(*args, log=_log, prog=_prog, cancel=self.cancel, **kw)
                self.q.put(("end", tag, res))
            except Cancelled:
                self.q.put(("stopped", tag))
            except Exception as e:                       # noqa: BLE001
                self.q.put(("fail", tag, f"{type(e).__name__}: {e}",
                            traceback.format_exc()))

        self._thread = threading.Thread(target=_body, daemon=True)
        self._thread.start()
        self.q.put(("begin", tag))
        return True

    def stop(self):
        self.cancel.set()

    # ── 큐 비우기 ─────────────────────────────────────────
    def _pump(self):
        try:
            while True:
                self.on_msg(self.q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(80, self._pump)

    # ── 화면 쪽에서 직접 한 줄 남길 때 ─────────────────────
    def log(self, msg, level="info"):
        self.q.put(("log", str(msg), level))
