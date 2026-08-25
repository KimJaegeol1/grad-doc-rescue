# -*- coding: utf-8 -*-
"""
paths — 프로그램이 자기 파일을 어디서 찾는지
================================================================
배포본 폴더를 깨끗하게 두려고 설정과 라이브러리를 한 폴더로 몰았다.

    보조도구/
        보조도구.exe
        보조도구 검수기준.xlsx      ← ★ 정본.  여기를 고치면 검사 규칙이 바뀐다
        엑셀_넣기_전에_읽어주세요.txt
        secrets/                 ← 열쇠 (밖에 둔다 · 가끔 손댈 일이 있다)
        프로그램파일/              ← 읽을거리 · 라이브러리

검수기준 엑셀을 **프로그램파일 안이 아니라 exe 옆**에 두는 것이 중요하다.
사람이 열어 고치는 물건이기 때문이다.  라이브러리 수백 개가 든 폴더에
숨겨 두면 아무도 못 찾는다.

소스로 돌 때는 이 개념이 없다 — 검수기준은 docs/ 안에 그대로 있다.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

INNER = "프로그램파일"          # build.spec 의 contents_directory 와 같아야 한다
SECRETS = "secrets"


def frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def exe_dir() -> Path:
    """exe 가 있는 폴더.  소스로 돌 때는 프로그램 루트."""
    if frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def inner() -> Path:
    """설정과 라이브러리가 든 폴더.  소스로 돌 때는 프로그램 루트."""
    return exe_dir() / INNER if frozen() else exe_dir()


def secrets() -> Path:
    """열쇠 폴더.  묶었을 때도 exe 옆에 그대로 둔다."""
    return exe_dir() / SECRETS


def config(name: str, fallback: str | os.PathLike | None = None) -> str:
    """설정 파일을 찾는다.

    묶었을 때  프로그램파일/<name>  →  없으면 exe 옆  →  없으면 fallback(번들 안 기본값)
    소스일 때  fallback (각 단계 폴더 안)
    """
    if frozen():
        for d in (inner(), exe_dir()):
            p = d / name
            if p.exists():
                return str(p)
    return str(fallback) if fallback else str(inner() / name)


def doc(name: str) -> str | None:
    """읽을거리를 찾는다 (안내문 등).  exe 옆 → 프로그램파일 안."""
    for d in (exe_dir(), inner()):
        p = d / name
        if p.exists():
            return str(p)
    return None
