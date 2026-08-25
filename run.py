# -*- coding: utf-8 -*-
"""
성과관리현황 검사 보조도구

시트 하나와 증빙 폴더 하나를 받아, 폴더 안 폴더까지 파일을 전부 열고
그 시트를 통째로 모형에게 맡겨 검사 칸을 채운다.

    python run.py

언제 쓰나
    학교가 파일 이름에 번호를 안 붙여 검수도구 2단계가 막힐 때.
    번호가 제대로 붙어 있으면 검수도구를 쓰는 게 맞다 — 번호는 사람이
    맺어 준 보증이고, 내용 추정보다 언제나 낫다.

필요한 것
    · 검사 칸(주황)이 든 성과관리현황 엑셀
    · 「보조도구 검수기준.xlsx」   프로그램 옆에.  ★ 검사 규칙 정본
    · secrets/.env 의 OPENAI_API_KEY
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    from preflight import guard
    if not guard():
        sys.exit(1)
    from ui.app import main
    sys.exit(main())
