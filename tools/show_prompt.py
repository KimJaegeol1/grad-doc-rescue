# -*- coding: utf-8 -*-
"""
show_prompt — 모형에게 **실제로 갈 글**을 파일로 뽑는다.

    python tools/show_prompt.py <성과관리현황.xlsx> <시트> [증빙폴더]

증빙 폴더를 안 주면 파일 글은 짧은 본으로 채운다 (글 짜임만 보려 할 때).
모형을 부르지 않는다.  돈이 안 든다.

왜 이런 도구가 필요한가
──────────────────────
검수기준을 고치면 프롬프트가 바뀐다.  바뀐 글을 눈으로 못 보면, 고친 것이
제대로 실렸는지 모형을 불러 봐야만 안다 — 돈도 시간도 든다.
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

from core import book as B, prompt as P, spec as S            # noqa: E402

기준 = HERE.parent / "docs" / "보조도구 검수기준.xlsx"

본보기 = [
    ("01_202406_인턴십_과학원.pdf",
     "인턴십 계획서\n인턴기관: 국립환경과학원\n인턴기간: 24년 06월 03일 ~ 24년 06월 16일\n"
     "참여학생: 서혜빈\n총 80시간\n"
     "===== [p.2] =====\n결과보고서 요약서\n대기오염 측정망 자료 분석 보조\n"
     "===== [p.3] =====\n세부 결과보고서\n1일차 오리엔테이션 …\n현장 사진 3매\n"),
    ("02_202406_인턴십_그리디그리드.pdf",
     "인턴십 계획서\n기업명: 그리디그리드\n기간: 24-06-03 ~ 24-06-19\n참여학생: 김은진\n"),
]


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    책, 시트 = sys.argv[1], sys.argv[2]
    폴더 = sys.argv[3] if len(sys.argv) > 3 else ""

    sp = S.load(기준)
    for 글, _ in sp.요약():
        print(글)
    if sp.탈:
        print("\n※ 검수기준에 어긋난 곳이 있습니다 — 위를 보세요\n")

    bk = B.Book(책, 시트)
    지도, 못, _셈 = B.열지도(bk, sp.항목(시트))
    줄들 = bk.줄들(지도)
    명단 = bk.별첨1()
    print(f"\n엑셀 「{bk.ws.title}」 · 머리글 {bk.머리글줄}행 · "
          f"자료 열 {len(지도)}개 · 줄 {len(줄들)}개 · 별첨1 {len(명단)}명")
    for x in 못:
        print(f"  ❌ 「{x['검증열']}」 이 쓸 열을 못 찾았습니다 "
              f"(찾은 이름: {' · '.join(x['찾은이름'])})")

    if 폴더:
        from engines import read as R
        todo, skip = R.from_folder(Path(폴더), recursive=True)
        print(f"증빙 폴더 — 읽을 것 {len(todo)}개 · 건너뜀 {len(skip)}개")
        파일들 = []
        for p in todo:
            try:
                t = R.read_one(Path(p)) if hasattr(R, "read_one") else None
                파일들.append((Path(p).name, t or ""))
            except Exception as e:                       # noqa: BLE001
                파일들.append((Path(p).name, f"(읽지 못했습니다 — {e})"))
    else:
        파일들 = 본보기
        print(f"증빙 폴더를 안 주셔서 본보기 파일 {len(파일들)}개로 채웁니다")

    묶음 = 8                      # core.ask.run 의 기본값과 같아야 한다
    물을줄 = [r["행"] for r in 줄들][:묶음]
    p = P.build(sp, 시트, 줄들, 명단, 파일들, 물을줄=물을줄)
    bk.close()

    안전 = "".join(ch for ch in 시트 if ch not in '\\/:*?"<>|')
    out = HERE.parent / "workspace" / f"프롬프트_{안전}.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        "◆ SYSTEM\n" + p["system"] + "\n\n◆ USER\n" + p["user"], "utf-8")

    print()
    print("═" * 70)
    for k, v in p["잰것"].items():
        print(f"  {k:<12} {v:,}" if isinstance(v, int) else f"  {k:<12} {v}")
    if p["자름"]:
        print(f"  ※ 자른 파일 {len(p['자름'])}개")
        for x in p["자름"][:5]:
            print(f"      {x['파일']}  {x['전체']:,}자 중 {x['실은것']:,}자만")
    print("═" * 70)
    print(f"뽑았습니다 — {out}")

    sch = P.schema(sp, 시트, 파일들, 물을줄)
    줄꼴 = sch["properties"]["줄"]
    칸 = 줄꼴["properties"][P.행키(물을줄[0])]["properties"]
    판정칸 = [k for k in 칸 if k.startswith("v") and k[1:].isdigit()]
    묶개 = -(-len(줄들) // 묶음)
    print(f"\n답 꼴: 줄 {len(줄들)}개를 {묶개}번에 나눠 묻는다 (한 번에 {묶음}줄)")
    print(f"      한 배치의 줄은 {' · '.join(줄꼴['required'][:4])}"
          f"{' …' if len(줄꼴['required']) > 4 else ''} — **전부 required**, "
          f"건너뛸 자리가 없다")
    print(f"      줄마다 값 {len(칸['값']['properties'])}가지 · 판정 {len(판정칸)}칸")
    표 = 칸["맺은파일"]["items"].get("enum") or []
    print(f"      맺은파일은 표 {len(표)}개 중에서만 고른다 "
          f"({', '.join(표[:3])}{' …' if len(표) > 3 else ''}) — 지어낼 자리가 없다")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
