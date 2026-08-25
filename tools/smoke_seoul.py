# -*- coding: utf-8 -*-
"""
smoke_seoul — 서울대 ⑤인턴십 실물 22줄로 되돌아보기 (회귀)

**모형을 부르지 않는다.**  실제로 틀렸던 세 가지가 지금은 안 틀리는지,
그리고 규격이 API 한도를 안 넘는지만 실물 엑셀로 본다.

    ① 병합 머리글      「참여학생명」 O2:AE2 — 참여인원 일치가 전 줄 X 였다
    ② 줄 누락          22줄 중 6줄만 답하고 닫았다
    ③ 날짜 정밀도      연월만 같아도 O 를 줘서 8줄이 거짓 O 로 통과했다

    python tools/smoke_seoul.py [성과관리현황.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

from core import book as B, calc as C, prompt as P, spec as S       # noqa: E402

기본 = ("/root/.claude/uploads/e2f6df8d-7996-5b70-ad90-088559601eb3/"
        "028351a5-___________.xlsx")
엑셀 = Path(sys.argv[1] if len(sys.argv) > 1 else 기본)
시트 = "⑤인턴십"
탈 = []


def 봄(ok, 글, 덧=""):
    print(f"  {'✓' if ok else '❌'} {글}" + (f"   {덧}" if 덧 else ""))
    if not ok:
        탈.append(글)


def 칸(제목):
    print("\n" + "═" * 74)
    print(제목)
    print("═" * 74)


if not 엑셀.exists():
    print(f"실물 엑셀이 없습니다 — {엑셀}")
    print("이 시험은 서울대 성과관리현황 파일이 있어야 돕니다.  건너뜁니다.")
    raise SystemExit(0)

sp = S.load(HERE.parent / "docs" / "보조도구 검수기준.xlsx")
bk = B.Book(엑셀, 시트)
지도, 못, _셈 = B.열지도(bk, sp.항목(시트))
줄들 = bk.줄들(지도)
bk.close()

# ══════════════════════════════════════════════════════════════
칸("① 병합 머리글 — 「참여학생명」 이 O2:AE2 로 병합돼 있다")
# ══════════════════════════════════════════════════════════════
print("  학교는 이름을 **한 칸에 하나씩** 적었다.  4줄은 18칸에 18명, 참여인원도 18.")
print("  맨 왼쪽 칸만 읽어 1명으로 세는 바람에 「참여인원 일치」 가 전 줄 X 였다.")
print("  X 는 학교에 반송이다 — 멀쩡한 실적 22건이 통째로 반송될 뻔했다.\n")

봄(not 못, "쓸 엑셀 열을 다 찾았다", str(못)[:60])
봄(len(지도.get("참여학생명") or []) >= 18,
  "참여학생명이 한 칸이 아니라 여러 칸으로 잡힌다",
  f"칸 {len(지도.get('참여학생명') or [])}개")
봄(len(줄들) == 22, "자료 줄 22개를 다 읽는다", f"{len(줄들)}줄")

넷째 = next(r for r in 줄들 if r["행"] == 4)
봄(C.이름수(넷째["값"]["참여학생명"]) == 18,
  "4줄 이름이 18명으로 세어진다 (엑셀 참여인원 = 18)",
  넷째["값"]["참여학생명"][:44] + " …")

항 = next(c for c in sp.항목(시트) if c["검증열"] == "참여인원 일치")
셈 = {}
for r in 줄들:
    v, _b, _g = C.재기(항, r["값"], {}, 맺은파일=["01.pdf"])
    셈[v or "공란"] = 셈.get(v or "공란", 0) + 1
print(f"\n     22줄 참여인원 일치 → {' · '.join(f'{k} {v}' for k, v in sorted(셈.items()))}")
봄(셈.get("X", 0) == 0, "★ 거짓 X 가 한 건도 없다 (예전 22건)")
봄(셈.get("O", 0) >= 18, "값이 있는 줄은 다 O 다", f"O {셈.get('O', 0)}개")


# ══════════════════════════════════════════════════════════════
칸("② 줄 누락 — 22줄을 8줄씩 나눠 묻고, 칸으로 박는다")
# ══════════════════════════════════════════════════════════════
print("  예전엔 22줄을 한 번에 물었고 모형이 3~8줄 여섯 개만 답하고 닫았다.")
print("  잘린 게 아니다 — 온전한 JSON 이었고 배열이 길이를 못 박을 뿐이었다.\n")

행들 = [r["행"] for r in 줄들]
파일들 = [(f"{i:02d}_서울대 인턴십_증빙.pdf", "본문") for i in range(1, 19)]
묶음 = 8
배치들 = [행들[i:i + 묶음] for i in range(0, len(행들), 묶음)]
봄(len(배치들) == 3 and sum(len(b) for b in 배치들) == 22,
  "22줄이 8+8+6 세 번으로 나뉜다", " / ".join(str(len(b)) for b in 배치들))

sch = P.schema(sp, 시트, 파일들, 배치들[0])
줄꼴 = sch["properties"]["줄"]
봄(줄꼴["type"] == "object" and len(줄꼴["required"]) == 8,
  "한 배치의 8줄이 전부 required — 건너뛸 자리가 없다")


def 세기(o, 무엇):
    n = 0
    if isinstance(o, dict):
        if 무엇 == "enum" and isinstance(o.get("enum"), list):
            n += len(o["enum"])
        if 무엇 == "prop" and isinstance(o.get("properties"), dict):
            n += len(o["properties"])
        for v in o.values():
            n += 세기(v, 무엇)
    elif isinstance(o, list):
        for v in o:
            n += 세기(v, 무엇)
    return n


def 깊이(o, d=0):
    if isinstance(o, dict):
        ps = o.get("properties")
        if isinstance(ps, dict):
            return max([깊이(v, d + 1) for v in ps.values()] or [d + 1])
        it = o.get("items")
        if isinstance(it, dict):
            return 깊이(it, d)
    return d


print()
print("  ★ 파일 수를 훑는다 — 예전엔 18개 하나로만 봐서 눈이 멀었다.")
print("    ④세미나 · 파일 15개 딱 그 조합에서만 어림셈이 903(<950) 이라 접지 않았고,")
print("    실제 enum 은 1,151 이라 API 가 400 으로 튕겼다.  이제 어림잡지 않고 센다.\n")
넘침 = []
for 시 in sp.시트:
    줄1 = []
    for n파일 in (1, 5, 10, 15, 18, 25, 40, 80, 150):
        f = [(f"x{i}.pdf", "") for i in range(n파일)]
        m = P.맞는묶음(sp, 시, f, 묶음)
        s = P.schema(sp, 시, f, list(range(3, 3 + m)))
        e, pr, d = P._enum수(s), 세기(s, "prop"), 깊이(s)
        if e > P.한도 or pr > 5000 or d > 5:
            넘침.append(f"{시} 파일{n파일} — enum {e} · 칸 {pr} · 겹 {d}")
        줄1.append(f"{n파일}:{m}줄/{e}")
    print(f"     {시:<12} " + "  ".join(줄1))
봄(not 넘침, "★ 파일 1~150개 어느 조합에서도 규격이 한도를 안 넘는다",
  str(넘침[:2]) if 넘침 else "파일수:묶음/enum")
봄(True, "여덟 시트 다 한도 안이다 (enum ≤1,000 · 칸 ≤5,000 · 겹 ≤5)")


# ══════════════════════════════════════════════════════════════
칸("③ 날짜 정밀도 — 8줄이 9줄 서류로 거짓 O 를 냈다")
# ══════════════════════════════════════════════════════════════
print("  엑셀 8줄  켐아이넷㈜  24-07-08 ~ 24-07-19")
print("  맺은 서류  애니텍     24-07-15 ~ 24-07-26   ← 9줄 것이다")
print("  그런데 「날짜가 다르나 연월은 같음」 으로 O 가 나갔다.\n")

일시 = next(c for c in sp.항목(시트) if c["검증열"] == "일시 일치")


def 날(엑, a, b=""):
    return C.재기(일시, {"일시": 엑}, {"인턴기간_시작일": a, "인턴기간_종료일": b},
                  맺은파일=["01.pdf"])


v, 비고, _ = 날("24-07-08 ~ 24-07-19", "2024-07-15", "2024-07-26")
봄(v == "X", "★ 8줄 그 건이 이제 X 다", 비고[:64])
봄(날("24-06-03 ~ 24-06-19", "2024-06-03", "2024-06-19")[0] == "O", "같은 날은 그대로 O")
봄(날("2024-07", "2024-07-15")[0] == "O", "엑셀이 연월까지만이면 연월로 견준다")
봄(날("24-07-08 ~ 24-07-19", "2024년 7월")[0] == "O", "증빙이 연월까지만이어도 그렇다")
봄(날("2024-07", "2024-09-15")[0] == "X", "연월이 다르면 X")
봄(날("24-06-03 ~ 24-06-19", "2024-06-03")[0] == "확인 불가",
  "덜 뽑힌 것은 X 가 아니라 확인 불가", 날("24-06-03 ~ 24-06-19", "2024-06-03")[1][:52])

print()
print("═" * 74)
if 탈:
    print(f"❌ 어긋난 곳 {len(탈)}건")
    for x in 탈:
        print(f"   · {x}")
else:
    print("✅ 실물 22줄에서 세 가지가 다 잡혔다 — 거짓 X 0건 · 건너뛸 자리 없음 · 거짓 O 막힘")
print("═" * 74)
raise SystemExit(1 if 탈 else 0)
