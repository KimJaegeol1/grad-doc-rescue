# -*- coding: utf-8 -*-
"""
core.calc — 수와 날짜는 코드가 센다.
================================================================
검수기준의 「판정 주체」 가 '계산' 인 항목을 여기서 판정한다.  모형은 값을
옮겨 적는 데까지만 하고, 세는 것은 코드가 한다.

왜 모형에게 안 맡기나
────────────────────
모형은 40시간과 4시간을 헷갈리고, ISBN 체크디지트를 못 세고, 「25-02-10 ~
25-02-14」 를 5일이라 했다가 4일이라 했다가 한다.  수와 날짜는 코드가 언제나
같은 답을 낸다.  **같은 답을 내는 것 자체가 값어치다** — 어제 O 였던 칸이
오늘 X 가 되면 아무도 그 도구를 못 믿는다.

무슨 셈인지는 어디서 아나
────────────────────────
검수기준의 「셈법」 칸이다.  판정 방법 글에도 "3개월 이상" 이라 적혀 있지만
**글을 읽어 셈을 고르지 않는다** — 사람이 문구를 다듬는 순간 셈이 바뀐다.

    날짜        양쪽을 연월일로 환산해 대조
    수          양쪽을 수로 읽어 대조
    이름개수     엑셀 수 ↔ 엑셀 이름 개수 (증빙을 안 본다)
    기간 ≥ N개월  엑셀 기간의 길이
    값 ≥ N       증빙 값이 임계 이상인가
    ISBN 유효    형식과 검증숫자
    ISBN 같음    하이픈 무시 후 같은가
    학기        연도와 학기 번호로 환산

못 읽으면 X 가 아니다
────────────────────
「2024년 5월」 처럼 날까지 없거나, 「8H×5일」 처럼 곱해야 알거나, 글자가
깨졌으면 **확인 불가**다.  못 읽은 것을 반송하면 안 된다.
"""

from __future__ import annotations

import re
import unicodedata

from . import spec as _S

O, X, UNK, BLANK = "O", "X", "확인 불가", ""

# ★ 엑셀 안에서 끝나는 셈 — 맺은 서류가 없어도 낼 수 있다 (개요 제0조 ③)
#   기간      엑셀에 적힌 기간이 3개월 이상인가
#   이름개수   엑셀 인원 수 ↔ 엑셀 이름 칸의 이름 개수
#   ISBN유효   엑셀에 적힌 ISBN 의 검증숫자가 맞는가
#   나머지(날짜·수·값·ISBN같음·학기)는 **증빙과 견주는 셈**이라 서류가 있어야 한다.
엑셀만셈 = {"기간", "이름개수", "ISBN유효"}


# ══════════════════════════════════════════════════════════════
# 읽기
# ══════════════════════════════════════════════════════════════
_날짜 = re.compile(r"(\d{2,4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})")
_연월 = re.compile(r"(\d{4})\s*[.\-/년]\s*(\d{1,2})(?!\s*[.\-/]?\s*\d)")
_짧은날 = re.compile(r"[~\-–]\s*(\d{1,2})\s*[.\-/일]?\s*$")


def _연도(y: str) -> int:
    """두 자리 연도를 편다.  24 → 2024.

    실제로 이것 하나를 못 읽어 인턴십 22줄이 통째로 미배정된 적이 있다.
    """
    n = int(y)
    return n if n >= 1000 else 2000 + n


def 날짜들(s) -> list:
    """글에서 연월일을 전부 뽑는다.  ["20240501", "20240831"].

    「2024.05.01-2024.08.31」  「24-06-03 ~ 24-06-16」  「24년 06월 03일」
    「25-08-25~29」 처럼 뒤가 생략된 것은 앞 날짜의 연·월을 이어받는다.
    """
    t = unicodedata.normalize("NFC", str(s or ""))
    out = []
    for y, m, d in _날짜.findall(t):
        out.append(f"{_연도(y):04d}{int(m):02d}{int(d):02d}")
    m = _짧은날.search(t)
    if m and len(out) == 1:                      # 「25-08-25~29」
        앞 = out[0]
        out.append(f"{앞[:6]}{int(m.group(1)):02d}")
    return out


def 연월들(s) -> list:
    """연월까지만 적힌 것.  ["202408"].  날까지 있으면 그것이 우선이다."""
    t = unicodedata.normalize("NFC", str(s or ""))
    return [f"{int(y):04d}{int(m):02d}" for y, m in _연월.findall(t)]


def 수(s, 힌트: str = "") -> float | None:
    """글에서 수 하나를 읽는다.  못 정하면 None.

    힌트를 주면 그 단위 앞의 수를 먼저 찾는다 — 「25-02-10 ~ 25-02-14(40시간)」
    에서 40 을 집으려면 그래야 한다.  힌트가 없고 수가 여럿이면 **정하지 않는다**
    (아무거나 집으면 거짓 판정이 난다).
    """
    t = str(s or "").replace(",", "")
    if 힌트:
        m = re.search(rf"(\d+(?:\.\d+)?)\s*(?:{힌트})", t, re.IGNORECASE)
        if m:
            return float(m.group(1))
    수들 = re.findall(r"\d+(?:\.\d+)?", t)
    return float(수들[0]) if len(수들) == 1 else None


def 시간(s) -> float | None:
    """시간 값.  「40」 「40H」 「40시간」 「(20H)+(20H)」 를 읽는다.

    나눠 적힌 것은 합산한다 (검수기준: 분할 표기는 합산).
    곱셈 표기(「8H×5일」)는 **읽지 않는다** — 무엇을 곱하라는 건지 코드가
    단정할 수 없다.  확인 불가로 두고 사람에게 넘긴다.
    """
    t = str(s or "")
    if re.search(r"[×xX*]", t):
        return None
    조각 = re.findall(r"(\d+(?:\.\d+)?)\s*(?:시간|H|h)", t)
    if len(조각) > 1:
        return sum(float(x) for x in 조각)
    if len(조각) == 1:
        return float(조각[0])
    return 수(t)


def 이름수(s) -> int:
    """「김재걸, 이수민 · 박민수」 → 3.  쉼표·가운뎃점·줄바꿈·빗금으로 가른다.

    ★ 길이 상한이 12자였다.  외국인 학생 「Gokce Nul AYAZ」(14자) ·
      「Salmaddina Aulia」(16자) 가 이름이 아닌 것으로 빠져, 서울대 4·5·9줄이
      「엑셀 18 ≠ 이름 16」 처럼 어긋나 X 가 났다.  X 는 학교에 반송이다.
      상한을 30자로 올린다 — 이름은 이제 core.book 이 **칸마다 하나씩**
      모아 오므로 긴 글이 섞여 들 자리가 거의 없다.
    """
    t = str(s or "").strip()
    if not t:
        return 0
    조각 = [x.strip() for x in re.split(r"[,·、/\n;]+", t)]
    return len([x for x in 조각 if 1 < len(x) <= 30 and not x.isdigit()])


def 명단수(값: dict) -> int | None:
    """증빙에서 뽑은 사람 수.  못 찾았으면 None (0 과 다르다).

    ★ 「찾음」 이라면서 이름을 하나도 안 적은 답이 온다.  규격상 `이름: []` 도
      유효하기 때문이다.  그것을 0 으로 세면 「엑셀 40 ↔ 증빙 0」 같은 **거짓
      X** 가 만들어져 학교에 반송이 나간다.  실제로 서울대 11~18줄 여덟 건이
      그렇게 X 를 받았다.  이름이 없으면 못 찾은 것이다 — 0 이 아니라 None.
    """
    v = 값.get(_S.명단칸)
    if isinstance(v, dict):
        if v.get("상태") != "찾음":
            return None
        n = len([x for x in (v.get("이름") or []) if str(x).strip()])
        return n or None
    if isinstance(v, list):
        return len([x for x in v if str(x).strip()]) or None
    n = 이름수(v)
    return n or None


_ISBN숫자 = re.compile(r"[\dXx]")


def isbn숫자(s) -> str:
    return "".join(_ISBN숫자.findall(str(s or ""))).upper()


def isbn유효(s) -> bool | None:
    """체크디지트까지 본다.  자릿수가 10·13 이 아니면 None (판단 불가)."""
    d = isbn숫자(s)
    if len(d) == 13 and d.isdigit():
        합 = sum(int(x) * (1 if i % 2 == 0 else 3) for i, x in enumerate(d[:12]))
        return (10 - 합 % 10) % 10 == int(d[12])
    if len(d) == 10:
        try:
            합 = sum((10 - i) * (10 if x == "X" else int(x)) for i, x in enumerate(d))
        except ValueError:
            return None
        return 합 % 11 == 0
    return None


def 학기(s) -> tuple | None:
    """「2025-1」 「2025학년도 1학기」 「25-1」 → (2025, 1)."""
    t = unicodedata.normalize("NFC", str(s or ""))
    m = re.search(r"(\d{2,4})\s*(?:학년도)?\s*[-.\s]?\s*(\d)\s*학기", t)
    if not m:
        m = re.search(r"(\d{2,4})\s*[-.]\s*([12])\b", t)
    if not m:
        return None
    return _연도(m.group(1)), int(m.group(2))


# ══════════════════════════════════════════════════════════════
# 판정
# ══════════════════════════════════════════════════════════════
def 재기(항목: dict, 엑셀값: dict, 증빙값: dict, *,
        맺은파일=None, 앞선판정=None) -> tuple:
    """계산 항목 하나를 판정한다.  반환 (판정, 비고, 근거)."""
    셈 = (항목.get("셈법") or {})
    이름 = _짧게(항목["검증열"])
    if not 셈.get("셈"):
        return UNK, f"{이름} 을(를) 어떻게 셀지 검수기준에 적혀 있지 않음", "셈법 빈칸"

    # 앞선 항목이 X·확인 불가면 물을 것이 없다
    앞 = 셈.get("앞선")
    if 앞 is not None:
        v = (앞선판정 or {}).get(앞)
        if v in (X, UNK):
            return BLANK, f"앞선 {앞}번이 {v or '공란'}", "종속"

    엑 = _엑셀(항목, 엑셀값)
    종류 = 셈["셈"]

    # ★ 맺은 서류가 없으면 증빙에 기댄 판정을 내지 아니한다 (개요 제0조 ③)
    # ────────────────────────────────────────────────────────
    #   모형 판정에는 이 장치가 있었는데(ask._내리기) **계산 판정에는 없었다.**
    #   그래서 서울대 22·23줄이 파일을 하나도 못 맺었는데도
    #
    #       40시간 이상 여부 = O   「코드가 셈 · 증빙 40 ≥ 40」
    #       시간 일치        = O   「코드가 셈 · 엑셀 40 = 증빙 40」
    #
    #   을 받았다.  모형이 맺지도 않은 서류의 값을 값 칸에 적어 온 것을
    #   코드가 그대로 믿은 것이다 — **거짓 O** 다.
    #
    #   새 규칙이 아니라 ③ 을 계산 쪽에 구현하지 않은 구멍이었다.
    #   기간·이름개수는 엑셀 안에서 끝나는 셈이라 서류가 없어도 낼 수 있다.
    if 종류 not in 엑셀만셈 and not (맺은파일 or []):
        return UNK, "맺은 서류가 없어 증빙과 대조할 수 없음", "증빙 못 맺음"

    if 종류 == "기간":
        return _기간(셈, 엑, 이름)
    if 종류 == "이름개수":
        return _이름개수(셈, 엑, 엑셀값, 이름)
    if 종류 == "값":
        return _값임계(셈, 증빙값, 항목, 이름)
    if 종류 == "날짜":
        return _날짜대조(엑, 증빙값, 항목, 이름)
    if 종류 == "수":
        return _수대조(엑, 증빙값, 항목, 이름)
    if 종류 == "ISBN유효":
        return _isbn유효(엑, 증빙값, 이름)
    if 종류 == "ISBN같음":
        return _isbn같음(엑, 증빙값, 이름)
    if 종류 == "학기":
        return _학기(엑, 증빙값, 이름)
    return UNK, f"{이름} — 모르는 셈법 「{셈.get('글')}」", "셈법 미구현"


def _엑셀(항목: dict, 엑셀값: dict) -> str:
    for 이름 in (항목.get("엑셀열") or []):
        v = str((엑셀값 or {}).get(이름, "") or "").strip()
        if v:
            return v
    return ""


def _증빙(항목: dict, 증빙값: dict, 골라=None) -> str:
    for 이름 in (골라 or 항목.get("뽑을값") or []):
        v = (증빙값 or {}).get(이름)
        if isinstance(v, (list, dict)):
            continue
        v = str(v or "").strip()
        if v:
            return v
    return ""


# ── 셈마다 ────────────────────────────────────────────────────
def _기간(셈, 엑: str, 이름: str) -> tuple:
    if not 엑:
        return BLANK, "엑셀에 기간이 없음", "엑셀 공란"
    ds = 날짜들(엑)
    if len(ds) < 2:
        return UNK, f"기간을 날짜 둘로 읽지 못함 — 엑셀 「{엑[:30]}」", "날짜 못 읽음"
    a, b = sorted(ds)[0], sorted(ds)[-1]
    개월 = (int(b[:4]) - int(a[:4])) * 12 + (int(b[4:6]) - int(a[4:6]))
    if int(b[6:]) < int(a[6:]):
        개월 -= 1
    임계 = 셈.get("임계", 3)
    남 = (int(b[:4]) * 12 + int(b[4:6])) - (int(a[:4]) * 12 + int(a[4:6]))
    말 = f"{_보기(a)} ~ {_보기(b)} · 약 {개월}개월"
    if 개월 >= 임계:
        return O, "", 말
    return X, f"{말} — {임계:.0f}개월 미만", 말


def _이름개수(셈, 엑: str, 엑셀값: dict, 이름: str) -> tuple:
    칸 = 셈.get("이름칸") or ""
    이름글 = str((엑셀값 or {}).get(칸, "") or "").strip()
    if not 엑 or not 이름글:
        빈 = "인원 수" if not 엑 else f"「{칸}」"
        return BLANK, f"엑셀에 {빈} 가 없음", "엑셀 공란"
    n = 수(엑)
    if n is None:
        return UNK, f"엑셀 인원 「{엑[:20]}」 을 수로 읽지 못함", "수 못 읽음"
    m = 이름수(이름글)
    if int(n) == m:
        return O, "", f"엑셀 {int(n)}명 · 이름 {m}명"
    return X, f"엑셀 {int(n)}명  ↔  이름 {m}명", f"엑셀 {int(n)} ≠ 이름 {m}"


def _값임계(셈, 증빙값: dict, 항목: dict, 이름: str) -> tuple:
    글 = _증빙(항목, 증빙값)
    if not 글:
        return UNK, "증빙에서 값을 옮겨 적지 못함", "증빙 값 없음"
    v = 시간(글)
    if v is None:
        return UNK, f"증빙 「{글[:24]}」 를 수 하나로 읽지 못함", "수 못 읽음"
    임계 = 셈.get("임계", 0)
    if v >= 임계:
        return O, "", f"증빙 {v:g} ≥ {임계:g}"
    return X, f"{v:g} — {임계:g} 미만", f"증빙 {v:g} < {임계:g}"


def _날짜대조(엑: str, 증빙값: dict, 항목: dict, 이름: str) -> tuple:
    """엑셀 날짜와 증빙 날짜를 견준다.

    ★ 연월만 같아도 O 를 주던 규칙을 걷어냈다
    ─────────────────────────────────────
    양쪽 다 날이 온전한데 서로 다른데도 연월이 같으면 O 를 냈다.  그래서
    서울대 8줄이 이렇게 통과했다.

        엑셀   켐아이넷㈜   24-07-08 ~ 24-07-19
        증빙   애니텍      24-07-15 ~ 24-07-26     ← 9줄 서류다
        일시 일치 = O   「날짜가 다르나 연월은 같음」

    엉뚱한 서류에 맺힌 줄이 **거짓 O** 로 통과했다.  거짓 X 는 사람이 다시
    볼 자리라도 생기지만 거짓 O 는 그대로 나간다.  그래서 이제

        양쪽 다 연월일   →  일자까지 같아야 O.  다르면 X
        한쪽만 연월까지  →  그때만 연월로 견준다 (24-07 ↔ 24-07-15 는 O)

    연월 비교는 **적힌 것이 연월까지뿐일 때** 봐주는 것이지, 날이 어긋난 것을
    눈감아 주는 규칙이 아니었다.
    """
    if not 엑:
        return BLANK, "엑셀에 값이 없음", "엑셀 공란"
    증글 = " ~ ".join(
        str((증빙값 or {}).get(f, "") or "").strip()
        for f in (항목.get("뽑을값") or []) if str((증빙값 or {}).get(f, "") or "").strip())
    if not 증글:
        return UNK, "증빙에서 날짜를 옮겨 적지 못함", "증빙 값 없음"
    엑d, 증d = 날짜들(엑), 날짜들(증글)

    # ── 양쪽 다 날이 온전하다 — 일자까지 같아야 한다 ──────
    #    ★ 목록이 아니라 **모둠(set)** 으로 견준다.  「2024-05-03」 한 날짜를
    #      날짜들() 이 시작·끝 둘로 돌려주는 일이 있어, 목록으로 견주면
    #      같은 날인데도 개수가 달라 X 가 났다 (④세미나 #7).
    if 엑d and 증d:
        엑s, 증s = set(엑d), set(증d)
        if 엑s == 증s:
            return O, "", " · ".join(_보기(x) for x in sorted(엑s))
        if 엑s < 증s or 증s < 엑s:
            # 한쪽이 다른 쪽에 통째로 들어 있다 — 어긋난 게 아니라 **덜 뽑힌** 것이다.
            # 못 뽑은 것을 X 로 하면 멀쩡한 실적이 반송된다.
            덜 = "증빙" if 증s < 엑s else "엑셀"
            return UNK, (f"{덜} 에서 날짜를 다 옮겨 적지 못해 가리지 못함 — "
                         f"엑셀 {_보기목록(엑s)}  ↔  증빙 {_보기목록(증s)}"), "날짜 덜 뽑힘"
        같은달 = {x[:6] for x in 엑s} == {x[:6] for x in 증s}
        return X, (f"엑셀 {_보기목록(엑s)}  ↔  증빙 {_보기목록(증s)}"
                   + ("  — 연월은 같으나 날이 다릅니다" if 같은달 else "")), "날짜 다름"

    # ── 한쪽이 연월까지만 적혀 있다 — 그때만 연월로 견준다 ──
    엑ym = set(연월들(엑)) or {x[:6] for x in 엑d}
    증ym = set(연월들(증글)) or {x[:6] for x in 증d}
    if not 엑ym:
        return UNK, f"엑셀 「{엑[:26]}」 를 날짜로 읽지 못함", "엑셀 날짜 못 읽음"
    if not 증ym:
        return UNK, f"증빙 「{증글[:26]}」 를 날짜로 읽지 못함", "증빙 날짜 못 읽음"
    어느 = "증빙" if not 증d else "엑셀"
    if 엑ym & 증ym:
        return O, f"{어느} 이 연월까지만 적혀 연월로 견주었습니다 — 「{증글[:24]}」", "연월 일치"
    return X, (f"엑셀 {_보기목록(sorted(엑ym))}  ↔  증빙 {_보기목록(sorted(증ym))}"
               f"  ({어느} 은 연월까지만 적혀 있습니다)"), "연월 다름"


def _수대조(엑: str, 증빙값: dict, 항목: dict, 이름: str) -> tuple:
    if not 엑:
        return BLANK, "엑셀에 값이 없음", "엑셀 공란"
    a = 시간(엑) if "시간" in 항목["검증열"] else 수(엑)
    if a is None:
        return UNK, f"엑셀 「{엑[:24]}」 를 수로 읽지 못함", "수 못 읽음"
    글 = _증빙(항목, 증빙값)
    b = 시간(글) if 글 else None
    if b is None:
        # ★ 명단 수로 갈음하는 것은 **검수기준이 「참여명단」 을 뽑을 값으로
        #   적어 둔 항목에서만** 한다.  ③프로젝트 「참여학생 수 일치」 는
        #   뽑을값이 참여인원수·참여명단 이라 갈음이 옳고, ⑤인턴십 「시간 일치」 는
        #   총시간뿐이라 옳지 않다.  사람 수를 시간으로 갖다 쓰면 말이 안 된다.
        #
        #   실제로 「시간 일치」 가 증빙이 통째로 빈 줄에서 명단 0명을 0시간으로
        #   읽어 「엑셀 40 ↔ 증빙 0」 거짓 X 를 여덟 건 냈다.
        #
        #   검증열 이름의 낱말로 가리지 않는다 — 사람이 이름을 다듬는 순간
        #   셈이 바뀐다 (이 파일 첫머리의 규약).
        b = 명단수(증빙값) if _S.명단칸 in (항목.get("뽑을값") or []) else None
        if b is None:
            return UNK, "증빙에서 수를 옮겨 적지 못함", "증빙 값 없음"
        b = float(b)
    if abs(a - b) < 1e-6:
        return O, "", f"엑셀 {a:g} = 증빙 {b:g}"
    return X, f"엑셀 {a:g}  ↔  증빙 {b:g}", f"{a:g} ≠ {b:g}"


def _isbn유효(엑: str, 증빙값: dict, 이름: str) -> tuple:
    글 = 엑 or str((증빙값 or {}).get("ISBN", "") or "")
    if not 글.strip():
        return BLANK, "엑셀에 ISBN 이 없음", "엑셀 공란"
    좋 = isbn유효(글)
    if 좋 is None:
        d = isbn숫자(글)
        return UNK, f"ISBN 자릿수가 {len(d)}자리 — 10·13자리가 아님", f"「{글[:24]}」"
    if 좋:
        return O, "", f"검증숫자 맞음 ({isbn숫자(글)})"
    return X, f"검증숫자가 맞지 않음 — 「{글[:24]}」", "체크디지트 불일치"


def _isbn같음(엑: str, 증빙값: dict, 이름: str) -> tuple:
    if not 엑:
        return BLANK, "엑셀에 ISBN 이 없음", "엑셀 공란"
    증 = str((증빙값 or {}).get("ISBN", "") or "").strip()
    if not 증:
        return UNK, "증빙에서 ISBN 을 옮겨 적지 못함", "증빙 값 없음"
    a, b = isbn숫자(엑), isbn숫자(증)
    if a == b:
        return O, "", f"{a}"
    return X, f"엑셀 「{엑[:20]}」  ↔  증빙 「{증[:20]}」", f"{a} ≠ {b}"


def _학기(엑: str, 증빙값: dict, 이름: str) -> tuple:
    if not 엑:
        return BLANK, "엑셀에 학기가 없음", "엑셀 공란"
    a = 학기(엑)
    if a is None:
        return UNK, f"엑셀 「{엑[:20]}」 를 학기로 읽지 못함", "학기 못 읽음"
    증 = str((증빙값 or {}).get("학기", "") or "").strip()
    if not 증:
        return UNK, "증빙에서 학기를 옮겨 적지 못함", "증빙 값 없음"
    b = 학기(증)
    if b is None:
        return UNK, f"증빙 「{증[:20]}」 를 학기로 읽지 못함", "학기 못 읽음"
    if a == b:
        return O, "", f"{a[0]}-{a[1]}"
    return X, f"엑셀 {a[0]}-{a[1]}  ↔  증빙 {b[0]}-{b[1]}", "학기 다름"


# ── 잔손 ──────────────────────────────────────────────────────
def _보기(d: str) -> str:
    return f"{d[:4]}-{d[4:6]}-{d[6:]}" if len(d) == 8 else f"{d[:4]}-{d[4:6]}"


def _보기목록(ds: list) -> str:
    return " ~ ".join(_보기(x) for x in sorted(ds))


def _짧게(검증열: str) -> str:
    t = re.sub(r"\s*\([^)]*\)\s*$", "", str(검증열 or "").strip())
    for 꼬리 in (" 제시 여부", " 포함 여부", " 제출여부", " 일치", " 여부", " 확인"):
        if t.endswith(꼬리):
            return t[: -len(꼬리)].strip() or 검증열
    return t or 검증열


# ══════════════════════════════════════════════════════════════
def 채우기(sp, 시트: str, 줄: dict, 엑셀값: dict) -> dict:
    """그 줄의 계산 항목을 전부 판정해 모형 판정 옆에 얹는다.

    줄["판정"] 은 모형이 낸 것 {검증열: {판정,근거,비고}}.  여기에 계산 몫을
    더해 돌려준다.  앞선 항목 종속은 모형 판정을 보고 가린다.
    """
    본 = dict(줄.get("판정") or {})
    번호별 = {c["#"]: c for c in sp.항목(시트)}
    앞선판정 = {n: (본.get(c["검증열"]) or {}).get("판정")
                for n, c in 번호별.items()}
    for c in sp.항목(시트):
        # '사람' 은 글자로 볼 수 없는 것이라 **언제나 확인 불가**다.
        # 모형에게 묻지도 않았으므로 여기서 채워야 칸이 비지 않는다.
        if c["주체"] == "사람":
            본[c["검증열"]] = {
                "판정": UNK,
                "비고": _짧게(c["검증열"]) + " — 눈으로 봐야 합니다",
                "근거": "코드가 둠 · 텍스트로는 볼 수 없는 항목",
            }
            continue
        if c["주체"] != "계산":
            continue
        v, 비고, 근거 = 재기(c, 엑셀값, 줄.get("값") or {},
                             맺은파일=줄.get("맺은파일"), 앞선판정=앞선판정)
        본[c["검증열"]] = {"판정": v, "비고": 비고, "근거": f"코드가 셈 · {근거}"}
    return 본
