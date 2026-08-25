# -*- coding: utf-8 -*-
"""
core.ask — 모형에게 시트를 통째로 맡기고, 낸 답을 코드가 검사한다.
================================================================
두 가지 일을 한다.

    ① 부른다      core.prompt 가 만든 글 + 답 꼴(strict json_schema)
    ② 검사한다    모형이 뭐라 하든 **코드가 지키는 것**이 있다

★ 코드가 지키는 것
─────────────────
모형은 시트 하나를 통째로 보고 답한다.  잘하지만 가끔 어긋난다.  그때
그대로 검사 칸에 옮기면 학교에 잘못 나간다.  그래서 아래는 코드가 막는다.

    · 엑셀에 없는 줄번호를 냈다              →  버린다
    · 없는 파일 표를 냈다                    →  그것만 버린다 (이제 거의 못 온다)
    · 답을 빠뜨린 줄이 있다                  →  알린다 (그 줄은 아무것도 안 쓴다)
    · 근거 없이 O·X 를 냈다                  →  **확인 불가로 내린다**
    · 맺은 파일이 없는데 O·X 를 냈다          →  **확인 불가로 내린다**
    · 확신이 '낮음'·'못맺음' 인데 X 를 냈다    →  **확인 불가로 내린다**
    · 한 파일을 여러 줄이 가져갔다            →  **막지 않는다.  알리기만 한다**
    · 판정 출처에만 있고 맺은파일엔 없는 표    →  **주워 담는다.  알리기만 한다**

★ 파일은 표로 온다
─────────────────
core.prompt 가 파일마다 `file_01` 같은 표를 붙여 싣고, 규격(enum)이 그
표들만 받는다.  그래서 여기 오는 파일은 **반드시 있는 파일**이다.

옛날엔 이름을 자유 글자로 받았고, 모형이 파일을 제대로 찾아 놓고도 이름을
기억으로 다시 썼다 — 「01_…과학원.pdf」 를 「12_…과학원.pdf」 로.  없는
이름이라 여기서 전부 버렸고, 서울대 21줄이 모두 맺은파일 = [] 이 되어 판정
160개가 확인 불가로 내려갔다.  판정이 틀린 게 아니라 **이름을 옮겨 적는
통로가 막힌 것**이었다.

표를 실제 이름으로 되돌리는 일은 여기서 한다.  파일 목록을 아는 쪽이 여기다.

왜 마지막만 다른가
────────────────
옛 도구는 "한 파일은 한 줄에만" 을 규칙으로 못박고 겹치면 둘 다 물렀다.
파일명 번호도 없이 점수로만 맺던 시절의 안전장치다.

지금은 다르다.  모형이 엑셀 줄과 파일을 **한눈에 놓고** 고른다.  그리고
실제로 한 파일이 여러 줄의 증빙인 경우가 있다 — ⑧학술발표의 프로그램북
하나에 발표 여럿이 실리고, ④세미나 참석자 명단이 회차 여럿을 덮는다.
규칙으로 막으면 그런 줄이 통째로 미배정이 된다.  그래서 **알리고 넘긴다.**

왜 '근거 없으면 확인 불가' 인가
──────────────────────────────
통짜로 맡기면 왜 그 판정이 나왔는지 되짚을 길이 근거뿐이다.  근거를 못 적는
판정은 사람이 확인할 수 없는 판정이고, 확인할 수 없는 X 는 학교에 나가면
안 된다.  프롬프트에도 그렇게 적어 두었고, 여기서 실제로 지킨다.

설치:  pip install openai python-dotenv
"""

from __future__ import annotations

import json
import os
import re
import threading
from pathlib import Path

from . import prompt as P

CANCEL = threading.Event()          # 화면 [중지] 가 set()
확신순 = {"높음": 3, "보통": 2, "낮음": 1, "못맺음": 0}


class Cancelled(Exception):
    pass


def log(msg: str, level: str = "info"):
    print(msg)


def _tick():
    if CANCEL.is_set():
        raise Cancelled()


# ══════════════════════════════════════════════════════════════
# 열쇠
# ══════════════════════════════════════════════════════════════
def 열쇠읽기(자리: Path | None = None):
    """secrets/.env 를 읽는다.  이미 환경에 있으면 덮어쓰지 않는다."""
    자리 = Path(자리) if 자리 else None
    if 자리 is None:
        try:
            import paths
            자리 = Path(paths.secrets())
        except Exception:                                # noqa: BLE001
            자리 = Path(__file__).resolve().parent.parent / "secrets"
    env = 자리 / ".env"
    if not env.exists():
        return
    try:
        from dotenv import load_dotenv
        load_dotenv(env, override=False)
    except ImportError:
        for ln in env.read_text("utf-8").splitlines():
            if "=" in ln and not ln.strip().startswith("#"):
                k, v = ln.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip("\"'"))


def available() -> tuple:
    if not os.getenv("OPENAI_API_KEY"):
        return False, ("OPENAI_API_KEY 가 없습니다.\n"
                       "secrets/.env 에 넣어 주세요.  이 도구는 열쇠가 없으면 "
                       "할 수 있는 일이 없습니다.")
    try:
        import openai  # noqa: F401
    except ImportError:
        return False, "openai 미설치 — pip install openai"
    return True, ""


def _client():
    from openai import OpenAI
    return OpenAI()


def _model() -> str:
    return os.getenv("OPENAI_MODEL") or "gpt-5-mini"


# ══════════════════════════════════════════════════════════════
# 부르기
# ══════════════════════════════════════════════════════════════
def run(sp, 시트: str, 줄들: list, 명단: list, 파일들: list, *,
        model: str = "", 재시도: int = 2, prog=None, 묶음: int = 8) -> dict:
    """시트를 **묶음씩 나눠** 맡긴다.

    ★ 왜 나누나
      한 번에 22줄을 물었더니 모형이 3~8줄 여섯 개만 답하고 닫았다.  한 줄이
      700~900토큰이라 22줄이면 2만 토큰인데, 그 앞에서 손을 놓은 것이다.
      8줄이면 6~7천 토큰이라 끝까지 간다.

    ★ 앞부분은 배치마다 똑같다
      「이번에 답할 줄」 만 끝에서 바뀐다.  24만 토큰짜리 앞부분이 그대로라
      프롬프트 캐시가 먹는다.  안 그러면 값이 배치 수만큼 곱해진다.

    반환 {"줄": {행: {...}}, "안맺힌파일": [...], "신호": [...],
          "잰것": {...}, "자름": [...]}
    """
    ok, why = available()
    if not ok:
        raise RuntimeError(why)

    mdl = model or _model()
    행들 = [r["행"] for r in 줄들]
    # ★ 규격이 API 한도(enum 1,000)에 드는 크기로 줄인다.  넘치면 400 이 나고
    #   같은 규격으로 다시 불러 봐야 똑같이 튕긴다 — 부르기 전에 맞춘다.
    바라던 = max(1, int(묶음 or 8))
    묶음 = P.맞는묶음(sp, 시트, 파일들, 바라던)
    if 묶음 < 바라던:
        log(f"     한 번에 {바라던}줄이면 규격이 API 한도를 넘어 "
            f"{묶음}줄씩으로 줄였습니다 (파일 {len(파일들)}개)", "warn")
    배치들 = [행들[i:i + 묶음] for i in range(0, len(행들), 묶음)]

    맛 = P.build(sp, 시트, 줄들, 명단, 파일들, 물을줄=행들[:묶음])
    log(f"[{시트}]  모형에게 맡깁니다 — 엑셀 {len(줄들)}줄 · 파일 {len(파일들)}개 · "
        f"{len(배치들)}번에 나눠 (한 번에 {묶음}줄) · "
        f"어림 {맛['잰것']['어림토큰']:,}토큰 / 한도 {P._한도():,}", "head")
    if 맛["자름"]:
        전체 = sum(x["전체"] for x in 맛["자름"])
        실은 = sum(x["실은것"] for x in 맛["자름"])
        log(f"     ※ 파일 {len(맛['자름'])}개는 앞뒤만 실었습니다 — "
            f"그 파일들의 {전체:,}자 중 {실은:,}자 "
            f"(앞 {1 - P.꼬리몫:.0%} · 뒤 {P.꼬리몫:.0%}).  "
            f"가운데에만 있는 것은 모형이 못 봅니다", "warn")
        for x in 맛["자름"][:6]:
            log(f"          {x['파일'][:44]}  {x['전체']:,} → {x['실은것']:,}자", "skip")
        if len(맛["자름"]) > 6:
            log(f"          … 그 밖 {len(맛['자름']) - 6}개", "skip")

    본: dict = {}
    신호: list = []
    안맺힌: list = []
    임자: dict = {}                      # 표 → [행…]   앞선 배치가 가져간 파일
    이름표 = dict(zip(P.표들(파일들), P.이름들(파일들)))
    표찾기 = {v: k for k, v in 이름표.items()}
    들어감 = 나옴 = 다시 = 0
    끝난까닭: list = []
    for i, 물을줄 in enumerate(배치들, 1):
        _tick()
        if prog:
            prog(i - 1, len(배치들))
        log(f"     [{i}/{len(배치들)}] {물을줄[0]}~{물을줄[-1]}줄 "
            f"({len(물을줄)}개) 를 묻습니다"
            + (f" · 임자 있는 파일 {len(임자)}개는 알려 줍니다" if 임자 else ""), "skip")
        답, 쓴, 잰 = _한판(sp, 시트, 줄들, 명단, 파일들, 물을줄,
                           mdl=mdl, 재시도=재시도, 임자=임자)
        다시 += 잰.get("다시", 0)
        if 잰.get("끝", "") not in ("", "stop"):
            끝난까닭.append(f"{i}번째 '{잰['끝']}'")
        if 잰.get("빈답"):
            신호.append({"글": f"{물을줄[0]}~{물을줄[-1]}줄 {len(물을줄)}개는 "
                               f"다시 물어도 끝내 빈 답이었습니다 — 맺은 파일도 "
                               f"못 맺은 까닭도 없습니다.  사람이 봐 주세요",
                         "수준": "err"})
        조각, 신 = 검사(sp, 시트, 답, 줄들, 파일들, 물을줄=물을줄)
        본.update(조각)
        신호 += 신
        안맺힌 += _안맺힌풀기(답, 파일들)
        for 행, v in 조각.items():
            for 이름 in v["맺은파일"]:
                t = 표찾기.get(이름)
                if t:
                    임자.setdefault(t, []).append(행)
        if 쓴 is not None:
            들어감 += getattr(쓴, "prompt_tokens", 0) or 0
            나옴 += getattr(쓴, "completion_tokens", 0) or 0
    if prog:
        prog(len(배치들), len(배치들))

    신호 += _빌린증빙내리기(본, 줄들)
    신호 += _마무리검사(행들, 본, 파일들)
    잰것 = dict(맛["잰것"])
    잰것["엑셀줄"] = len(줄들)
    잰것["나눔"] = f"{len(배치들)}번 × {묶음}줄"
    잰것["다시부름"] = 다시
    if 끝난까닭:
        잰것["멈춘까닭"] = " · ".join(끝난까닭)
    잰것.pop("이번줄", None)
    if 들어감 or 나옴:
        잰것["쓴토큰_들어감"] = 들어감
        잰것["쓴토큰_나옴"] = 나옴
    # 같은 파일이 배치마다 올라오므로 한 번만 둔다
    본것, 안맺힌하나 = set(), []
    for x in 안맺힌:
        if x["파일"] in 본것:
            continue
        본것.add(x["파일"])
        안맺힌하나.append(x)
    return {"줄": 본, "안맺힌파일": 안맺힌하나,
            "신호": 신호, "잰것": 잰것, "자름": 맛["자름"], "쓴것": "모형"}


def _다시해도같나(e) -> bool:
    """다시 불러 봐야 결과가 같은 잘못인가 (규격·인증 같은 것)."""
    이름 = type(e).__name__
    글 = str(e)
    if 이름 in ("BadRequestError", "AuthenticationError", "PermissionDeniedError"):
        return True
    return "Invalid schema" in 글 or "invalid_request_error" in 글


def _빈답(나온: list) -> bool:
    """배치가 통째로 **어깨만 으쓱한** 답인가.

    ★ 실제로 있었던 일:  2번째 배치(11~18줄) 여덟 줄이 모두
      맺은파일 없음 · 값 전부 빈 글자 · 근거 없음 · 안맺힌파일에도 말 없음
      으로 왔다.  재시도도 아니었다 — **한 번에 저 답을 냈다.**

      그 여덟 줄의 서류는 폴더에 다 있었다.  1번 배치가 오히려 그것들을
      "2025 자료 — 다른 엑셀 행 대상" 이라고 또박또박 알아봤다.

      줄을 칸으로 박아 **줄은** 빠질 수 없게 했더니, 이번엔 그 칸을 **빈
      내용으로** 채웠다.  `맺은파일: []` 도 `이름: []` 도 규격에 맞기 때문이다.
      빈 답이 통하는 자리가 있으면 언젠가 빈 답이 온다 — 네 번째다.

      길이는 규격으로 못 박는다(strict 가 minItems 를 안 받는다).  그러니
      **코드가 답이 아니라고 판정하고 다시 묻는다.**

    까닭을 댄 것은 답으로 친다 — 21~24줄처럼 "엑셀에 참여학생명이 없어
    특정 불가" 는 어깨 으쓱이 아니라 제대로 된 답이다.
    """
    if not 나온:
        return True
    for x in 나온:
        if x.get("맺은파일"):
            return False
        if str(x.get("맺은근거") or "").strip():
            return False
    return True


def _한판(sp, 시트, 줄들, 명단, 파일들, 물을줄, *, mdl, 재시도, 임자=None) -> tuple:
    """배치 하나를 묻는다.  반환 (답, usage, {"다시", "끝"}).

    다시 부르는 자리 둘 — 둘 다 **조용히 넘기지 않는다.**
        · 요청한 줄이 빠졌다        (규격이 막지만 그래도 본다)
        · 배치가 통째로 빈 답이다    (_빈답)
    """
    sch = P.schema(sp, 시트, 파일들, 물을줄)   # ★ 파일 표·줄 칸을 규격에 박는다
    마지막탈, 되짚기, 잰 = "", "", {"다시": 0, "끝": "", "빈답": False}
    for 번째 in range(1, 재시도 + 2):
        _tick()
        글 = P.build(sp, 시트, 줄들, 명단, 파일들,
                     물을줄=물을줄, 임자=임자, 되짚기=되짚기)
        try:
            r = _client().chat.completions.create(
                model=mdl,
                messages=[{"role": "system", "content": 글["system"]},
                          {"role": "user", "content": 글["user"]}],
                response_format={"type": "json_schema", "json_schema": {
                    # ★ 이름은 ASCII 여야 한다 — API 가 ^[a-zA-Z0-9_-]+$ 만 받는다.
                    #   「검사」 라고 두었다가 400 으로 세 번 다 튕겼다.
                    #   스키마 **안**의 칸 이름은 한글이어도 된다.  여기만 다르다.
                    "name": "grad_check", "strict": True, "schema": sch}},
            )
            답 = json.loads(r.choices[0].message.content)
            잰["끝"] = getattr(r.choices[0], "finish_reason", "") or ""
            나온, 빠짐 = P.줄풀기(답, 물을줄)
            if 빠짐:
                # 규격이 required 로 막고 있으니 여기 오는 일은 드물다.  와도
                # **그냥 넘기지 않는다** — 조용히 빠진 줄이 이 도구의 지난
                # 실패였다.  다시 부른다.
                raise ValueError(
                    f"요청한 줄 {len(물을줄)}개 중 {len(빠짐)}개가 빠졌습니다 "
                    f"({', '.join(map(str, 빠짐[:6]))})"
                    + (f" · 모형이 멈춘 까닭 '{잰['끝']}'" if 잰["끝"] not in ("", "stop")
                       else ""))
            잰["빈답"] = _빈답(나온)
            if 잰["빈답"] and 번째 <= 재시도:
                되짚기 = (
                    f"지난번에 {물을줄[0]}~{물을줄[-1]}줄 {len(물을줄)}개를 "
                    f"**하나도 못 맺고** 값도 근거도 비운 채 냈다.\n"
                    "그 줄들의 서류는 위에 실려 있다.  엑셀 줄의 값과 서류 안의 "
                    "값으로 다시 찾아라.  **글자가 똑같기를 기다리지 마라** — "
                    "학교가 오타·띄어쓰기·㈜ 표기·줄임말로 다르게 적는 일이 흔하다.\n"
                    "정말 못 찾겠으면 **줄마다 왜 못 찾았는지 맺은근거에 적어라.** "
                    "빈 채로 두면 사람이 무엇을 봐야 할지 알 수 없다.")
                raise ValueError(
                    f"{물을줄[0]}~{물을줄[-1]}줄 {len(물을줄)}개가 통째로 빈 답입니다 "
                    f"— 맺은 파일도 까닭도 없습니다")
            return 답, getattr(r, "usage", None), 잰
        except Cancelled:
            raise
        except Exception as e:                           # noqa: BLE001
            마지막탈 = f"{type(e).__name__}: {str(e)[:160]}"
            잰["다시"] += 1
            # ★ 규격이 틀렸다는 400 은 **다시 불러도 똑같다.**  ④세미나에서
            #   같은 400 을 세 번 던지고 죽었다.  헛돈만 쓰고 화면만 시끄럽다.
            if _다시해도같나(e):
                raise RuntimeError(
                    f"규격이 잘못돼 모형이 받지 않습니다 — 다시 불러도 같습니다.\n"
                    f"{마지막탈}")
            log(f"     {번째}번째 호출이 쓸모없었습니다 — {마지막탈}",
                "warn" if 번째 <= 재시도 else "err")
            if 번째 > 재시도:
                raise RuntimeError(
                    f"모형을 {재시도 + 1}번 불렀으나 모두 실패했습니다.\n{마지막탈}")
    return {}, None, 잰                                  # 여기 오지 않는다


def _빌린증빙내리기(본: dict, 줄들: list) -> list:
    """제 값이 이웃 줄에 통째로 들어 있는 줄은 O·X 를 못 낸다 (개요 제0조 ⑤).

    ★ 실제로 있었던 일
    ─────────────────
        r19  ㈜청명기연환경  26-02-09~13  인원 7  학생 최성일 외 6명
        r21  ㈜청명기연환경  일시 없음    인원 없음  학생 없음
        둘 다 → 01_202602_청명기연환경.pdf

    21줄은 학교가 아직 안 채운 줄인데, 같은 회사라는 이유로 19줄 서류를
    빌려 맺고 「계획서 제출여부 = X」 를 받았다.  **학교에 반송이 나가는
    판정**이다.  학교는 그 줄을 쓰지도 않았다.

    한 파일을 여러 줄이 쓰는 것은 일부러 허용한 것이다 (⑧학술발표 프로그램북).
    다만 그 대가로 이런 자리가 생겼다.  막는 자리는 여기다 —
    **제 값으로 서류를 지목하지 못한 줄은 확인 불가로 둔다.**

    O 도 같이 내린다.  그 줄에 대해 확인된 것이 아무것도 없기 때문이다.
    """
    값 = {r["행"]: {k: str(v).strip() for k, v in (r["값"] or {}).items()
                    if str(v).strip()} for r in 줄들}
    임자: dict = {}
    for 행, v in 본.items():
        for f in v["맺은파일"]:
            임자.setdefault(f, []).append(행)

    신호, 내린줄 = [], []
    for 행, v in 본.items():
        파일 = v["맺은파일"]
        if not 파일:
            continue
        # 이 줄이 쓴 파일을 **다 같이 쓰는** 이웃 줄 가운데, 값이 더 많은 줄
        이웃 = [n for n in set(sum((임자.get(f) or [] for f in 파일), []))
                if n != 행]
        더큰 = [n for n in 이웃
                if set(값.get(행, {})) < set(값.get(n, {}))
                and all(값[행][k] == 값[n].get(k) for k in 값.get(행, {}))]
        if not 더큰:
            continue
        고친 = False
        for 이름, x in (v["판정"] or {}).items():
            if (x.get("판정") or "") in ("O", "X"):
                x["판정"] = "확인 불가"
                x["비고"] = (f"{_짧게(이름)} — 이 줄은 제 값이 "
                             f"{더큰[0]}줄에 다 들어 있어, 같은 서류로는 "
                             f"이 줄만의 것을 가릴 수 없습니다")
                x["근거"] = (x.get("근거") or "") + " · 코드가 내림(빌린 증빙)"
                고친 = True
        if 고친:
            내린줄.append((행, 더큰[0]))
    for 행, n in 내린줄[:6]:
        신호.append({"글": f"{행}줄은 제 값이 {n}줄에 다 들어 있고 서류도 같아 "
                           f"판정을 확인 불가로 내렸습니다 — 학교가 아직 안 "
                           f"채운 줄일 수 있습니다", "수준": "warn"})
    return 신호


def _마무리검사(행들: list, 본: dict, 파일들: list) -> list:
    """배치를 다 돌고 나서 한 번만 보는 것 — 빠진 줄 · 안 쓰인 파일."""
    신호 = []
    빠짐 = [n for n in 행들 if n not in 본]
    if 빠짐:
        신호.append({"글": f"끝내 답을 못 받은 줄이 {len(빠짐)}개 있습니다 "
                           f"({', '.join(map(str, 빠짐[:8]))}"
                           f"{' …' if len(빠짐) > 8 else ''}).  "
                           f"그 줄은 아무것도 쓰지 않습니다", "수준": "err"})
    쓴 = {f for v in 본.values() for f in v["맺은파일"]}
    안쓴 = sorted(set(P.이름들(파일들)) - 쓴)
    if 안쓴:
        신호.append({"글": f"어느 줄에도 안 쓰인 파일이 {len(안쓴)}개 있습니다 — "
                           f"엑셀에 없는 실적이거나 다른 시트 증빙일 수 있습니다",
                     "수준": "info"})
    return 신호


def _안맺힌풀기(답: dict, 파일들: list) -> list:
    """'안맺힌파일' 의 표도 실제 이름으로 되돌린다."""
    이름 = dict(zip(P.표들(파일들), P.이름들(파일들)))
    out = []
    for x in (답.get("안맺힌파일") or []):
        t = str((x or {}).get("파일") or "")
        out.append({"파일": 이름.get(t, t), "왜": str((x or {}).get("왜") or "")})
    return out


# ══════════════════════════════════════════════════════════════
# 검사 — 여기가 안전장치다
# ══════════════════════════════════════════════════════════════
def 검사(sp, 시트: str, 답: dict, 줄들: list, 파일들: list, *, 물을줄=None) -> tuple:
    """모형이 낸 것을 코드가 훑는다.  반환 (정리된 줄, 신호).

    파일은 표(file_01)로 온다.  여기서 실제 이름으로 되돌린다.
    줄은 `r3`·`r4` 칸으로 온다.  옛 꼴(배열 · 파일 이름 그대로)도 받아 준다.

    물을줄 을 주면 **그 배치의 몫만** 본다 — 빠진 줄과 안 쓰인 파일은 배치를
    다 돌고 나서 _마무리검사 가 한 번에 본다.
    """
    있는줄 = {r["행"] for r in 줄들}
    한판 = 물을줄 is not None
    답 = dict(답 or {})
    답["줄"], _빠짐 = P.줄풀기(답, 물을줄 or sorted(있는줄))
    이름표 = dict(zip(P.표들(파일들), P.이름들(파일들)))
    있는파일 = set(이름표.values())

    def 되돌리기(t) -> str:
        """표면 이름으로, 이미 이름이면 그대로, 아무것도 아니면 빈 글자."""
        t = str(t or "").strip()
        if t in 이름표:
            return 이름표[t]
        return t if t in 있는파일 else ""

    신호: list = []
    본: dict = {}
    파일임자: dict = {}

    for x in (답.get("줄") or []):
        행 = x.get("행")
        if 행 not in 있는줄:
            신호.append({"글": f"모형이 엑셀에 없는 줄번호 {행} 을 냈습니다 — 버립니다",
                         "수준": "warn"})
            continue
        if 행 in 본:
            신호.append({"글": f"{행}줄이 두 번 나왔습니다 — 앞엣것만 씁니다",
                         "수준": "warn"})
            continue

        # ── 맺은파일: 표 → 이름.  겹치는 것은 차례를 지키며 추린다 ──
        #    (strict 규격이 uniqueItems 를 안 받아 같은 표가 두 번 올 수 있다)
        파일, 지어냄 = [], []
        for t in (x.get("맺은파일") or []):
            n = 되돌리기(t)
            if not n:
                지어냄.append(str(t))
            elif n not in 파일:
                파일.append(n)
        if 지어냄:
            신호.append({"글": f"{행}줄 — 없는 파일 표 {', '.join(지어냄[:2])} 을 "
                               f"냈습니다.  그것만 버립니다", "수준": "warn"})

        판정 = P.판정풀기(sp, 시트, x)

        # ── 판정 출처를 이름으로 되돌리고, 빠진 파일은 주워 담는다 ──
        #    ★ 줄에서는 빠뜨렸는데 판정에서 그 서류를 봤다고 한 경우가 있다.
        #      그걸 버리면 '증빙 없는 O' 가 되어 애먼 확인 불가가 된다.
        주움 = []
        for v in 판정.values():
            n = 되돌리기(v.get("출처"))
            v["출처"] = n
            if n and n not in 파일:
                파일.append(n)
                if n not in 주움:
                    주움.append(n)
        if 주움:
            신호.append({"글": f"{행}줄 — 판정 근거로 든 「{주움[0]}」"
                               f"{f' 외 {len(주움) - 1}개' if len(주움) > 1 else ''} 가 "
                               f"맺은파일에 빠져 있어 채워 넣었습니다", "수준": "info"})

        확신 = str(x.get("확신") or "못맺음")
        for f in 파일:
            파일임자.setdefault(f, []).append(행)

        판정, 내린것 = _내리기(판정, 파일, 확신)
        for 이름, 왜 in 내린것:
            신호.append({"글": f"{행}줄 「{이름}」 — {왜}", "수준": "info"})
        _출처붙이기(판정)

        값 = dict(x.get("값") or {})
        본[행] = {
            "맺은파일": 파일, "맺은근거": str(x.get("맺은근거") or ""),
            "확신": 확신, "값": 값, "판정": 판정,
            "메모": str(x.get("메모") or ""),
        }

    # ── 빠진 줄 ───────────────────────────────────────────
    #    배치로 물을 때는 _한판 이 다시 부르고 _마무리검사 가 한 번에 알린다
    빠짐 = sorted(있는줄 - set(본))
    if 빠짐 and not 한판:
        신호.append({"글": f"모형이 답하지 않은 줄이 {len(빠짐)}개 있습니다 "
                           f"({', '.join(map(str, 빠짐[:8]))}"
                           f"{' …' if len(빠짐) > 8 else ''}).  "
                           f"그 줄은 아무것도 쓰지 않습니다", "수준": "err"})

    # ── 한 파일을 여러 줄이 가져감 — 막지 않고 알린다 ──────
    겹침 = {f: rs for f, rs in 파일임자.items() if len(rs) > 1}
    for f, rs in list(겹침.items())[:6]:
        신호.append({"글": f"「{f}」 를 {len(rs)}줄이 함께 씁니다 "
                           f"({', '.join(map(str, rs[:5]))}) — 맞는지 봐 주세요",
                     "수준": "info"})

    # ── 안 맺힌 파일 ──────────────────────────────────────
    안쓴 = sorted(있는파일 - set(파일임자))
    if 안쓴 and not 한판:
        신호.append({"글": f"어느 줄에도 안 쓰인 파일이 {len(안쓴)}개 있습니다 — "
                           f"엑셀에 없는 실적이거나 다른 시트 증빙일 수 있습니다",
                     "수준": "info"})
    return 본, 신호


def _내리기(판정: dict, 파일: list, 확신: str) -> tuple:
    """근거 없는 판정 · 증빙 없는 판정 · 헐거운 맺음의 X 를 확인 불가로 내린다."""
    낸것 = []
    for 이름, v in list(판정.items()):
        값 = v.get("판정") or ""
        if 값 not in ("O", "X"):
            continue                                   # 확인 불가·공란은 그대로
        근거 = str(v.get("근거") or "").strip()
        왜 = ""
        if not 근거:
            왜 = f"근거 없이 {값} 를 내어 확인 불가로 내렸습니다"
        elif not 파일:
            왜 = f"맺은 파일이 없는데 {값} 를 내어 확인 불가로 내렸습니다"
        elif 값 == "X" and 확신순.get(확신, 0) <= 1:
            왜 = (f"맺음 확신이 '{확신}' 인데 X 를 내어 확인 불가로 내렸습니다 — "
                  f"엉뚱한 파일이면 억울한 반송이 됩니다")
        if 왜:
            v["판정"] = "확인 불가"
            v["비고"] = (v.get("비고") or "").strip() or _짧게(이름) + " 을(를) 가리지 못함"
            v["근거"] = (근거 + " · " if 근거 else "") + "코드가 내림"
            낸것.append((이름, 왜))
    return 판정, 낸것


def _출처붙이기(판정: dict):
    """근거 앞에 어느 서류였는지를 적는다 — 「파일」 근거.

    ★ 반드시 _내리기 **뒤에** 부른다.  근거가 빈 판정을 확인 불가로 내리는
      안전장치가 앞에 있는데, 여기서 먼저 파일 이름을 채워 넣으면 '근거가
      있다' 로 보여 그 장치가 헛돈다.
    """
    for v in 판정.values():
        n = str(v.get("출처") or "").strip()
        if n and not str(v.get("근거") or "").startswith("「"):
            v["근거"] = f"「{n}」 {v.get('근거') or ''}".strip()


def _짧게(검증열: str) -> str:
    t = re.sub(r"\s*\([^)]*\)\s*$", "", str(검증열 or "").strip())
    for 꼬리 in (" 제시 여부", " 포함 여부", " 제출여부", " 일치", " 여부", " 확인"):
        if t.endswith(꼬리):
            return t[: -len(꼬리)].strip() or 검증열
    return t or 검증열


def format_summary(res: dict) -> list:
    c = res["잰것"]
    줄 = res["줄"]
    맺 = sum(1 for v in 줄.values() if v["맺은파일"])
    센다: dict = {}
    for v in 줄.values():
        for x in v["판정"].values():
            k = x["판정"] or "공란"
            센다[k] = 센다.get(k, 0) + 1
    out = [(f"엑셀 {c['엑셀줄']}줄 · 파일 {c['파일']}개 → 답한 줄 {len(줄)}개 "
            f"(파일을 맺은 줄 {맺}개)", "head"),
           ("     판정 " + " · ".join(f"{k} {v}" for k, v in sorted(센다.items())),
            "ok")]
    if "쓴토큰_들어감" in c:
        out.append((f"     토큰 들어감 {c['쓴토큰_들어감']:,} · "
                    f"나옴 {c['쓴토큰_나옴']:,}", "skip"))
    for s in res["신호"][:10]:
        out.append(("     " + s["글"], s["수준"]))
    return out
