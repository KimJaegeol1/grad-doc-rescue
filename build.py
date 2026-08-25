# -*- coding: utf-8 -*-
"""
배포본 만들기 — 보조도구
================================================================
    python build.py             열쇠를 함께 넣는다 (기본) — 받는 사람은 바로 실행
    python build.py --no-keys   열쇠 없이 — .env 를 따로 전달할 때

하는 일 (검수도구 build.py 와 같은 여섯 단계다)
    1. 빌드 전 점검 — 라이브러리·검수기준이 다 있나.  **없으면 여기서 멈춘다**
    2. PyInstaller 로 묶기
    3. 검수기준 엑셀·읽을거리를 제자리에
    4. secrets/ 를 다시 만든다        ← 손으로 넣은 것은 여기서 지워진다
    5. 열쇠가 새어 나갔는지 전수 검사  ← 걸리면 zip 을 안 만든다
    6. 배포용 zip                    ← **열쇠가 들어가도 만든다**

만들어지는 것
    dist/보조도구/                     폴더째 배포
    보조도구_배포_YYYYMMDD.zip          또는 …_열쇠포함.zip

★ 빌드는 배포 대상과 같은 운영체제에서 해야 한다.
  윈도우용 exe 는 윈도우에서만 만들어진다 (PyInstaller 는 교차 빌드를 못 한다).

★ 검수도구와 다른 점 하나
  검수도구는 설정 json 을 「프로그램파일/」 안에 넣지만, 보조도구는
  **검수기준 엑셀을 exe 바로 옆에 둔다.**  사람이 열어 고치는 물건이라
  라이브러리 수백 개가 든 폴더에 숨기면 아무도 못 찾는다 (paths.py 참고).
"""
from __future__ import annotations

import argparse
import datetime
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

WITH_KEYS = True                  # 기본은 넣는다.  빼려면 --no-keys

ROOT = Path(__file__).resolve().parent
NAME = "보조도구"
DIST = ROOT / "dist" / NAME
INNER = "프로그램파일"             # build.spec 의 contents_directory · paths.py 와 같아야 한다

# exe 옆에 둘 것 — 사람이 열어 보거나 고치는 물건
곁에 = [
    ("docs/보조도구 검수기준.xlsx", "보조도구 검수기준.xlsx"),
    ("엑셀_넣기_전에_읽어주세요.txt", "엑셀_넣기_전에_읽어주세요.txt"),
]
# 프로그램파일/ 안에 둘 읽을거리 — 있으면 넣고 없으면 넘어간다
안쪽읽을거리 = ["배포_안내.md", "배포본_점검표.txt"]

# 묶였는지 반드시 확인할 모듈 — 동적으로 부르는 것은 정적 분석이 놓친다
MUST = ["paths", "preflight"] + \
       [f"core.{m}" for m in ("spec", "book", "prompt", "ask", "calc",
                              "report", "flow")] + \
       [f"engines.{m}" for m in ("read", "ocr", "write")] + \
       [f"ui.{m}" for m in ("app", "theme", "runner", "widgets")]

# 빌드 PC 에 있어야 하는 것
PKGS = [("openpyxl", "openpyxl", "엑셀 읽기·쓰기", "전부"),
        ("openai", "openai", "모형 부르기", "검사 전체"),
        ("dotenv", "python-dotenv", "secrets/.env 읽기", "검사 전체"),
        ("lxml", "lxml", "docx·hwpx 파싱", "그 형식"),
        ("pypdf", "pypdf", "PDF 내장 글자", "pdf"),
        ("hwp5", "pyhwp", ".hwp 읽기", "hwp"),
        ("google.cloud.documentai_v1", '"google-cloud-documentai>=2.29.0"',
         "그림 PDF OCR", "그림 pdf"),
        ("tkinterdnd2", "tkinterdnd2", "끌어다 놓기", "편의")]
꼭 = {"openpyxl", "openai", "dotenv"}

# 배포본에 있으면 안 되는 것.  안내문의 예시(sk-... 처럼 점만 찍힌 것)는
# 걸리지 않도록 **진짜 열쇠 모양만** 본다.
열쇠꼴 = [
    (re.compile(r"sk-[A-Za-z0-9_\-]{20,}"), "OpenAI 열쇠"),
    (re.compile(r"AIza[0-9A-Za-z_\-]{30,}"), "Google API 열쇠"),
    (re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----"), "개인키"),
    (re.compile(r'"private_key"\s*:\s*"[^"]{40,}'), "서비스계정 개인키"),
]


def 말(s, mark=" "):
    print(f"  {mark} {s}")


def 머리(s):
    print("\n" + "═" * 66)
    print(f" {s}")
    print("═" * 66)


def _있나(m: str) -> bool:
    try:
        __import__(m)
        return True
    except Exception:                                    # noqa: BLE001
        return False


# ══════════════════════════════════════════════════════════════
def 점검() -> bool:
    """빠진 것은 **빌드 때** 잡는다.

    ★ 없어도 exe 는 만들어지고 그 기능만 조용히 죽는다.  실행할 때까지
      모른다.  그래서 여기서 멈춘다.
    """
    머리("1. 빌드 전 점검")
    ok = True

    for rel, _ in 곁에:
        있 = (ROOT / rel).exists()
        말(rel, "O" if 있 else "X")
        ok = ok and 있

    if _있나("PyInstaller"):
        말("PyInstaller 있음", "O")
    else:
        말("PyInstaller 가 없습니다 — pip install pyinstaller", "X")
        ok = False

    빠짐 = []
    for mod, 설치, 무엇, 없으면 in PKGS:
        있 = _있나(mod.split(".")[0]) if "." not in mod else _있나(mod)
        if 있:
            말(f"{설치:<38} {무엇}", "O")
        elif mod in 꼭 or 설치 in 꼭:
            말(f"{설치:<38} 없으면 {없으면} 못 씀 — pip install {설치}", "X")
            ok = False
        else:
            말(f"{설치:<38} 없음 → {없으면} 만 빠집니다", "!")
            빠짐.append(설치)

    if WITH_KEYS:
        src = ROOT / "secrets"
        env = src / ".env"
        js = sorted(src.glob("*.json")) if src.is_dir() else []
        if env.exists():
            말(".env 있음 — 배포본에 넣습니다", "O")
        else:
            말(f".env 가 없습니다 — {env}", "X")
            말("  열쇠 없이 만들려면:  python build.py --no-keys", "X")
            ok = False                 # ★ 모르고 열쇠 없이 배포하는 사고를 막는다
        if js:
            말("서비스계정 열쇠 " + " · ".join(j.name for j in js)
               + " — 배포본에 넣습니다", "O")
            말("  ⚠ 쓰지 않는 옛 열쇠가 섞여 있지 않은지 이 목록을 꼭 보세요", "!")
        else:
            말("서비스계정 json 없음 — 그림 PDF(OCR)는 못 읽습니다", "!")

    if not ok:
        print("\n  위를 먼저 갖춘 뒤 다시 실행하세요.")
    return ok


def 만들기() -> bool:
    머리("2. 묶기 (PyInstaller)")
    말("몇 분 걸립니다…")
    for d in (ROOT / "build", ROOT / "dist"):
        if d.exists():
            shutil.rmtree(d)
    r = subprocess.run([sys.executable, "-m", "PyInstaller", "build.spec",
                        "--noconfirm", "--clean", "--log-level", "WARN"],
                       cwd=ROOT)
    if r.returncode != 0:
        말("빌드가 멈췄습니다 — 위 메시지를 보세요.", "X")
        return False
    if not DIST.exists():
        말(f"산출 폴더가 없습니다 — {DIST}", "X")
        return False
    말("묶기 끝", "O")
    return True


def 곁에두기():
    머리("3. 검수기준을 exe 옆에")
    for rel, 낼이름 in 곁에:
        src = ROOT / rel
        if src.exists():
            shutil.copy2(src, DIST / 낼이름)
            말(낼이름, "O")
    inner = DIST / INNER
    inner.mkdir(parents=True, exist_ok=True)
    for rel in 안쪽읽을거리:
        if (ROOT / rel).exists():
            shutil.copy2(ROOT / rel, inner / Path(rel).name)
            말(f"{INNER}/{rel}", "O")
    print()
    말("검수기준 엑셀은 프로그램파일/ 안이 아니라 **exe 바로 옆**에 있습니다.")
    말("사람이 열어 고치는 물건이라 그렇습니다 — 고쳐도 재빌드가 필요 없습니다.")


def 열쇠두기():
    머리("4. secrets/")
    d = DIST / "secrets"
    if d.exists():
        shutil.rmtree(d)            # 딸려 들어온 것이 있으면 통째로 지운다
    d.mkdir(parents=True)

    if not WITH_KEYS:
        (d / "여기에_키를_넣으세요.txt").write_text(
            "secrets — 열쇠 두는 곳\n"
            "================================================================\n"
            "이 폴더는 비어 있는 것이 정상입니다.\n"
            "\n"
            "담당자에게 '.env' 라는 파일을 따로 받아 이 폴더에 그대로 넣으세요.\n"
            "넣은 뒤 프로그램을 껐다 다시 켜면 됩니다.\n"
            "\n"
            "  보조도구\\\n"
            "      보조도구.exe\n"
            "      보조도구 검수기준.xlsx\n"
            "      secrets\\\n"
            "          여기에_키를_넣으세요.txt\n"
            "          .env               <- 받은 파일을 여기에\n"
            "\n"
            "  .env 안에는 이렇게 들어 있습니다\n"
            "      OPENAI_API_KEY=sk-...\n"
            "      OPENAI_MODEL=gpt-5-mini\n"
            "\n"
            "그림으로만 된 PDF 를 읽어야 하면 Google Document AI 서비스계정\n"
            "JSON 도 함께 받아 같은 폴더에 넣습니다.  이름은 무엇이든 됩니다.\n"
            "\n"
            "★ 이 도구는 열쇠가 없으면 검사를 못 합니다.\n"
            "  받은 .env 는 남에게 주지 마세요.  이 폴더째로 복사해 주지도 마세요.\n",
            encoding="utf-8")
        말("빈 secrets/ 와 안내를 만들었습니다", "O")
        return

    src = ROOT / "secrets"
    넣음 = []
    if (src / ".env").exists():
        shutil.copy2(src / ".env", d / ".env")
        넣음.append(".env")
    for j in sorted(src.glob("*.json")):
        shutil.copy2(j, d / j.name)
        넣음.append(j.name)
    if not 넣음:
        말(f"넣을 열쇠가 없습니다 — {src} 를 확인하세요", "X")
        return
    말("열쇠를 배포본에 넣었습니다: " + " · ".join(넣음), "!")
    말("  받는 사람은 그대로 실행만 하면 됩니다.", "!")
    말("  이 폴더·zip 이 밖으로 나가면 열쇠도 함께 나갑니다.", "!")
    (d / "이_폴더는_건드리지_마세요.txt").write_text(
        "secrets — 열쇠가 들어 있습니다\n"
        "================================================================\n"
        "이 폴더에는 프로그램이 쓰는 열쇠가 이미 들어 있습니다.\n"
        "지우거나 옮기면 검사가 아예 안 됩니다.\n"
        "\n"
        "이 폴더를 남에게 복사해 주지 마세요.\n"
        "프로그램 폴더째로 다른 사람에게 전달하지도 마세요.\n"
        "\n"
        "  .env                    OpenAI 열쇠 — 검사 전체에 씁니다\n"
        "  *.json (있으면)          Google Document AI — 그림 PDF 를 읽을 때만\n",
        encoding="utf-8")


def 확인() -> bool:
    """엔진이 실제로 묶였는지 본다.

    빠져도 exe 는 만들어지고, **그 기능을 눌러야 비로소 죽는다.**
    묶인 파일은 압축돼 있어 내용으로는 알 수 없으므로 Analysis TOC 로 본다.
    """
    머리("4-2. 빠진 모듈이 없나")
    tocs = sorted((ROOT / "build").rglob("Analysis-*.toc"))
    if not tocs:
        말("빌드 기록(Analysis TOC)을 못 찾아 확인을 건너뜁니다", "!")
        return True
    글 = "\n".join(t.read_text(errors="ignore") for t in tocs)
    빠짐 = [m for m in MUST if f"'{m}'" not in 글]
    for m in 빠짐:
        말(f"{m} 이 안 묶였습니다 — build.spec 의 hiddenimports 에 넣으세요", "X")
    if not 빠짐:
        말(f"{len(MUST)}개 모두 들어 있습니다", "O")
    return not 빠짐


def 열쇠샜나() -> bool:
    머리("5. 열쇠가 새어 나갔는지 검사")
    나쁨, 훑음 = [], 0
    for p in DIST.rglob("*"):
        if not p.is_file():
            continue
        안 = (p.parent == DIST / "secrets")
        if WITH_KEYS and 안:
            continue                       # 일부러 넣은 것 — 넘어간다
        if p.name.lower() == ".env" or (안 and p.suffix.lower() == ".json"):
            나쁨.append((p, "이 이름은 secrets/ 밖에 있으면 안 됩니다"))
            continue
        if p.suffix.lower() not in (".json", ".txt", ".env", ".cfg", ".ini",
                                    ".md", ".py", ".yaml", ".yml"):
            continue
        if p.stat().st_size > 2_000_000:
            continue
        try:
            글 = p.read_text("utf-8", errors="ignore")
        except OSError:
            continue
        훑음 += 1
        for 꼴, 무엇 in 열쇠꼴:
            if 꼴.search(글):
                나쁨.append((p, 무엇))
                break
    if WITH_KEYS:
        말("secrets/ 안은 일부러 넣은 것이라 검사에서 뺐습니다", "!")
    말(f"{훑음}개 파일을 훑었습니다")
    if 나쁨:
        말("배포본에 열쇠로 보이는 것이 있습니다 — 배포하지 마세요:", "X")
        for p, 왜 in 나쁨:
            말(f"  {p.relative_to(DIST)}  ({왜})", "X")
        return False
    말("secrets/ 밖에는 열쇠가 없습니다", "O")
    return True


def 묶기() -> Path:
    """★ 열쇠가 들어가도 zip 을 만든다.

    예전엔 열쇠가 들면 zip 을 안 만들고 "폴더째 옮기세요" 라고 했다.  그런데
    윈도우 기본 압축은 한글 이름을 깨뜨린다 — 실제로 `코드_구조.md` 가
    `肄붾뱶_援ъ“.md` 로 깨진 적이 있다.  손으로 압축하게 두면 그 사고가 난다.
    그래서 여기서 UTF-8 이름으로 만들어 주고, **이름에 열쇠포함을 박아** 둔다.
    """
    머리("6. 배포용 zip")
    날 = datetime.datetime.now().strftime("%Y%m%d")
    낼 = ROOT / (f"{NAME}_배포_{날}_열쇠포함.zip" if WITH_KEYS
                 else f"{NAME}_배포_{날}.zip")
    if 낼.exists():
        낼.unlink()
    with zipfile.ZipFile(낼, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(DIST.rglob("*")):
            if not p.is_file():
                continue
            i = zipfile.ZipInfo(str(Path(NAME) / p.relative_to(DIST)).replace("\\", "/"))
            i.flag_bits |= 0x800                      # ★ 이름은 UTF-8
            i.compress_type = zipfile.ZIP_DEFLATED
            i.external_attr = (p.stat().st_mode & 0xFFFF) << 16   # 실행 권한 보존
            z.writestr(i, p.read_bytes())
    말(f"{낼.name}  ({낼.stat().st_size / 1048576:,.0f} MB)", "O")
    if WITH_KEYS:
        말("이름에 '열쇠포함' 이 박혀 있습니다 — 공유드라이브·메일에 올리지 마세요", "!")
    return 낼


# ══════════════════════════════════════════════════════════════
def main() -> int:
    global WITH_KEYS
    ap = argparse.ArgumentParser(
        description="배포본 만들기 — 보조도구",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예)
  python build.py             열쇠를 함께 넣는다 (기본) — 받는 사람은 바로 실행
  python build.py --no-keys   열쇠 없이 — .env 를 따로 전달할 때
""")
    ap.add_argument("--no-keys", action="store_true",
                    help="secrets/ 의 열쇠를 넣지 않는다 (기본은 넣음)")
    WITH_KEYS = not ap.parse_args().no_keys

    print("배포본 만들기 —", NAME)
    print("  열쇠:", "함께 넣음" if WITH_KEYS else "넣지 않음 (--no-keys)")

    if not 점검():
        return 1
    if not 만들기():
        return 1
    곁에두기()
    열쇠두기()
    if not 확인():
        return 1
    if not 열쇠샜나():
        return 1
    낼 = 묶기()

    머리("끝")
    말(f"배포 파일 : {낼.name}")
    말(f"폴더      : {DIST}")
    print()
    말("배포본 생김새")
    말(f"     {NAME}.exe")
    말( "     보조도구 검수기준.xlsx      ← 여기를 고치면 검사 규칙이 바뀝니다")
    말( "     엑셀_넣기_전에_읽어주세요.txt")
    말( "     secrets/                  열쇠")
    말(f"     {INNER}/             읽을거리 · 라이브러리 (건드리지 않음)")
    print()
    말("받는 사람이 할 일")
    말("  1. zip 을 풀어 아무 폴더에나 둔다 (설치 없음)")
    if WITH_KEYS:
        말("  2. 보조도구.exe 를 실행한다   ← 그게 전부입니다")
        print()
        말("이 zip 에는 열쇠가 들어 있습니다 — 받는 사람에게만 주세요.", "!")
        말("공유드라이브·메일에 그대로 올리지 마세요.", "!")
    else:
        말("  2. 따로 받은 .env 를 secrets/ 폴더에 넣는다")
        말("  3. 보조도구.exe 를 실행한다")
        print()
        말("이 도구는 열쇠가 없으면 검사를 못 합니다 — .env 를 꼭 함께 전달하세요.", "!")
    print()
    말("주의")
    말("  · 처음 실행 때 백신·SmartScreen 이 막을 수 있습니다 — 예외로 등록하세요")
    말("  · exe 하나만 빼내면 안 돌아갑니다.  폴더째 옮기세요")
    말("  · 산출물에는 학생 이름이 들어 있습니다.  공유드라이브에 올리지 마세요")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
