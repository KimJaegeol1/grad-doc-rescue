# -*- coding: utf-8 -*-
r"""
stage3_read — 파일 → 텍스트 (3단계 읽기 층)
================================================================
경계: "파일 → 텍스트"만 한다.  문서 종류 판별·필드 추출·판정은 하지 않는다
      (stage3_extract · stage3_compare 소관).  ocr_folder.py 의 경계 원칙을 그대로 승계.

네 갈래 경로 — 형식마다 최선의 수단을 쓴다.
  .hwp   → pyhwp XML 모델 직접 파싱      (무비용 · 원문 그대로 · 표 보존)
  .hwpx  → zip + OWPML XML 파싱          (무비용 · 원문 그대로 · 표 보존)
  .docx  → zip + WordprocessingML 파싱   (무비용 · 원문 그대로 · 표 보존)
  .pdf   → ① 내장 텍스트 레이어 추출
           ② 품질 검사(자모 분리·문자 밀도)
           ③ 자모 분리는 NFC 정규화로 복구 시도
           ④ 그래도 못 쓰면 Document AI OCR (stage3_ocr — stage_3 내장)

  ※ hwp5txt(pyhwp 기본 변환기)는 표 내용을 "<표>" 자리표시자로 버린다.
    강의계획서처럼 본문이 전부 표 안에 있는 문서는 전멸하므로 쓰지 않는다.
    대신 Hwp5File.xmlevents() 로 XML 을 받아 TableRow/TableCell 을 직접 걷는다.

캐시 — 내용 주소 방식(content-addressed)
  파일 sha256 을 이름에 넣어 저장한다.  같은 파일은 두 번 읽지 않고,
  증빙이 재제출되면 해시가 달라져 새 항목이 생기고 종전 항목이 남는다.
  → 검수정의서 제13조(판정 이력)의 기반.

설치:
  pip install lxml pyhwp pypdf
  pip install "google-cloud-documentai>=2.29.0"      # PDF OCR 폴백을 쓸 때만
  ※ 없어도 hwp·hwpx·docx·텍스트레이어 PDF 는 그대로 읽는다.

사용(엔진 직접):
  python stage3_read.py --manifest 인계표.json -o 텍스트캐시
  python stage3_read.py --folder  D:\제출자료      -o 텍스트캐시
  python stage3_read.py --manifest 인계표.json --dry-run     # 대상·경로만 확인
  python stage3_read.py --manifest 인계표.json --no-ocr      # 네이티브·텍스트레이어만
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import sys
import threading
import time
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

from lxml import etree

import logging
logging.getLogger("pypdf").setLevel(logging.ERROR)      # 손상 PDF 경고는 우리가 따로 알린다

# Windows 콘솔 한글 깨짐 방지
if sys.platform == "win32":
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8")
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════
# 설정
# ══════════════════════════════════════════════════════════════
NATIVE_EXT = {".hwp", ".hwpx", ".docx"}
PDF_EXT = {".pdf"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".gif", ".bmp"}
SUPPORTED = NATIVE_EXT | PDF_EXT | IMAGE_EXT

# PDF 텍스트 레이어 품질 기준 — 미달이면 OCR 로 넘긴다
MIN_CHARS_PER_PAGE = 40     # 페이지당 이보다 적으면 스캔본으로 본다
JAMO_MAX = 0.10             # 한글 중 자모(분리) 비율 상한
CELL_SEP = " | "            # 표 칸 구분자
PAGE_MARK = "\n===== [p.{page}] =====\n"    # PDF 전용 (네이티브는 페이지 개념 없음)

CANCEL = threading.Event()  # GUI [중지] 가 set()
_print_lock = threading.Lock()

# 0·1·2단계와 같은 표기 — err ❌ / warn ⚠ / ok ✓ / done ★ / head ■
_PREFIX = {"err": "❌ ", "warn": "⚠ ", "ok": "✓ ", "done": "★ ", "head": "■ ",
           "skip": "⏭ ", "info": ""}


def log(msg: str, level: str = "info"):
    """GUI 가 이 함수를 갈아끼워 로그를 가져간다."""
    print(_PREFIX.get(level, "") + str(msg), flush=True)


_emit_lock = threading.Lock()


def log_many(lines):
    """한 파일의 로그를 통째로 낸다 — 동시 처리 중에도 줄이 섞이지 않는다."""
    with _emit_lock:
        for msg, lv in lines:
            log(msg, lv)


# 사람이 읽는 이름 — 화면·요약에 그대로 쓴다
METHOD_KO = {
    "hwp":            "한글 문서(.hwp) 원문",
    "hwpx":           "한글 문서(.hwpx) 원문",
    "docx":           "워드 문서(.docx) 원문",
    "pdf-layer":      "PDF 내장 텍스트",
    "pdf-layer(NFC)": "PDF 내장 텍스트(자모 복구)",
    "pdf-layer(미달)": "PDF 내장 텍스트(품질 미달)",
    "ocr":            "Document AI OCR",
    "ocr(image)":     "Document AI OCR(이미지)",
    "이미지":          "이미지 — OCR 필요",
    "네이티브":        "원문 그대로 읽기",
    "PDF 레이어→OCR":  "PDF 내장 텍스트 → 필요 시 OCR",
    "OCR(이미지)":     "이미지 OCR",
}
STATUS_KO = {
    "ok": "성공", "skipped": "캐시 사용", "empty": "텍스트 없음",
    "ocr필요": "OCR 필요", "error": "실패", "cancelled": "중지",
    "dry-run": "미리보기", "unsupported": "미지원 형식",
}


def ko_method(m: str) -> str:
    return METHOD_KO.get(m, m or "-")


def ko_status(s: str) -> str:
    return STATUS_KO.get(s, s)


class Cancelled(Exception):
    """사용자 중지 — 진행 중인 파일만 중단, 이미 만든 캐시는 보존."""


def _why(e: Exception) -> str:
    """예외를 사람이 읽는 사유로 옮긴다."""
    n, m = type(e).__name__, str(e)
    table = [
        ("FileNotFoundError", "파일을 찾을 수 없습니다 (경로가 바뀌었거나 삭제됨)"),
        ("PermissionError", "파일이 다른 프로그램에서 열려 있거나 접근 권한이 없습니다"),
        ("BadZipFile", "문서가 손상되었습니다 (압축 구조를 열 수 없음)"),
        ("DefaultCredentialsError", "Document AI 인증 정보가 없습니다 — secrets 폴더를 확인하세요"),
        ("PermissionDenied", "Document AI 권한이 없습니다 — 서비스계정 역할·결제 계정을 확인하세요"),
        ("ResourceExhausted", "Document AI 할당량을 넘었습니다 — 초당 요청을 낮추고 다시 시도하세요"),
        ("InvalidArgument", "Document AI 가 파일을 받지 못했습니다 (형식·용량 확인)"),
        ("ModuleNotFoundError", "필요한 파이썬 패키지가 없습니다"),
    ]
    for key, ko in table:
        if key in n or key in m:
            return ko
    if "hwp5" in m.lower() or "ole" in m.lower():
        return "한글 문서를 해석하지 못했습니다 (구버전 .hwp 이거나 손상된 파일)"
    return f"{n}: {m[:140]}"


# ══════════════════════════════════════════════════════════════
# 유틸
# ══════════════════════════════════════════════════════════════
def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for blk in iter(lambda: f.read(1 << 20), b""):
            h.update(blk)
    return h.hexdigest()


def natkey(s: str):
    """section2.xml 이 section10.xml 보다 앞서도록."""
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)]


def jamo_ratio(s: str) -> float:
    """한글 중 자모(분리 상태) 비율.  HWP→PDF 변환본의 깨진 텍스트 레이어 탐지용."""
    syl = jam = 0
    for ch in s:
        if "\uAC00" <= ch <= "\uD7A3":          # 완성형 음절
            syl += 1
        elif "\u1100" <= ch <= "\u11FF" or "\u3130" <= ch <= "\u318F":
            jam += 1                            # 첫/중/종성 · 호환 자모
    return jam / (syl + jam) if (syl + jam) else 0.0


def squeeze(s: str) -> str:
    """빈 줄 3개 이상 → 2개.  줄 끝 공백 제거."""
    s = "\n".join(ln.rstrip() for ln in s.replace("\r\n", "\n").split("\n"))
    return re.sub(r"\n{3,}", "\n\n", s).strip()


def safe_name(s: str, n: int = 60) -> str:
    """캐시 파일명용 — 경로 구분자·따옴표 제거."""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s)
    return s[:n]


# ══════════════════════════════════════════════════════════════
# 네이티브 ① .hwp  — pyhwp XML 모델
# ══════════════════════════════════════════════════════════════
def _hwp_cell_text(el) -> str:
    """el 아래 <Text> 중 '더 안쪽 표'에 속하지 않은 것만 이어붙인다."""
    out = []
    for t in el.iter("Text"):
        nested = False
        for a in t.iterancestors():
            if a is el:
                break
            if a.tag == "TableRow":     # el 과 t 사이에 또 다른 표가 있다
                nested = True
                break
        if not nested and t.text:
            out.append(t.text)
    return "".join(out).strip()


def read_hwp(path: Path) -> str:
    from hwp5.xmlmodel import Hwp5File           # 지연 임포트 — 없어도 다른 형식은 동작

    buf = io.BytesIO()
    h = Hwp5File(str(path))
    try:
        h.xmlevents(embedbin=False).dump(buf)
    finally:
        try:
            h.close()
        except Exception:
            pass

    root = etree.fromstring(buf.getvalue())
    lines = []
    for el in root.iter():
        if el.tag == "TableRow":
            if any(a.tag == "TableRow" for a in el.iterancestors()):
                continue                          # 중첩 표의 안쪽 행은 바깥에서 처리
            cells = [_hwp_cell_text(tc) for tc in el if tc.tag == "TableCell"]
            if any(cells):
                lines.append(CELL_SEP.join(cells))
        elif el.tag == "Paragraph":
            if any(a.tag == "TableCell" for a in el.iterancestors()):
                continue                          # 표 안 문단은 위에서 처리
            s = _hwp_cell_text(el)
            if s:
                lines.append(s)
    return squeeze("\n".join(lines))


# ══════════════════════════════════════════════════════════════
# 네이티브 ② .hwpx  — OWPML (zip + XML)
# ══════════════════════════════════════════════════════════════
def _ln(el) -> str:
    """네임스페이스 판본 차이에 견디게 로컬명만 쓴다."""
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def _in_tbl(el, stop=None) -> bool:
    for a in el.iterancestors():
        if a is stop:
            return False
        if _ln(a) == "tbl":
            return True
    return False


def _hwpx_text(el) -> str:
    return "".join(t.text or "" for t in el.iter()
                   if _ln(t) == "t" and not _in_tbl(t, stop=el)).strip()


def _hwpx_rows(tbl) -> list[str]:
    out = []
    for tr in tbl.iter():
        if _ln(tr) != "tr":
            continue
        cells = []
        for tc in tr:
            if _ln(tc) == "tc":
                cells.append("".join(t.text or "" for t in tc.iter()
                                     if _ln(t) == "t").strip())
        if any(cells):
            out.append(CELL_SEP.join(cells))
    return out


def read_hwpx(path: Path) -> str:
    lines = []
    with zipfile.ZipFile(path) as z:
        secs = [n for n in z.namelist()
                if re.search(r"section\d+\.xml$", n, re.I) and "Contents" in n]
        if not secs:                                    # 판본에 따라 경로가 다를 수 있다
            secs = [n for n in z.namelist() if re.search(r"section\d+\.xml$", n, re.I)]
        for s in sorted(secs, key=natkey):
            root = etree.fromstring(z.read(s))
            for el in root.iter():
                n = _ln(el)
                if n == "p" and not _in_tbl(el):
                    t = _hwpx_text(el)
                    if t:
                        lines.append(t)
                elif n == "tbl" and not _in_tbl(el):
                    lines.extend(_hwpx_rows(el))
    return squeeze("\n".join(lines))


# ══════════════════════════════════════════════════════════════
# 네이티브 ③ .docx  — WordprocessingML (zip + XML)
# ══════════════════════════════════════════════════════════════
NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_para(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{NS_W}t"))


def _docx_block(el, out: list):
    """문단·표를 문서 순서대로.  표 안의 표도 재귀 처리."""
    tag = el.tag.replace(NS_W, "")
    if tag == "p":
        s = _docx_para(el)
        if s.strip():
            out.append(s)
    elif tag == "tbl":
        for tr in el.findall(f"{NS_W}tr"):
            cells = []
            for tc in tr.findall(f"{NS_W}tc"):
                sub = []
                for ch in tc:
                    _docx_block(ch, sub)
                cells.append(" ".join(x.strip() for x in sub).strip())
            if any(cells):
                out.append(CELL_SEP.join(cells))


def read_docx(path: Path) -> str:
    out = []
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(f"{NS_W}body")
    if body is not None:
        for el in body:
            _docx_block(el, out)
    return squeeze("\n".join(out))


# ══════════════════════════════════════════════════════════════
# PDF ① 내장 텍스트 레이어
# ══════════════════════════════════════════════════════════════
def read_pdf_layer(path: Path) -> tuple[list[str], dict]:
    """반환: (페이지별 텍스트, 품질 정보).  실패하면 ([], {...})."""
    from pypdf import PdfReader

    info = {"pages": 0, "chars": 0, "jamo": 0.0, "repaired": False, "usable": False,
            "reason": ""}
    try:
        rd = PdfReader(str(path))
        if rd.is_encrypted:
            try:
                rd.decrypt("")
            except Exception:
                pass
        pages = [(p.extract_text() or "") for p in rd.pages]
    except Exception as e:
        info["reason"] = ("PDF 를 열 수 없습니다 (손상되었거나 PDF 형식이 아님)"
                          if "PdfStream" in type(e).__name__ or "PdfRead" in type(e).__name__
                          else f"PDF 내장 텍스트를 읽지 못했습니다 ({type(e).__name__})")
        return [], info

    info["pages"] = len(pages)
    if not pages:
        info["reason"] = "페이지가 없습니다 (빈 PDF)"
        return [], info

    joined = "\n".join(pages)
    info["chars"] = len(joined.strip())
    info["jamo"] = round(jamo_ratio(joined), 3)

    # 자모 분리(HWP→PDF 변환본에서 흔함) → NFC 정규화로 복구 시도
    if info["jamo"] > JAMO_MAX:
        fixed = [unicodedata.normalize("NFC", p) for p in pages]
        r2 = jamo_ratio("\n".join(fixed))
        if r2 <= JAMO_MAX:
            pages, info["repaired"], info["jamo"] = fixed, True, round(r2, 3)
        else:
            info["reason"] = (f"글자가 자모로 흩어져 있습니다 (분리 비율 {info['jamo']}) "
                              f"— 자동 복구에 실패")
            return pages, info

    if info["chars"] < MIN_CHARS_PER_PAGE * info["pages"]:
        info["reason"] = (f"글자가 거의 없습니다 ({info['chars']}자 / {info['pages']}쪽) "
                          f"— 스캔한 그림 문서로 보입니다")
        return pages, info

    info["usable"] = True
    return pages, info


# ══════════════════════════════════════════════════════════════
# PDF ② Document AI OCR 폴백 — stage3_ocr (stage_3 내장 엔진)
# ══════════════════════════════════════════════════════════════
_ocr_client = None
_ocr_throttle = None
_ocr_lock = threading.Lock()


def ocr_available() -> tuple[bool, str]:
    """(사용 가능 여부, 사유) — 창이 미리 알려주기 위함."""
    try:
        from . import ocr as stage3_ocr
    except Exception as e:
        return False, f"stage3_ocr 임포트 실패: {type(e).__name__}: {e}"
    return stage3_ocr.available()


def _ocr_init(rps: float):
    """클라이언트·스로틀을 한 번만 만든다 (스레드 공용)."""
    global _ocr_client, _ocr_throttle
    if _ocr_client is not None:
        return
    with _ocr_lock:
        if _ocr_client is not None:
            return
        from . import ocr as stage3_ocr
        _ocr_throttle = stage3_ocr.Throttle(rps)
        _ocr_client = stage3_ocr.build_client()


def ocr_pdf(path: Path, rps: float = 2.0, warn=lambda m: None) -> tuple[list[str], int]:
    """반환: (페이지별 텍스트, API 요청 건수)."""
    from . import ocr as stage3_ocr
    _ocr_init(rps)
    try:
        return stage3_ocr.ocr_pdf(path, _ocr_client, _ocr_throttle,
                                  warn=warn, cancel=CANCEL)
    except KeyboardInterrupt:
        raise Cancelled()


def ocr_image(path: Path, rps: float = 2.0) -> tuple[list[str], int]:
    from . import ocr as stage3_ocr
    _ocr_init(rps)
    return stage3_ocr.ocr_image(path, _ocr_client, _ocr_throttle)


# ══════════════════════════════════════════════════════════════
# 캐시 — 내용 주소(sha256) 방식
# ══════════════════════════════════════════════════════════════
def cache_paths(cache_dir: Path, digest: str, src: Path) -> tuple[Path, Path]:
    stem = f"{digest[:12]}_{safe_name(src.stem)}"
    return cache_dir / f"{stem}.txt", cache_dir / f"{stem}.meta.json"


def load_cached(cache_dir: Path, digest: str, src: Path):
    txt_p, meta_p = cache_paths(cache_dir, digest, src)
    if not (txt_p.exists() and meta_p.exists()):
        return None
    try:
        meta = json.loads(meta_p.read_text("utf-8"))
        if meta.get("sha256") != digest:
            return None
        return meta
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════
# 파일 1개 → 텍스트
# ══════════════════════════════════════════════════════════════
def read_file(src: Path, cache_dir: Path, *, force=False, allow_ocr=True,
              rps=2.0, dry_run=False) -> dict:
    """반환 계약 — 형식과 무관하게 항상 같은 모양.
       {file, name, ext, sha256, method, status, chars, pages, requests,
        txt, meta, sec, error}
       method : hwp | hwpx | docx | pdf-layer | pdf-layer(NFC) | ocr | ocr(image)
       status : ok | empty | skipped | error | cancelled | dry-run | unsupported
    """
    t0 = time.time()
    rec = {"file": str(src), "name": src.name, "ext": src.suffix.lower(),
           "sha256": "", "method": "", "status": "", "chars": 0, "pages": 0,
           "requests": 0, "txt": "", "meta": "", "sec": 0.0, "error": ""}
    out: list[tuple[str, str]] = []           # 이 파일의 로그 — 끝에서 한 번에 낸다

    def say(msg, level="info"):
        out.append((str(msg), level))

    def done(status, **kw):
        rec.update(status=status, sec=round(time.time() - t0, 2), **kw)
        if out:
            log_many(out)
        return rec

    if CANCEL.is_set():
        return done("cancelled")
    if not src.exists():
        say(f"{src.name}", "err")
        say("     파일을 찾을 수 없습니다 — 증빙 루트가 맞는지 확인하세요")
        return done("error", error="파일을 찾을 수 없습니다")
    ext = rec["ext"]
    if ext not in SUPPORTED:
        say(f"{src.name}", "warn")
        say(f"     읽을 수 없는 형식입니다 ({ext})")
        return done("unsupported", error=f"읽을 수 없는 형식 {ext}")

    if dry_run:
        plan = ("네이티브" if ext in NATIVE_EXT else
                "PDF 레이어→OCR" if ext in PDF_EXT else "OCR(이미지)")
        return done("dry-run", method=plan)

    digest = sha256(src)
    rec["sha256"] = digest
    txt_p, meta_p = cache_paths(cache_dir, digest, src)
    rec["txt"], rec["meta"] = str(txt_p), str(meta_p)

    # ── 캐시 적중 ──────────────────────────────────────────
    if not force:
        m = load_cached(cache_dir, digest, src)
        if m:
            say(f"{src.name}", "skip")
            say(f"     이미 읽은 파일입니다 — 캐시를 씁니다 "
                f"({ko_method(m.get('method',''))} · {m.get('chars', 0):,}자)")
            return done("skipped", method=m.get("method", ""), chars=m.get("chars", 0),
                        pages=m.get("pages", 0))

    pages: list[str] = []
    quality: dict = {}
    try:
        # ── 네이티브 ───────────────────────────────────────
        if ext == ".hwp":
            pages, rec["method"] = [read_hwp(src)], "hwp"
        elif ext == ".hwpx":
            pages, rec["method"] = [read_hwpx(src)], "hwpx"
        elif ext == ".docx":
            pages, rec["method"] = [read_docx(src)], "docx"

        # ── PDF ────────────────────────────────────────────
        elif ext in PDF_EXT:
            pages, quality = read_pdf_layer(src)
            if quality.get("usable"):
                rec["method"] = "pdf-layer(NFC)" if quality.get("repaired") else "pdf-layer"
            elif allow_ocr:
                say(f"{src.name}", "warn")
                say(f"     PDF 안에 쓸 만한 텍스트가 없어 OCR 로 읽습니다 "
                    f"— {quality.get('reason', '')}")
                pages, rec["requests"] = ocr_pdf(
                    src, rps, warn=lambda m: say(f"     {m}"))
                rec["method"] = "ocr"
            else:
                # OCR 이 꺼져 있어 확정할 수 없다 → 캐시에 남기지 않는다.
                # (남기면 나중에 OCR 을 켜고 돌려도 캐시 적중으로 건너뛰게 된다)
                say(f"{src.name}", "warn")
                say(f"     OCR 이 필요합니다 — {quality.get('reason', '')}")
                return done("ocr필요", method="pdf-layer(미달)",
                            chars=len("".join(pages).strip()), pages=len(pages),
                            error=quality.get("reason", ""))

        # ── 이미지 ─────────────────────────────────────────
        else:
            if not allow_ocr:
                say(f"{src.name}", "warn")
                say("     그림 파일이라 OCR 로만 읽을 수 있습니다")
                return done("ocr필요", method="이미지", error="그림 파일 — OCR 필요")
            pages, rec["requests"] = ocr_image(src, rps)
            rec["method"] = "ocr(image)"

        # ── 조립 ───────────────────────────────────────────
        multi = len(pages) > 1
        body = []
        for i, t in enumerate(pages, 1):
            if multi:
                h = PAGE_MARK.format(page=i)
                body.append(h.lstrip("\n") if i == 1 else h)
            body.append(t if t.endswith("\n") else t + "\n")
        text = "".join(body).strip() + "\n"

        cache_dir.mkdir(parents=True, exist_ok=True)
        txt_p.write_text(text, encoding="utf-8", newline="\n")
        meta = {
            "source": str(src), "name": src.name, "ext": ext, "sha256": digest,
            "method": rec["method"], "pages": len(pages), "chars": len(text.strip()),
            "requests": rec["requests"], "quality": quality,
            "read_at": datetime.now().isoformat(timespec="seconds"),
            "engine": "stage3_read/1.0",
        }
        meta_p.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

        rec["chars"], rec["pages"] = len(text.strip()), len(pages)
        status = "empty" if not text.strip() else "ok"
        say(f"{src.name}", "warn" if status == "empty" else "ok")
        detail = f"     {ko_method(rec['method'])}"
        if len(pages) > 1:
            detail += f" · {len(pages)}쪽"
        detail += f" · {rec['chars']:,}자"
        if rec["requests"]:
            detail += f" · OCR 요청 {rec['requests']}건"
        say(detail)
        if status == "empty":
            say("     읽어낸 글자가 없습니다. 원본이 그림만 있는 문서인지 확인하세요.")
        return done(status)

    except Cancelled:
        return done("cancelled")
    except Exception as e:
        say(f"{src.name}", "err")
        say(f"     읽지 못했습니다 — {_why(e)}")
        return done("error", error=f"{type(e).__name__}: {str(e)[:300]}")


# ══════════════════════════════════════════════════════════════
# 대상 수집
# ══════════════════════════════════════════════════════════════
def from_manifest(manifest: Path, root: Path | None = None, *, skip_collision=True,
                  skip_sheets=("별첨1. 명단",)) -> tuple[list[Path], dict]:
    """3단계 인계표(JSON) → 개봉 대상 파일 목록.
    ※ 인계표의 파일 경로는 '증빙 루트 기준 상대경로'다 (인계표에 절대경로가 없음).
      root 를 주면 그 아래로 해석하고, 주지 않으면 인계표 파일이 있는 폴더를 루트로 본다.
    키충돌 행은 매칭 미확정이므로 기본 제외(검수정의서 제6조).
    폴더로 끝나는 경로(묶음 폴더)는 파일이 아니므로 제외."""
    manifest = Path(manifest)
    root = Path(root) if root else manifest.parent
    d = json.loads(manifest.read_text("utf-8"))
    files, seen = [], set()
    stat = {"시트": 0, "행": 0, "제외_키충돌": 0, "제외_시트": 0, "경로_폴더": 0,
            "없는파일": 0}
    for s in d.get("sheets", []):
        if s.get("sheet") in skip_sheets:
            stat["제외_시트"] += len(s.get("rows", []))
            continue
        if not s.get("rows"):
            continue
        stat["시트"] += 1
        for r in s["rows"]:
            if skip_collision and r.get("키충돌"):
                stat["제외_키충돌"] += 1
                continue
            stat["행"] += 1
            for sub in r.get("제출", []):
                for p in sub.get("파일", []):
                    if not p or p.endswith("/") or p.endswith("\\"):
                        stat["경로_폴더"] += 1
                        continue
                    q = root / Path(p.replace("\\", "/"))
                    k = str(q).lower()
                    if k not in seen:
                        seen.add(k)
                        if not q.exists():
                            stat["없는파일"] += 1
                        files.append(q)
    return files, stat


def manifest_rels(manifest: Path, n=8) -> list[str]:
    """인계표에서 표본 상대경로 n개를 뽑는다 (루트 추정용)."""
    d = json.loads(Path(manifest).read_text("utf-8"))
    out = []
    for s in d.get("sheets", []):
        for r in s.get("rows", []):
            for sub in r.get("제출", []):
                for p in sub.get("파일", []):
                    if p and not p.endswith(("/", "\\")):
                        out.append(p.replace("\\", "/"))
                        if len(out) >= n:
                            return out
    return out


def guess_root(manifest: Path, extra: list[Path] | None = None) -> Path | None:
    """인계표의 상대경로가 실제로 존재하는 증빙 루트를 추정한다.
    인계표에 절대경로가 없으므로, 후보 폴더를 훑어 표본 경로가 맞는 곳을 고른다."""
    manifest = Path(manifest)
    rels = manifest_rels(manifest)
    if not rels:
        return None

    cands: list[Path] = list(extra or [])
    base = manifest.parent
    cands.append(base)
    cands.extend(list(base.parents)[:4])          # 위로 4단계
    try:
        for d in sorted(base.iterdir()):          # 형제 폴더 한 단계
            if d.is_dir() and not d.name.startswith("."):
                cands.append(d)
    except Exception:
        pass

    best, best_hit = None, 0
    for c in cands:
        try:
            hit = sum(1 for r in rels if (c / r).exists())
        except Exception:
            continue
        if hit > best_hit:
            best, best_hit = c, hit
            if hit == len(rels):
                break
    return best if best_hit else None


def from_folder(root: Path, recursive=True) -> tuple[list[Path], list[Path]]:
    it = root.rglob("*") if recursive else root.glob("*")
    todo, skip = [], []
    for p in sorted(it):
        if not p.is_file() or p.name.startswith(("~$", ".")):
            continue
        (todo if p.suffix.lower() in SUPPORTED else skip).append(p)
    return todo, skip


# ══════════════════════════════════════════════════════════════
# 배치
# ══════════════════════════════════════════════════════════════
def read_many(paths, cache_dir: Path, *, workers=4, force=False, allow_ocr=True,
              rps=2.0, dry_run=False, on_done=None) -> list[dict]:
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    recs, idx_lock = [], threading.Lock()
    index = cache_dir / "_index.jsonl"

    with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
        futs = [ex.submit(read_file, Path(p), cache_dir, force=force,
                          allow_ocr=allow_ocr, rps=rps, dry_run=dry_run) for p in paths]
        for i, fu in enumerate(as_completed(futs), 1):
            r = fu.result()
            recs.append(r)
            if not dry_run and r["status"] in ("ok", "empty", "skipped"):
                with idx_lock, index.open("a", encoding="utf-8") as f:
                    f.write(json.dumps({
                        "source": r["file"], "sha256": r["sha256"], "txt": r["txt"],
                        "method": r["method"], "chars": r["chars"],
                        "at": datetime.now().isoformat(timespec="seconds"),
                    }, ensure_ascii=False) + "\n")
            if on_done:
                on_done(i, len(futs), r)
    return recs


def format_summary(recs: list[dict], sec: float, dry=False) -> list[tuple[str, str]]:
    """완료 요약을 (문구, 레벨) 목록으로.  CLI 와 창이 같은 문구를 쓴다."""
    s = summarize(recs)
    st, me = s["status"], s["method"]
    out: list[tuple[str, str]] = [("─" * 66, "info")]

    if dry:
        out.append((f"미리보기 — 열어야 할 파일 {len(recs)}개", "done"))
        for k, v in sorted(me.items(), key=lambda x: -x[1]):
            out.append((f"     {ko_method(k):<28} {v:>4}개", "info"))
        out.append(("     실제로 읽으려면 [추출 시작]을 누르세요.", "info"))
        return out

    out.append((f"텍스트 추출 완료 — {sec}초 걸렸습니다", "done"))
    order = ["ok", "skipped", "empty", "ocr필요", "error", "unsupported", "cancelled"]
    parts = [f"{ko_status(k)} {st[k]}" for k in order if st.get(k)]
    out.append(("     " + "  ·  ".join(parts), "info"))

    if me:
        out.append(("     읽은 방식", "info"))
        for k, v in sorted(me.items(), key=lambda x: -x[1]):
            out.append((f"        {ko_method(k):<28} {v:>4}개", "info"))

    ocr_files = sum(1 for r in recs if r["requests"])
    tail = f"     글자 {s['chars']:,}자"
    if s["requests"]:
        tail += f"  ·  OCR 요청 {s['requests']:,}건 (파일 {ocr_files}개)"
    else:
        tail += "  ·  OCR 요청 없음 (전부 원문·내장 텍스트로 읽었습니다)"
    out.append((tail, "info"))

    bad = [r for r in recs if r["status"] in ("error", "empty", "unsupported", "ocr필요")]
    if bad:
        out.append(("─" * 66, "info"))
        out.append((f"확인이 필요한 {len(bad)}건", "warn"))
        for r in bad[:40]:
            lv = "err" if r["status"] == "error" else "warn"
            out.append((f"[{ko_status(r['status'])}] {r['name']}", lv))
            if r["error"]:
                out.append((f"        {r['error']}", "info"))
        if len(bad) > 40:
            out.append((f"     … 외 {len(bad)-40}건 (캐시 폴더의 _index.jsonl 참조)", "info"))
    return out


def summarize(recs: list[dict]) -> dict:
    st, me = {}, {}
    for r in recs:
        st[r["status"]] = st.get(r["status"], 0) + 1
        if r["method"]:
            me[r["method"]] = me.get(r["method"], 0) + 1
    return {"status": st, "method": me,
            "chars": sum(r["chars"] for r in recs),
            "requests": sum(r["requests"] for r in recs)}


# ══════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════
def main() -> int:
    ap = argparse.ArgumentParser(description="3단계 읽기 층 — 파일 1개당 텍스트 1개 (캐시)")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--manifest", help="3단계 인계표 JSON")
    g.add_argument("--folder", help="폴더 (하위 포함)")
    ap.add_argument("--root", help="증빙 루트 폴더 — 인계표의 상대경로를 여기 기준으로 해석 "
                                   "(생략 시 인계표가 있는 폴더)")
    ap.add_argument("-o", "--out", default=str(Path.cwd() / "text_cache"), help="텍스트 캐시 폴더")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--rps", type=float, default=2.0, help="OCR 초당 요청 상한")
    ap.add_argument("--force", action="store_true", help="캐시 무시하고 다시 읽기")
    ap.add_argument("--no-ocr", dest="allow_ocr", action="store_false",
                    help="OCR 폴백 없이 네이티브·텍스트레이어만")
    ap.add_argument("--dry-run", action="store_true", help="읽지 않고 대상·경로만 확인")
    ap.add_argument("--include-collision", action="store_true",
                    help="인계표의 키충돌 행도 포함")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    cache = Path(args.out).expanduser().resolve()
    if args.manifest:
        mf = Path(args.manifest).expanduser().resolve()
        rt = Path(args.root).expanduser().resolve() if args.root else mf.parent
        paths, stat = from_manifest(mf, rt, skip_collision=not args.include_collision)
        src_desc = f"인계표 {mf.name}\n  루트   {rt}\n  집계   {stat}"
        skipped = []
        if stat["없는파일"]:
            print(f"  ⚠ 인계표 경로 {stat['없는파일']}개가 루트 아래에 없습니다 — "
                  f"--root 를 확인하세요")
    else:
        root = Path(args.folder).expanduser().resolve()
        if not root.is_dir():
            print(f"✖ 폴더가 아님: {root}")
            return 1
        paths, skipped = from_folder(root)
        src_desc = f"폴더 {root}"
    if args.limit:
        paths = paths[:args.limit]

    ok, why = ocr_available()
    print("═" * 70)
    print(f"  대상     {src_desc}")
    print(f"  파일     {len(paths)}개")
    print(f"  캐시     {cache}")
    print(f"  OCR      {'준비됨' if ok else '쓸 수 없음 — ' + why}"
          f"{'' if args.allow_ocr else '   (--no-ocr 로 끔)'}")
    print("═" * 70)
    if skipped:
        kinds = sorted({p.suffix.lower() or "(무확장자)" for p in skipped})
        print(f"  ⚠ 읽을 수 없는 형식 {len(skipped)}개는 뺐습니다: {', '.join(kinds)}")
    if not paths:
        print("  처리할 파일이 없습니다.")
        return 0

    t0 = time.time()
    recs = read_many(paths, cache, workers=args.workers, force=args.force,
                     allow_ocr=args.allow_ocr, rps=args.rps, dry_run=args.dry_run)
    for msg, lv in format_summary(recs, round(time.time() - t0, 1), args.dry_run):
        print(_PREFIX.get(lv, "") + msg)
    return 0


if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("이 파일은 엔진입니다. 창으로 쓰려면:  python stage3_main.py")
        print("CLI:  python stage3_read.py --manifest 인계표.json -o 텍스트캐시")
        sys.exit(1)
    sys.exit(main())
