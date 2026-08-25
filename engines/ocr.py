# -*- coding: utf-8 -*-
r"""
stage3_ocr — Document AI OCR 엔진 (stage_3 내장)
================================================================
ocr_folder.py 의 검증된 엔진을 stage_3 안으로 들여온 것.  외부 파일에 의존하지 않는다.
stage3_read 가 "PDF 텍스트 레이어를 못 믿겠다"고 판단한 파일만 여기로 온다.

경계: "PDF·이미지 → 페이지별 텍스트"만 한다.  캐시·분류·판정은 하지 않는다.

핵심 — 페이지·용량 한도 처리:
  · Document AI 온라인 처리는 요청 1건당 한도가 있다.
      기본 15페이지 / 20MB,  imageless 30페이지 / 20MB  ← 기본값
  · 한도를 넘는 PDF 는 pypdf 로 물리 분할 → 청크별 요청 → 텍스트 순차 연결.
      ① 페이지 수로 자르고 ② 조각이 여전히 20MB 를 넘으면 절반씩 재귀 분할
  · 한도 미만이면 분할하지 않고 원본 바이트 그대로 전송(재직렬화 훼손 회피).
  · 서버가 페이지 한도 오류를 돌려주면 그 파일만 한도를 절반으로 낮춰 재시도.

인증 (셋 중 하나):
  ① stage_3\secrets\*.json  에 서비스계정 키를 넣는다      ← 권장
  ② 환경변수 GOOGLE_APPLICATION_CREDENTIALS 에 키 경로
  ③ gcloud auth application-default login
  필요 권한: roles/documentai.apiUser

설치:
  pip install "google-cloud-documentai>=2.29.0" "pypdf>=4.0"
  ※ 이 패키지가 없어도 stage_3 의 hwp·hwpx·docx·텍스트레이어 PDF 경로는 그대로 동작한다.
"""

from __future__ import annotations

import io
import json
import os
import random
import sys
import threading
import time
from pathlib import Path

# ══════════════════════════════════════════════════════════════
# 설정 — 환경변수로 덮어쓸 수 있다
# ══════════════════════════════════════════════════════════════
PROJECT_ID = os.getenv("DOCAI_PROJECT", "grad-doc-verify")
LOCATION = os.getenv("DOCAI_LOCATION", "us")
PROCESSOR_ID = os.getenv("DOCAI_PROCESSOR", "b73e8cddad753efb")      # doc-ocr

IMAGELESS_MODE = True             # True → 페이지 한도 15 → 30 상향 (이미지 미반환)
PAGE_LIMIT_IMAGELESS = 30
PAGE_LIMIT_PLAIN = 15
MAX_DOC_BYTES = 19 * 1024 * 1024  # 콘솔 표기 20MB — 여유 1MB
LANGUAGE_HINTS = ["ko", "en"]
RPC_TIMEOUT = 600                 # 초 — 30p 스캔본은 수 분 걸릴 수 있다
MAX_RETRY = 5                     # 일시 오류(429/503/504) 재시도 횟수

MIME = {
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
}


class PageLimitError(Exception):
    """서버가 '페이지 한도 초과'를 돌려준 경우 — 그 파일만 한도를 낮춰 재시도."""


# ══════════════════════════════════════════════════════════════
# 지연 임포트 — google 패키지가 없어도 모듈 자체는 불러와진다
# ══════════════════════════════════════════════════════════════
_documentai = None
_gexc = None
_RETRY_EXC: tuple = ()
_load_lock = threading.Lock()


def _load():
    global _documentai, _gexc, _RETRY_EXC
    if _documentai is not None:
        return
    with _load_lock:
        if _documentai is not None:
            return
        from google.api_core import exceptions as gexc
        from google.cloud import documentai_v1 as documentai
        _gexc, _documentai = gexc, documentai
        _RETRY_EXC = tuple(e for e in (
            getattr(gexc, "ResourceExhausted", None),      # 429 할당량
            getattr(gexc, "TooManyRequests", None),
            getattr(gexc, "ServiceUnavailable", None),     # 503
            getattr(gexc, "DeadlineExceeded", None),       # 504
            getattr(gexc, "InternalServerError", None),    # 500
            getattr(gexc, "Aborted", None),
            getattr(gexc, "Unknown", None),
        ) if e is not None)


# ══════════════════════════════════════════════════════════════
# 인증
# ══════════════════════════════════════════════════════════════
def base_dir() -> Path:
    """secrets\\ 가 있는 폴더.

    ★ 이 한 줄이 「열쇠를 넣었는데 없다고 한다」 를 만들었다
    ──────────────────────────────────────────────────
    예전엔 이 파일이 secrets\\ 와 **같은 층**에 있어서 `__file__` 의 부모가
    곧 프로그램 뿌리였다.  이 엔진을 `engines\\` 폴더로 옮기면서 한 겹
    깊어졌고, 이 줄도 같이 깊어져 `engines\\secrets\\` 를 보게 됐다.

        ask (OpenAI)  →  grad-rescue\\secrets\\           ← 사람이 넣는 곳
        ocr (문서AI)  →  grad-rescue\\engines\\secrets\\   ← 아무것도 없는 곳

    OpenAI 열쇠는 먹고 Document AI 열쇠만 "없습니다" 가 떴다.  ④세미나에서
    그림 PDF 10개가 통째로 못 읽혔다.  exe 로 묶으면 frozen 가지를 타서
    멀쩡했으므로 **소스로 돌릴 때만** 걸렸다 — 그래서 늦게 드러났다.

    파일을 옮기는 것만으로 뜻이 바뀌는 경로였다.  이제 프로그램이 정한
    한 곳(paths.secrets)을 쓴다.  paths 를 못 불러오면 한 겹 위로 올라간다.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    try:
        import paths
        return Path(paths.secrets()).parent
    except Exception:                                    # noqa: BLE001
        return Path(__file__).resolve().parent.parent


def find_key() -> Path | None:
    """base_dir()/secrets/ 안의 서비스계정 JSON.  없으면 None → 환경변수·ADC 폴백."""
    d = base_dir() / "secrets"
    if not d.is_dir():
        return None
    for k in sorted(d.glob("*.json")):
        try:
            j = json.loads(k.read_text("utf-8"))
        except Exception:
            continue
        if j.get("type") == "service_account" and j.get("private_key"):
            return k
    return None


def available() -> tuple[bool, str]:
    """(사용 가능 여부, 사유) — 창이 미리 알려주기 위함."""
    try:
        import google.cloud.documentai_v1  # noqa: F401
    except Exception as e:
        return False, (f"google-cloud-documentai 미설치 ({type(e).__name__}) — "
                       f'pip install "google-cloud-documentai>=2.29.0"')
    try:
        import pypdf  # noqa: F401
    except Exception:
        return False, 'pypdf 미설치 — pip install "pypdf>=4.0"'
    if find_key() is None and not os.getenv("GOOGLE_APPLICATION_CREDENTIALS"):
        return False, f"서비스계정 키 없음 — {base_dir()}\\secrets\\ 에 JSON 을 넣으세요"
    return True, f"준비됨 (프로세서 {PROCESSOR_ID} · {LOCATION})"


def processor_name() -> str:
    return f"projects/{PROJECT_ID}/locations/{LOCATION}/processors/{PROCESSOR_ID}"


def build_client():
    _load()
    from google.api_core.client_options import ClientOptions
    opts = ClientOptions(api_endpoint=f"{LOCATION}-documentai.googleapis.com")
    key = find_key()
    if key:                                   # 환경변수·gcloud 불필요 — 파일만 두면 됨
        return _documentai.DocumentProcessorServiceClient.from_service_account_file(
            str(key), client_options=opts)
    return _documentai.DocumentProcessorServiceClient(client_options=opts)


# ══════════════════════════════════════════════════════════════
# 할당량 보호 — 요청 시작 간 최소 간격
# ══════════════════════════════════════════════════════════════
class Throttle:
    def __init__(self, rps: float):
        self._interval = (1.0 / rps) if rps > 0 else 0.0
        self._lock = threading.Lock()
        self._next = 0.0

    def wait(self):
        if self._interval <= 0:
            return
        with self._lock:
            now = time.monotonic()
            if now < self._next:
                time.sleep(self._next - now)
                now = time.monotonic()
            self._next = max(now, self._next) + self._interval


# ══════════════════════════════════════════════════════════════
# PDF 분할 — ① 페이지 한도로 자르고 ② 20MB 넘으면 절반씩 재귀
# ══════════════════════════════════════════════════════════════
def _slice_bytes(reader, start: int, end: int) -> bytes:
    from pypdf import PdfWriter
    w = PdfWriter()
    for i in range(start, end):
        w.add_page(reader.pages[i])
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue()


def _split_by_size(reader, start, end, max_bytes, warn):
    data = _slice_bytes(reader, start, end)
    if len(data) <= max_bytes or (end - start) <= 1:
        if len(data) > max_bytes:
            warn(f"p.{start+1} 단일 페이지가 {len(data)/1048576:.1f}MB — 한도 초과, 그대로 시도")
        return [(start, end, data)]
    mid = (start + end) // 2
    return (_split_by_size(reader, start, mid, max_bytes, warn)
            + _split_by_size(reader, mid, end, max_bytes, warn))


def plan_pdf(path: Path, page_limit: int, warn) -> tuple[int, list]:
    """반환: (총 페이지 수, [(시작0based, 끝exclusive, bytes|None), ...])
    분할이 불필요하면 bytes=None → 원본 파일을 그대로 전송한다."""
    from pypdf import PdfReader
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")            # 빈 암호 (열람 제한만 걸린 흔한 경우)
        except Exception:
            pass
    n = len(reader.pages)
    if n == 0:
        raise ValueError("페이지 0 — 손상되었거나 빈 PDF")

    if n <= page_limit and path.stat().st_size <= MAX_DOC_BYTES:
        return n, [(0, n, None)]          # 무분할 — 원본 바이트 사용

    chunks = []
    for s in range(0, n, page_limit):
        chunks.extend(_split_by_size(reader, s, min(s + page_limit, n),
                                     MAX_DOC_BYTES, warn))
    return n, chunks


# ══════════════════════════════════════════════════════════════
# 요청
# ══════════════════════════════════════════════════════════════
def _ocr_config(native_pdf: bool):
    # enable_native_pdf_parsing: 디지털 PDF 의 내장 텍스트 레이어를 그대로 사용(빠름).
    #   stage3_read 가 이미 레이어를 검사해 못 쓴다고 판단한 뒤에 오므로 기본은 끈다.
    return _documentai.OcrConfig(
        hints=_documentai.OcrConfig.Hints(language_hints=LANGUAGE_HINTS),
        enable_native_pdf_parsing=native_pdf,
        enable_image_quality_scores=False,
        enable_symbol=False,
        disable_character_boxes_detection=True,    # 응답 크기·시간 절감 (좌표 미사용)
    )


def _is_page_limit_error(e: Exception) -> bool:
    m = str(e).lower()
    return "page" in m and ("limit" in m or "exceed" in m or "too many" in m)


def process_once(client, name, content: bytes, mime: str, native_pdf: bool,
                 throttle: Throttle):
    """요청 1건 — 일시 오류는 지수 백오프 재시도, 페이지 한도 오류는 위로 올린다."""
    _load()
    req = _documentai.ProcessRequest(
        name=name,
        raw_document=_documentai.RawDocument(content=content, mime_type=mime),
        skip_human_review=True,
        imageless_mode=IMAGELESS_MODE,
        process_options=_documentai.ProcessOptions(ocr_config=_ocr_config(native_pdf)),
    )
    last = None
    for attempt in range(1, MAX_RETRY + 1):
        if throttle:
            throttle.wait()
        try:
            return client.process_document(request=req, timeout=RPC_TIMEOUT).document
        except _gexc.InvalidArgument as e:
            if _is_page_limit_error(e):
                raise PageLimitError(str(e)[:200])
            raise
        except _RETRY_EXC as e:
            last = e
            if attempt == MAX_RETRY:
                break
            time.sleep(min(2 ** attempt + random.random(), 60))
    raise last


def page_texts(doc) -> list[str]:
    """Document → 페이지별 텍스트.  text_anchor 오프셋으로 doc.text 를 잘라낸다."""
    full = doc.text or ""
    if not doc.pages:
        return [full]
    out = []
    for pg in doc.pages:
        segs = pg.layout.text_anchor.text_segments
        if not segs:
            out.append("")
            continue
        out.append("".join(full[int(s.start_index):int(s.end_index)] for s in segs))
    return out


# ══════════════════════════════════════════════════════════════
# 공개 API — stage3_read 가 부르는 두 함수
# ══════════════════════════════════════════════════════════════
def ocr_pdf(path: Path, client, throttle: Throttle, *, warn=lambda m: None,
            cancel=None) -> tuple[list[str], int]:
    """반환: (페이지별 텍스트, API 요청 건수).  페이지 한도 오류는 한도를 반감해 재시도."""
    name = processor_name()
    limit = PAGE_LIMIT_IMAGELESS if IMAGELESS_MODE else PAGE_LIMIT_PLAIN

    for _try in range(3):
        try:
            _total, chunks = plan_pdf(path, limit, warn)
            raw = path.read_bytes() if (len(chunks) == 1 and chunks[0][2] is None) else None
            texts = []
            for (s, e, blob) in chunks:
                if cancel is not None and cancel.is_set():
                    raise KeyboardInterrupt("중지")
                content = blob if blob is not None else raw
                if len(content) > MAX_DOC_BYTES:
                    warn(f"{len(content)/1048576:.1f}MB — 20MB 한도 초과 가능")
                doc = process_once(client, name, content, "application/pdf", False, throttle)
                texts.extend(page_texts(doc))
                if len(chunks) > 1:
                    warn(f"p.{s+1}-{e} 완료 ({len(chunks)}청크 중)")
            return texts, len(chunks)
        except PageLimitError:
            if limit <= 1:
                raise
            limit = max(1, limit // 2)
            warn(f"서버 페이지 한도 오류 → 청크 크기 {limit}p 로 낮춰 재시도")
    raise RuntimeError("페이지 한도 재시도 3회 실패")


def ocr_image(path: Path, client, throttle: Throttle) -> tuple[list[str], int]:
    mime = MIME.get(path.suffix.lower())
    if not mime:
        raise ValueError(f"OCR 미지원 확장자 {path.suffix}")
    doc = process_once(client, processor_name(), path.read_bytes(), mime, False, throttle)
    return page_texts(doc), 1


if __name__ == "__main__":
    ok, why = available()
    print(f"OCR 엔진  {'✔ ' if ok else '✖ '}{why}")
    print(f"  프로세서  {processor_name()}")
    print(f"  키        {find_key() or os.getenv('GOOGLE_APPLICATION_CREDENTIALS') or '없음'}")
    print(f"  한도      {PAGE_LIMIT_IMAGELESS if IMAGELESS_MODE else PAGE_LIMIT_PLAIN}p / "
          f"{MAX_DOC_BYTES//1048576}MB · imageless={IMAGELESS_MODE}")
