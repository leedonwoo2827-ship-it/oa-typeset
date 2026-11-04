"""입력 문서 텍스트 추출 — docx / hwp / hwpx / xlsx / pptx / pdf / txt / md.

각 추출기는 의존성 미설치/파싱 실패에도 죽지 않고
"[추출 실패: ...]" 문자열을 돌려준다. (사용자가 더미 파일로 시험하므로 견고성 우선)
"""
from __future__ import annotations

import os
import zipfile
from typing import Callable, Dict

MAX_CHARS_PER_INPUT = int(__import__("os").environ.get("OA_MAX_CHARS_PER_INPUT", "200000"))


# 추출이 안 됐을 때 돌려주는 문자열의 머리말. 프롬프트에는 그대로 남기지만
# (모델이 '이 자료는 비었다' 를 알아야 [확인 필요] 를 제대로 찍는다) 근거 색인에서는
# 걸러야 한다 — pipeline._indexable() 이 이 목록을 본다.
DIAG_PREFIXES = ("[추출 실패", "[읽지 못함", "[빈 파일", "[지원하지 않는 형식")


def is_diagnostic(text: str) -> bool:
    """추출 결과가 본문이 아니라 진단 메시지인가."""
    return (text or "").lstrip().startswith(DIAG_PREFIXES)


def _clip(text: str) -> str:
    text = (text or "").strip()
    limit = MAX_CHARS_PER_INPUT
    if len(text) > limit:
        return text[:limit] + f"\n…(이하 {len(text) - limit}자 생략)…"
    return text


# ── 포맷별 추출기 ────────────────────────────────────────────────────
def _from_txt(path: str) -> str:
    for enc in ("utf-8", "cp949", "euc-kr", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            continue
    with open(path, "rb") as f:
        return f.read().decode("utf-8", errors="replace")


def _from_docx(path: str) -> str:
    from docx import Document  # python-docx
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    # 표 셀 텍스트도 포함
    for tbl in doc.tables:
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _from_xlsx(path: str) -> str:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"# 시트: {ws.title}")
        for row in ws.iter_rows(values_only=True):
            vals = ["" if v is None else str(v) for v in row]
            if any(v.strip() for v in vals):
                out.append(" | ".join(vals))
    wb.close()
    return "\n".join(out)


def _from_pptx(path: str) -> str:
    from pptx import Presentation  # python-pptx
    prs = Presentation(path)
    out = []
    for i, slide in enumerate(prs.slides, 1):
        out.append(f"# 슬라이드 {i}")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = "".join(run.text for run in para.runs).strip()
                    if t:
                        out.append(t)
            if shape.has_table:
                for row in shape.table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if any(cells):
                        out.append(" | ".join(cells))
    return "\n".join(out)


def _from_pdf(path: str) -> str:
    # 1순위 pdfplumber(레이아웃 우수), 폴백 pypdf
    try:
        import pdfplumber
        out = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                out.append(page.extract_text() or "")
        text = "\n".join(out).strip()
        if text:
            return text
    except Exception:
        pass
    from pypdf import PdfReader
    reader = PdfReader(path)
    return "\n".join((pg.extract_text() or "") for pg in reader.pages)


def _from_hwpx(path: str) -> str:
    """HWPX = ZIP 컨테이너. Contents/section*.xml 의 <hp:t> 텍스트 수집.

    표는 셀 순서대로 이어붙이며, 가능하면 행 구분을 살린다.
    """
    from lxml import etree

    HP_T = "{http://www.hancom.co.kr/hwpml/2011/paragraph}t"
    HP_P = "{http://www.hancom.co.kr/hwpml/2011/paragraph}p"

    out_lines = []
    with zipfile.ZipFile(path) as z:
        names = sorted(
            n for n in z.namelist()
            if n.startswith("Contents/section") and n.endswith(".xml")
        )
        if not names:
            # PrvText.txt(미리보기 텍스트)라도 활용
            if "Preview/PrvText.txt" in z.namelist():
                return z.read("Preview/PrvText.txt").decode("utf-8", "replace")
        for name in names:
            data = z.read(name)
            root = etree.fromstring(data)
            # 단락 단위로 텍스트를 모아 줄바꿈.
            # ★ 표를 품은 바깥 문단은 건너뛴다. p.iter() 가 **자손 전체**를 훑기 때문에,
            #   표가 든 문단을 그대로 세면 셀 안의 글이 통째로 한 번 더 실린다
            #   (실측: 계획서 1ch 이 중복 때문에 4만자 상한에 걸려 뒷부분이 잘렸다).
            for p in root.iter(HP_P):
                if any(True for _ in p.iterdescendants(HP_P)):
                    continue
                texts = [t.text for t in p.iter(HP_T) if t.text]
                line = "".join(texts).strip()
                if line:
                    out_lines.append(line)
            # 단락 태그가 없는 예외적 구조 대비
            if not out_lines:
                out_lines = [t.text for t in root.iter(HP_T) if t.text and t.text.strip()]
    return "\n".join(out_lines)


def _from_hwp(path: str) -> str:
    """HWP 5.0(한글 97~2007) — OLE 복합문서. 순수 파이썬으로 본문만 뽑는다.

    hwpx 와 달리 ZIP/XML 이 아니라 바이너리다. `BodyText/SectionN` 스트림이
    raw-deflate 로 눌려 있고, 그 안이 (태그, 레벨, 크기) 헤더를 가진 레코드의 나열이다.
    본문 글자는 `HWPTAG_PARA_TEXT`(67) 레코드에 UTF-16LE 로 들어 있다.

    한컴오피스 없이도 읽으려고 이 길을 열어 뒀다. 없으면 1차 사업 산출물(.hwp)을
    문체 기준으로 넣을 수가 없는데, 변환해 달라고 떠넘기는 것보다 읽는 편이 낫다.
    """
    import struct
    import zlib

    import olefile

    HWPTAG_PARA_TEXT = 67
    # 제어문자가 차지하는 폭이 셋으로 갈린다. 이걸 틀리면 그 뒤 글자가 통째로 밀린다.
    CHAR_CTRL = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}          # 1 글자
    WIDE_CTRL = {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16,
                 17, 18, 19, 20, 21, 22, 23}                          # 8 글자(16바이트)

    ole = olefile.OleFileIO(path)
    try:
        header = ole.openstream("FileHeader").read()
        flags = struct.unpack("<I", header[36:40])[0]
        # 암호·배포용은 본문(BodyText)을 못 연다. 그래도 **미리보기 텍스트**는
        # 눌리지 않은 채로 남아 있어서, 짧은 공문 한 장짜리는 그걸로 전부 건진다
        # (실측: 배포용 공문에서 578자 = 본문 전체가 PrvText 에 있었다).
        locked = ""
        if flags & 0x02:
            locked = "암호가 걸린"
        elif flags & 0x04:
            locked = "배포용(복사 방지)"
        if locked:
            prv = _hwp_prvtext(ole)
            if prv:
                return (f"[{locked} hwp 라 본문 대신 미리보기 텍스트만 읽었습니다 — "
                        f"{os.path.basename(path)}]\n{prv}")
            return f"[읽지 못함: {locked} hwp 입니다 — {os.path.basename(path)}]"
        compressed = bool(flags & 0x01)

        sections = sorted(
            ("/".join(s) for s in ole.listdir() if s[0] == "BodyText"),
            key=lambda n: int("".join(ch for ch in n if ch.isdigit()) or 0),
        )
        lines = []
        for name in sections:
            raw = ole.openstream(name).read()
            data = zlib.decompress(raw, -15) if compressed else raw
            pos = 0
            while pos + 4 <= len(data):
                head = struct.unpack("<I", data[pos:pos + 4])[0]
                pos += 4
                tag, size = head & 0x3FF, (head >> 20) & 0xFFF
                if size == 0xFFF:                       # 크기가 넘치면 다음 4바이트에 별도로
                    size = struct.unpack("<I", data[pos:pos + 4])[0]
                    pos += 4
                if tag == HWPTAG_PARA_TEXT:
                    lines.append(_hwp_para(data[pos:pos + size], CHAR_CTRL, WIDE_CTRL))
                pos += size
        return "\n".join(l for l in (x.strip() for x in lines) if l)
    finally:
        ole.close()


def _hwp_prvtext(ole) -> str:
    """hwp 의 `PrvText` 스트림(UTF-16LE 미리보기 글). 없거나 깨지면 ''."""
    try:
        if not ole.exists("PrvText"):
            return ""
        raw = ole.openstream("PrvText").read()
    except Exception:  # noqa: BLE001
        return ""
    text = raw.decode("utf-16-le", "ignore").replace("\x00", "").strip()
    return text if len(text) >= 20 else ""


def _hwp_para(raw: bytes, char_ctrl: set, wide_ctrl: set) -> str:
    import struct
    out = []
    i = 0
    while i + 2 <= len(raw):
        c = struct.unpack("<H", raw[i:i + 2])[0]
        if c in wide_ctrl:
            out.append("\t" if c == 9 else " ")
            i += 16
        elif c in char_ctrl:
            out.append("\n" if c == 10 else "" if c in (0, 13) else " ")
            i += 2
        else:
            out.append(chr(c))
            i += 2
    return "".join(out)


# ── 실물 판별 ────────────────────────────────────────────────────────
# 확장자는 거짓말을 한다. 실무에서 받은 자료 중 여럿이 **이름만 .hwpx 고 알맹이는
# 구 한글(OLE)** 이었다(한글에서 '다른 이름으로 저장' 대신 확장자만 고쳐 붙인 경우).
# 그대로 두면 ZIP 파서가 BadZipFile 을 내고 파일 하나가 통째로 사라진다.
_MAGIC = (
    (b"PK\x03\x04", "zip"),                       # OOXML(docx/xlsx/pptx) · HWPX
    (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "ole"),  # 구 한글/구 오피스
    (b"%PDF-", "pdf"),
)


def sniff(path: str) -> str:
    """파일 앞머리로 컨테이너 종류를 알아낸다. 모르면 ''."""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return ""
    for sig, kind in _MAGIC:
        if head.startswith(sig):
            return kind
    return ""


# 확장자가 기대하는 컨테이너. 여기 없는 확장자는 실물을 따지지 않는다(txt 등).
_EXPECTED = {"hwpx": "zip", "docx": "zip", "xlsx": "zip", "xlsm": "zip",
             "pptx": "zip", "hwp": "ole", "pdf": "pdf"}
# 실물이 이것이면 이 추출기로 돌린다.
_BY_KIND = {"ole": "hwp", "zip": "hwpx", "pdf": "pdf"}


_EXTRACTORS: Dict[str, Callable[[str], str]] = {
    "txt": _from_txt,
    "md": _from_txt,
    "csv": _from_txt,
    "docx": _from_docx,
    "xlsx": _from_xlsx,
    "xlsm": _from_xlsx,
    "pptx": _from_pptx,
    "pdf": _from_pdf,
    "hwpx": _from_hwpx,
    "hwp": _from_hwp,
}

SUPPORTED_EXTS = sorted(_EXTRACTORS.keys())


# 모르는 확장자라도 알아볼 수 있는 것은 무엇을 해야 하는지 알려 준다.
_HINTS: Dict[str, str] = {
    "hwt":  "한글 서식 파일입니다. 내용을 담은 .hwpx 로 저장해 주세요.",
    "doc":  "구형 Word 형식입니다. Word 에서 **.docx** 로 저장해 주세요.",
    "xls":  "구형 Excel 형식입니다. Excel 에서 **.xlsx** 로 저장해 주세요.",
    "ppt":  "구형 PowerPoint 형식입니다. **.pptx** 로 저장해 주세요.",
}


def _looks_binary(text: str) -> bool:
    """이 문자열이 사람이 읽을 글인가, 바이너리를 억지로 디코딩한 쓰레기인가.

    latin-1 은 **어떤 바이트든** 디코딩에 성공한다. 그래서 폴백이 늘 '성공'하고,
    .hwp 같은 바이너리가 수만 자짜리 깨진 글자로 프롬프트에 그대로 실려 간다.
    실패하지도 않으니 눈치채기 어렵고, 토큰만 태우면서 모델을 망친다.
    """
    sample = text[:4000]
    if not sample:
        return False
    bad = sum(1 for c in sample
              if (ord(c) < 32 and c not in "\t\n\r") or c in "�\x00")
    return bad / len(sample) > 0.05


def extract(path: str) -> str:
    """경로의 파일을 텍스트로 추출. 실패 시 진단 문자열 반환(예외 던지지 않음).

    ★ 확장자보다 **실물(매직바이트)** 을 먼저 믿는다. 이름만 .hwpx 인 구 한글 파일이
      흔해서, 확장자만 보고 ZIP 파서를 부르면 파일 하나가 통째로 날아간다.
    """
    name = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower().lstrip(".")

    try:
        if os.path.getsize(path) == 0:
            return f"[빈 파일: {name} — 0바이트라 읽을 내용이 없습니다]"
    except OSError:
        pass

    fn = _EXTRACTORS.get(ext)
    kind = sniff(path)
    want = _EXPECTED.get(ext)
    if kind and want and kind != want:
        # 확장자와 알맹이가 다르다 — 알맹이 쪽으로 간다.
        alt_ext = _BY_KIND.get(kind)
        alt = _EXTRACTORS.get(alt_ext or "")
        if kind == "ole" and want == "zip" and ext != "hwpx":
            # .docx/.xlsx/.pptx 인데 구 오피스 이진 — 우리 파서로는 못 읽는다.
            hint = _HINTS.get({"docx": "doc", "xlsx": "xls", "pptx": "ppt"}.get(ext, ""))
            return (f"[읽지 못함: {name} — 확장자는 .{ext} 지만 알맹이는 구 오피스 이진입니다]\n"
                    + (hint or "원본 프로그램에서 최신 형식으로 다시 저장해 주세요."))
        if alt is not None:
            text = _safe(alt, path)
            if not is_diagnostic(text):
                return text
            # 알맹이 기준으로도 실패하면 원래 확장자로 한 번 더 시도한다.
    if fn is not None:
        return _safe(fn, path)

    if kind and _BY_KIND.get(kind) in _EXTRACTORS:
        # 확장자를 모르지만 알맹이는 아는 것 — 예: 확장자 없는 hwp
        return _safe(_EXTRACTORS[_BY_KIND[kind]], path)

    hint = _HINTS.get(ext)
    if hint:
        return f"[읽지 못함: .{ext} — {name}]\n{hint}"

    text = _safe(_from_txt, path)
    if _looks_binary(text):
        return (f"[읽지 못함: .{ext} 는 텍스트가 아닙니다 — {name}]\n"
                f"읽을 수 있는 형식: {', '.join(SUPPORTED_EXTS)}")
    return f"[지원하지 않는 형식: .{ext} — 텍스트로 읽었습니다]\n{text}"


def _safe(fn: Callable[[str], str], path: str) -> str:
    try:
        return _clip(fn(path))
    except ModuleNotFoundError as e:
        return f"[추출 실패: 라이브러리 미설치 ({e.name}). setup.bat 을 다시 실행하세요. 파일: {os.path.basename(path)}]"
    except Exception as e:
        return f"[추출 실패: {type(e).__name__}: {e} — 파일: {os.path.basename(path)}]"
