"""HWPX 생성 — 회사 양식 스펙(hwpx_template_spec.md)에 맞춘 두 가지 경로.

(A) 한컴 OLE 모드 [권장·목표 PC]
    Windows + 한컴 한글 + pywin32 환경에서 HWPFrame.HwpObject 로 본문을 주입.
    템플릿(assets/hwpx_template.hwpx)이 있으면 열고 없으면 새 문서.

(B) ZIP 클론 모드 [폴백]
    회사 템플릿 .hwpx 를 그대로 복제하여 header.xml(폰트/스타일/페이지설정)을
    100% 재사용하고, 본문 섹션 내용만 생성 단락으로 교체한다.
    header 의 명명 스타일(제목/제목 1~3/본문/목록/인용)을 이름으로 찾아 매핑.

(C) 둘 다 불가하면 HwpxUnavailable 을 올려 pipeline 이 Markdown 으로 대체한다.
"""
from __future__ import annotations

import copy
import os
import re
import zipfile
from typing import Dict, List, Optional

from . import mdblocks
from . import md_gen

HH = "http://www.hancom.co.kr/hwpml/2011/head"
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HS = "http://www.hancom.co.kr/hwpml/2011/section"

# 진짜 표 렌더링 토글. 한글이 표를 못 열면 .env 에 STUDIO_HWPX_REAL_TABLES=0 → 즉시 평탄화 폴백.
_REAL_TABLES = os.environ.get("STUDIO_HWPX_REAL_TABLES", "1") not in ("0", "false", "False", "")


class HwpxUnavailable(RuntimeError):
    pass


# ── 진단 ────────────────────────────────────────────────────────────
def hancom_available() -> bool:
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except Exception:
        return False
    return os.name == "nt"


def render(payload, out_path: str, template: Optional[str] = None) -> str:
    md = md_gen.to_markdown(payload)

    # (A) 한컴 OLE
    if hancom_available():
        try:
            return _render_hancom(md, out_path, template)
        except Exception as e:
            print(f"[hwpx_export] 한컴 OLE 실패, ZIP 폴백 시도: {e}", file=__import__("sys").stderr)

    # (B) ZIP 클론 (템플릿 필요)
    if template and os.path.exists(template):
        return _render_zip(md, out_path, template)

    raise HwpxUnavailable(
        "HWPX 를 생성할 수 없습니다.\n"
        " - 한컴 한글+pywin32 가 설치된 Windows 에서 실행하거나\n"
        " - assets/hwpx_template.hwpx 회사 양식을 넣어주세요.\n"
        "지금은 Markdown(.md) 으로 결과를 제공합니다."
    )


# ── (A) 한컴 OLE 모드 ────────────────────────────────────────────────
def _render_hancom(md: str, out_path: str, template: Optional[str]) -> str:
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    try:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
        except Exception:
            pass
        if template and os.path.exists(template):
            hwp.Open(os.path.abspath(template), "HWPX", "")
            hwp.MovePos(3)  # 문서 끝으로
        else:
            hwp.HAction.Run("FileNew")

        text = _md_to_plain(md)
        pset = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault("InsertText", pset.HSet)
        pset.Text = text
        hwp.HAction.Execute("InsertText", pset.HSet)

        save = hwp.HParameterSet.HFileOpenSave
        hwp.HAction.GetDefault("FileSaveAs_S", save.HSet)
        save.filename = os.path.abspath(out_path)
        save.Format = "HWPX"
        hwp.HAction.Execute("FileSaveAs_S", save.HSet)
        try:
            hwp.Quit()
        except Exception:
            pass
        return out_path
    finally:
        pythoncom.CoUninitialize()


def _md_to_plain(md: str) -> str:
    """OLE InsertText 용 평문. 표는 탭 정렬 텍스트로, <br>는 이미 ' / '로 정규화됨."""
    lines = []
    for b in mdblocks.parse(md):
        t = b["type"]
        if t == "table":
            if b.get("header"):
                lines.append("\t".join(b["header"]))
            for row in b.get("rows", []):
                lines.append("\t".join(row))
        elif t == "heading":
            lines.append(b["text"])
        elif t == "bullet":
            lines.append("· " + b["text"])
        elif t == "quote":
            lines.append("  " + b["text"])
        else:
            lines.append(b["text"])
    return "\r\n".join(lines)


# ── (B) ZIP 클론 모드 ────────────────────────────────────────────────
def _render_zip(md: str, out_path: str, template: str) -> str:
    # 1) 템플릿 통째로 메모리에 로드 (압축정보 보존)
    src = zipfile.ZipFile(template)
    members = src.infolist()
    blobs: Dict[str, bytes] = {m.filename: src.read(m.filename) for m in members}

    # 2) header.xml 에서 명명 스타일 맵 구성
    style_map = _style_map(blobs.get("Contents/header.xml", b""))

    section_names = sorted(n for n in blobs if n.startswith("Contents/section") and n.endswith(".xml"))
    blocks = mdblocks.parse(md)
    title = next((b["text"] for b in blocks if b["type"] == "heading"), "")

    # 3) 본문은 '마지막' 섹션에 주입하고 앞 섹션(표지/간지)은 그대로 보존한다.
    #    (표지=양식 첫 페이지 그대로. 생성 콘텐츠의 도입부=표지정보는 떼어낸다.)
    blocks = _drop_cover_info(blocks)
    body_name = section_names[-1] if section_names else None

    # 3-b) 표 청사진은 '본문 섹션'의 표에서 복제(표지의 디자인 표 말고) → 음영 없는 본문용 표.
    tbl_tmpl = None
    if body_name:
        tbl_tmpl = _extract_table_template({body_name: blobs[body_name]})
    if tbl_tmpl is None:
        tbl_tmpl = _extract_table_template(blobs)

    bf_index = _borderfill_index(blobs.get("Contents/header.xml", b""))
    if body_name:
        _hdr = blobs.get("Contents/header.xml", b"")
        blobs[body_name] = _rewrite_section(blobs[body_name], blocks, style_map, tbl_tmpl,
                                            bf_index, _green_charpr(_hdr),
                                            _center_parapr(_hdr, style_map))
    # 앞 섹션들(표지 등)은 미수정 → 양식 그대로 보존(빈 섹션도 안 만들어 끝 빈 페이지 방지)

    # 4) 표지 자리표시자 치환. 제목이 길면 **두 줄로 나눈다** —
    #    한 줄에 욱여넣으면 한글이 글자를 눌러 겹쳐 보인다(실측: 사업명 38자).
    if title:
        line1, line2 = _split_title(title)
        for name in section_names:
            xml = blobs[name].decode("utf-8", "replace")
            if "_제목_입력_" in xml or "_제목2_입력_" in xml:
                xml = xml.replace("_제목2_입력_", _xml_escape(line2))
                xml = xml.replace("_제목_입력_", _xml_escape(line1))
                blobs[name] = xml.encode("utf-8")

    # 5) re-zip — mimetype 을 STORED 로 가장 먼저
    _write_hwpx(out_path, members, blobs)
    src.close()
    return out_path


TITLE_ONE_LINE = 22          # 이보다 짧으면 한 줄로 둔다


def _split_title(title: str):
    """표지 제목을 두 줄로. 가운데에 가장 가까운 **띄어쓰기**에서 자른다."""
    t = " ".join((title or "").split())
    if len(t) <= TITLE_ONE_LINE or " " not in t:
        return t, ""
    mid = len(t) / 2
    best, bestd = None, None
    for i, ch in enumerate(t):
        if ch != " ":
            continue
        d = abs(i - mid)
        if bestd is None or d < bestd:
            best, bestd = i, d
    if best is None:
        return t, ""
    return t[:best].strip(), t[best + 1:].strip()


_NUM_HEAD = re.compile(r"^(?:[IVXivx]+|[0-9]+|[가-하]|[ㅇ○◦])\s*[.)]?\s+")


def _center_parapr(header_bytes: bytes, style_map) -> Optional[str]:
    """가운데 정렬 문단모양 id. 표 캡션에 쓴다.

    이름 붙은 스타일이 쓰는 문단모양은 피한다 — 제목용(큰 글자·앞뒤 여백)을 잡으면
    캡션이 제목처럼 커진다.
    """
    if not header_bytes:
        return None
    used = {v.get("paraPrIDRef") for v in (style_map or {}).values()}
    found = []
    for m in re.finditer(rb'<hh:paraPr id="(\d+)"(.*?)</hh:paraPr>', header_bytes, re.S):
        if b'horizontal="CENTER"' in m.group(2):
            found.append(m.group(1).decode())
    for pid in found:
        if pid not in used:
            return pid
    return found[0] if found else None


def _caption_text(head_text: str) -> str:
    """캡션 글. 가장 가까운 제목에서 번호를 떼어 쓴다."""
    t = _NUM_HEAD.sub("", (head_text or "").strip()).strip()
    return "▲ " + (t or "표")


def _make_caption(text: str, normal, parapr: Optional[str]):
    """표 **아래** 가운데 정렬 캡션 문단."""
    from lxml import etree
    p = etree.Element(f"{{{HP}}}p")
    p.set("paraPrIDRef", parapr or normal["paraPrIDRef"])
    p.set("styleIDRef", normal["styleId"])
    run = etree.SubElement(p, f"{{{HP}}}run")
    run.set("charPrIDRef", normal["charPrIDRef"])
    etree.SubElement(run, f"{{{HP}}}t").text = text
    return p


def _green_charpr(header_bytes: bytes) -> Optional[str]:
    """양식에 심어 둔 '확인 필요' 강조용 녹색 charPr id. 없으면 None."""
    if not header_bytes:
        return None
    m = re.search(rb'<hh:charPr id="(\d+)"[^>]*textColor="#1F9155"', header_bytes)
    return m.group(1).decode() if m else None


def _style_map(header_bytes: bytes) -> Dict[str, Dict[str, str]]:
    """style name → {styleId, paraPrIDRef, charPrIDRef}. 실패해도 빈 dict."""
    from lxml import etree
    out: Dict[str, Dict[str, str]] = {}
    if not header_bytes:
        return out
    try:
        root = etree.fromstring(header_bytes)
    except Exception:
        return out
    for st in root.iter(f"{{{HH}}}style"):
        name = st.get("name")
        if not name:
            continue
        out[name.strip()] = {
            "styleId": st.get("id", "0"),
            "paraPrIDRef": st.get("paraPrIDRef", "0"),
            "charPrIDRef": st.get("charPrIDRef", "0"),
        }
    return out


def _pick(style_map: Dict[str, Dict[str, str]], *names: str) -> Optional[Dict[str, str]]:
    for n in names:
        if n in style_map:
            return style_map[n]
    return None


def _rewrite_section(section_bytes: bytes, blocks: List[Dict],
                     style_map: Dict[str, Dict[str, str]], tbl_tmpl=None, bf_index=None,
                     green=None, center=None) -> bytes:
    """섹션의 본문 단락을 교체. 첫 단락(secPr=페이지설정)은 보존."""
    from lxml import etree

    root = etree.fromstring(section_bytes)
    ptag = f"{{{HP}}}p"
    paras = [c for c in root if c.tag == ptag]

    # 페이지 설정(secPr)을 가진 첫 단락 찾기
    secpr_para = None
    for p in paras:
        if p.find(f".//{{{HP}}}secPr") is not None:
            secpr_para = p
            break
    if secpr_para is None and paras:
        secpr_para = paras[0]

    # 기존 단락 전부 제거
    for p in paras:
        root.remove(p)

    # 페이지설정 단락 복원(텍스트는 비움 → secPr 만 유지)
    if secpr_para is not None:
        _strip_text(secpr_para)
        root.append(secpr_para)

    # 생성 단락 추가
    normal = _pick(style_map, "본문") or {"styleId": "0", "paraPrIDRef": "0", "charPrIDRef": "0"}
    emph = _pick(style_map, "강조")
    head_level = 0            # 마지막으로 지난 제목의 깊이(# 개수)
    head_text = ""            # 그 제목의 글 — 표 캡션 이름으로 쓴다
    for blk in blocks:
        if _is_divider(blk):
            continue
        if blk["type"] == "heading":
            head_level = blk.get("level", 0) or 0
            head_text = blk.get("text", "") or ""
        if blk["type"] == "table":
            # 진짜 표(템플릿 복제). 실패하면 평탄화(불릿)로 폴백 → 문서는 무조건 나오게.
            if tbl_tmpl is not None and _REAL_TABLES:
                try:
                    root.append(_make_table(tbl_tmpl, blk, style_map, bf_index, normal))
                    root.append(_make_caption(_caption_text(head_text), normal, center))
                    continue
                except Exception as e:  # noqa: BLE001
                    print(f"[hwpx_export] 표 렌더 실패, 평탄화 폴백: {e}", file=__import__("sys").stderr)
            for b in _table_to_blocks(blk):
                st = _style_for_block(b, style_map, normal, head_level)
                root.append(_make_para(b, st, emph, normal, green))
            continue
        st = _style_for_block(blk, style_map, normal, head_level)
        root.append(_make_para(blk, st, emph, normal, green))

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _table_to_blocks(blk: Dict) -> List[Dict]:
    """표 블록 → 단락 블록들(헤더는 굵게 한 줄, 각 행은 ' / '로 이은 불릿)."""
    out: List[Dict] = []
    hdr = [c for c in (blk.get("header") or []) if c]
    if hdr:
        txt = " · ".join(hdr)
        out.append({"type": "para", "level": 0, "text": txt, "runs": [(txt, True)]})
    for row in blk.get("rows", []):
        cells = [c for c in row if c]
        if not cells:
            continue
        txt = " / ".join(cells)
        out.append({"type": "bullet", "level": 1, "text": txt, "runs": [(txt, False)]})
    return out


# ── 진짜 표(테이블) 렌더링 — 템플릿의 기존 표를 복제해 행/열/텍스트만 교체 ──────────
_ID_SEQ = [2000000000]


def _next_id() -> str:
    _ID_SEQ[0] += 1
    return str(_ID_SEQ[0])


def _extract_table_template(blobs: Dict[str, bytes], section_names_hint=None):
    """표를 감싼 단락(<hp:p>) 중 '셀이 가장 많은' 표를 deepcopy 로 반환(= 내용 표 샘플).
    작은 장식용 표(간지·표지 디자인 표)는 건너뛴다. 이 표의 디자인(색/테두리/스타일)을 그대로
    복제하므로, 사용자가 본문에 둔 '내용 표 샘플'을 꾸미면 생성 표에 그대로 반영된다."""
    from lxml import etree
    ptag = f"{{{HP}}}p"
    tbltag = f"{{{HP}}}tbl"
    tctag = f"{{{HP}}}tc"
    best = None
    best_cells = -1
    for name in sorted(n for n in blobs if n.startswith("Contents/section") and n.endswith(".xml")):
        try:
            root = etree.fromstring(blobs[name])
        except Exception:
            continue
        for tbl in root.iter(tbltag):
            ncells = len(tbl.findall(f".//{tctag}"))
            if ncells <= best_cells:
                continue
            p = tbl
            while p is not None and p.tag != ptag:
                p = p.getparent()
            if p is not None:
                best = copy.deepcopy(p)
                best_cells = ncells
    return best


def _borderfill_index(header_bytes: bytes) -> Dict[str, object]:
    """{borderFill id: element} 인덱스. 표 왼쪽 외곽선 보강용 매칭에 사용."""
    from lxml import etree
    idx: Dict[str, object] = {}
    if not header_bytes:
        return idx
    try:
        root = etree.fromstring(header_bytes)
    except Exception:
        return idx
    for bf in root.iter(f"{{{HH}}}borderFill"):
        idx[bf.get("id")] = bf
    return idx


def _bf_sides(bf) -> Dict[str, tuple]:
    out: Dict[str, tuple] = {}
    for s in ("leftBorder", "rightBorder", "topBorder", "bottomBorder"):
        e = bf.find(f"{{{HH}}}{s}")
        out[s] = (e.get("type"), e.get("width")) if e is not None else (None, None)
    return out


def _match_grid_bf(ref_id: Optional[str], bf_index: Dict[str, object]) -> Optional[str]:
    """ref 셀과 우/상/하 선 굵기가 같고 좌·우·상·하 모두 테두리인 borderFill id 반환.
    (왼쪽 열 셀에 적용 → 그 행과 같은 굵기로 표 좌측 외곽선만 추가.)
    정확히 일치하는 게 없으면 아무 사방 테두리, 그것도 없으면 None."""
    ref = bf_index.get(ref_id) if ref_id else None
    rs = _bf_sides(ref) if ref is not None else None
    fallback = None
    for bid, bf in bf_index.items():
        s = _bf_sides(bf)
        if not all(s[k][0] == "SOLID" for k in ("leftBorder", "rightBorder", "topBorder", "bottomBorder")):
            continue
        if fallback is None:
            fallback = bid
        if rs is not None and s["rightBorder"][1] == rs["rightBorder"][1] \
                and s["topBorder"][1] == rs["topBorder"][1] \
                and s["bottomBorder"][1] == rs["bottomBorder"][1]:
            return bid
    return fallback


def _set_cell(tc, text: str, style: Optional[Dict[str, str]], col: int, row: int,
              width: int, height: int) -> None:
    """복제한 <hp:tc> 의 주소/크기/스타일/텍스트를 교체. linesegarray 는 제거(한글이 재계산)."""
    P = f"{{{HP}}}"
    ca = tc.find(f"{P}cellAddr")
    if ca is not None:
        ca.set("colAddr", str(col)); ca.set("rowAddr", str(row))
    cspan = tc.find(f"{P}cellSpan")
    if cspan is not None:
        cspan.set("colSpan", "1"); cspan.set("rowSpan", "1")
    csz = tc.find(f"{P}cellSz")
    if csz is not None:
        csz.set("width", str(width)); csz.set("height", str(height))
    # 레이아웃 캐시 제거 → 한글이 새 텍스트로 재배치
    for lsa in list(tc.iter(f"{P}linesegarray")):
        lsa.getparent().remove(lsa)
    sub = tc.find(f"{P}subList")
    cp = sub.find(f"{P}p") if sub is not None else None
    if cp is None:
        return
    cp.set("id", _next_id())
    if style:
        cp.set("styleIDRef", style["styleId"])
        cp.set("paraPrIDRef", style["paraPrIDRef"])
    runs = cp.findall(f"{P}run")
    for extra in runs[1:]:
        cp.remove(extra)
    run = runs[0] if runs else None
    if run is None:
        from lxml import etree
        run = etree.SubElement(cp, f"{P}run")
    if style:
        run.set("charPrIDRef", style["charPrIDRef"])
    ts = run.findall(f"{P}t")
    if ts:
        ts[0].text = text
        for extra in ts[1:]:
            run.remove(extra)
    else:
        from lxml import etree
        etree.SubElement(run, f"{P}t").text = text


def _make_table(tbl_tmpl, blk: Dict, style_map: Dict[str, Dict[str, str]], bf_index=None,
                normal=None):
    """템플릿 표 단락(deepcopy 원본)을 복제해 blk(header/rows)로 채운 <hp:p> 반환.
    bf_index: borderFill 인덱스 → 맨 왼쪽 열 셀을 그 행과 같은 굵기의 '사방 테두리'로 보강."""
    P = f"{{{HP}}}"
    p = copy.deepcopy(tbl_tmpl)
    for lsa in list(p.iter(f"{P}linesegarray")):  # 표 단락 레이아웃 캐시 제거
        lsa.getparent().remove(lsa)
    tbl = p.find(f".//{P}tbl")
    if tbl is None:
        raise ValueError("표 템플릿에 <hp:tbl> 없음")

    # 양식의 표 단락에는 캡션 글자('세부일정')가 표와 **같은 문단**에 박혀 있다.
    # 그대로 복제하면 모든 표 위에 그 낱말이 따라붙는다 — 떼어 내고, 캡션은
    # 표 아래에 따로 만든다(_make_caption).
    for t in list(p.iter(f"{P}t")):
        anc, inside = t.getparent(), False
        while anc is not None:
            if anc.tag == f"{P}tbl":
                inside = True
                break
            anc = anc.getparent()
        if not inside:
            t.getparent().remove(t)

    # 개체 속성 '글자처럼 취급' — 표가 글자 흐름에 얹혀 앞뒤 문단과 함께 밀린다.
    pos = tbl.find(f"{P}pos")
    if pos is not None:
        pos.set("treatAsChar", "1")

    # 글자처럼 취급하면 **표가 앉은 문단의 서식**이 표 위치·여백을 좌우한다.
    # 양식에서 떠 온 문단은 제목 스타일이라 앞뒤 여백이 붙는다 → 본문으로 맞춘다.
    if normal:
        p.set("paraPrIDRef", normal["paraPrIDRef"])
        p.set("styleIDRef", normal["styleId"])

    rows_src = tbl.findall(f"{P}tr")
    if not rows_src:
        raise ValueError("표 템플릿에 <hp:tr> 없음")
    # 머리행 셀 / 내용행 셀을 따로 복제해 사용 → 템플릿 표의 머리행·내용행 디자인(색/스타일)을
    # 그대로 재현한다. (셀 스타일을 새로 덮어쓰지 않고 style=None 으로 템플릿 셀을 보존)
    head_tc = rows_src[0].find(f"{P}tc")
    body_tc = rows_src[1].find(f"{P}tc") if len(rows_src) >= 2 else head_tc
    if head_tc is None or body_tc is None:
        raise ValueError("표 템플릿에 <hp:tc> 없음")
    head_tc = copy.deepcopy(head_tc)
    body_tc = copy.deepcopy(body_tc)
    tr_tmpl = copy.deepcopy(rows_src[0])
    for tc in tr_tmpl.findall(f"{P}tc"):
        tr_tmpl.remove(tc)
    for tr in rows_src:
        tbl.remove(tr)

    header = [str(c) for c in (blk.get("header") or [])]
    body = [[str(c) for c in r] for r in (blk.get("rows") or [])]
    ncol = blk.get("ncol") or len(header) or (max((len(r) for r in body), default=1))
    ncol = max(1, int(ncol))
    data = ([header] if header else []) + body
    if not data:
        raise ValueError("빈 표")
    nrow = len(data)

    total_w = 45354          # A4 본문 폭(약 160mm, HWPUNIT). noAdjust=0 이라 한글이 재조정.
    colw = max(900, total_w // ncol)
    rowh = 1400

    tbl.set("rowCnt", str(nrow))
    tbl.set("colCnt", str(ncol))
    tbl.set("id", _next_id())
    sz = tbl.find(f"{P}sz")
    if sz is not None:
        sz.set("width", str(colw * ncol)); sz.set("height", str(rowh * nrow))

    # 왼쪽 열 보강용: 머리/내용 셀과 '같은 굵기'의 사방 테두리 borderFill (굵은 머리선 유지)
    head_grid = _match_grid_bf(head_tc.get("borderFillIDRef"), bf_index) if bf_index else None
    body_grid = _match_grid_bf(body_tc.get("borderFillIDRef"), bf_index) if bf_index else None

    for r, cells in enumerate(data):
        is_head = (r == 0 and bool(header))
        cell_tmpl = head_tc if is_head else body_tc
        gb = head_grid if is_head else body_grid
        tr = copy.deepcopy(tr_tmpl)
        for c in range(ncol):
            tc = copy.deepcopy(cell_tmpl)
            txt = cells[c] if c < len(cells) else ""
            _set_cell(tc, txt, None, c, r, colw, rowh)  # style=None → 템플릿 셀 디자인 보존
            # 맨 왼쪽 열만 '같은 굵기' 사방 테두리로 → 빠져 있던 좌측 외곽선을 그 행 굵기로 추가.
            if c == 0 and gb:
                tc.set("borderFillIDRef", gb)
            tr.append(tc)
        tbl.append(tr)
    return p


_DIVIDER_RE = re.compile(r"\s*[-_*=—–]{3,}\s*")

# 표지정보 제거: 생성 콘텐츠의 도입부(문서 제목 + '항목/내용' 표지정보표)를 떼고
# 첫 '번호 매겨진' 섹션(1. / I. 등)부터 본문으로 본다. 표지는 양식 첫 페이지를 쓴다.
_NUM_HEAD_RE = re.compile(r"^\s*(?:\d+|[IVXivxⅠ-Ⅹ]+)[.)]\s*\S")
_DROP_COVER = os.environ.get("STUDIO_HWPX_DROP_COVER", "1") not in ("0", "false", "False", "")


def _drop_cover_info(blocks: List[Dict]) -> List[Dict]:
    """도입부(문서 제목 + 표지정보표/항목 등)를 떼고 첫 '번호 매겨진' 섹션부터 본문으로.
    첫 16블록 안에서 '1.'/'I.' 형태의 섹션 제목을 찾으면 그 앞을 전부 버린다."""
    if not _DROP_COVER:
        return blocks
    for i, b in enumerate(blocks[:16]):
        if i > 0 and b.get("type") == "heading" and _NUM_HEAD_RE.match(b.get("text", "") or ""):
            return blocks[i:]
    return blocks


def _is_divider(blk: Dict) -> bool:
    """마크다운 구분선(--- *** === 등)은 단락으로 찍지 않고 버린다."""
    if blk.get("type") not in ("para", "bullet"):
        return False
    t = (blk.get("text") or "").strip()
    return bool(t) and _DIVIDER_RE.fullmatch(t) is not None


def _style_for_block(blk: Dict, style_map, normal, head_level: int = 0) -> Dict[str, str]:
    t = blk["type"]
    if t == "heading":
        # `#` 는 문서 제목(표지로 빠진다). 본문 제목은 `##` 부터 시작하므로 한 칸 당겨
        # `##`→제목 1(1.) · `###`→제목 2(가.) · `####`→제목 3(ㅇ) 로 **세 단계를 다 쓴다.**
        # 예전에는 `##`→제목 2 로 밀려서 제목 1 이 놀고, `###`·`####` 가 둘 다 제목 3 이 되어
        # '가.' 와 'ㅇ' 이 같은 모양으로 보였다.
        lvl = min(max(blk["level"] - 1, 1), 3)
        return _pick(style_map, f"제목 {lvl}", "제목", f"Heading {lvl}") or normal
    if t == "bullet":
        # 글머리표는 **지금 어느 제목 아래인지**로 정한다.
        #   가.(제목 2) 아래  → 동그라미(ㅇ), 들여쓴 것 → 마이너스(−)
        #   ㅇ(제목 3) 아래   → 곧바로 마이너스(−)
        # 제목 3 아래에서도 동그라미를 쓰면 제목과 본문이 같은 ㅇ 로 보여 위계가 뭉갠다
        # (실측: 보고서 '가. 목적' 아래가 통째로 ㅇ 한 단계로 눌렸다).
        lvl = (blk.get("level", 1) or 1) + (1 if head_level >= 4 else 0)
        if lvl >= 2:
            return _pick(style_map, "마이너스", "목록") or normal
        return _pick(style_map, "동그라미", "목록") or normal
    if t == "quote":
        return _pick(style_map, "인용") or normal
    return normal


# '아직 못 채운 자리' 표시 — 문서에서 눈에 띄어야 사람이 찾아 고친다.
_PENDING = re.compile(r"\[(?:확인\s*필요|확인필요|To be confirmed|TBC|TBD)[^\]]*\]")


def _split_pending(text: str):
    """[(조각, 확인필요인가)] 로 쪼갠다."""
    out, pos = [], 0
    for m in _PENDING.finditer(text or ""):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(0), True))
        pos = m.end()
    if pos < len(text or ""):
        out.append((text[pos:], False))
    return out or [(text or "", False)]


def _make_para(blk: Dict, st: Dict[str, str], emph, normal, green=None):
    from lxml import etree
    p = etree.Element(f"{{{HP}}}p")
    p.set("paraPrIDRef", st["paraPrIDRef"])
    p.set("styleIDRef", st["styleId"])
    # 글머리표(○/−)는 동그라미/마이너스 스타일의 문단모양에 내장 — 글자 접두를 넣지 않는다.
    prefix = ""
    runs = blk.get("runs") or [(blk.get("text", ""), False)]
    first = True
    for text, bold in runs:
        if not text:
            continue
        if first and prefix:
            text = prefix + text
        first = False
        char_id = (emph or normal)["charPrIDRef"] if bold and emph else st["charPrIDRef"]
        for seg, pending in _split_pending(text):
            if not seg:
                continue
            run = etree.SubElement(p, f"{{{HP}}}run")
            run.set("charPrIDRef", green if (pending and green) else char_id)
            t = etree.SubElement(run, f"{{{HP}}}t")
            t.text = seg
    if first:  # 빈 단락 보장
        run = etree.SubElement(p, f"{{{HP}}}run")
        run.set("charPrIDRef", st["charPrIDRef"])
        etree.SubElement(run, f"{{{HP}}}t").text = ""
    return p


def _strip_text(p) -> None:
    """단락에서 텍스트 run 만 제거하고 secPr/ctrl 등 구조는 유지."""
    for run in list(p):
        if run.tag == f"{{{HP}}}run":
            # secPr/ctrl 류가 들어있는 run 은 보존, 순수 텍스트 run 만 제거
            if run.find(f"{{{HP}}}secPr") is not None or run.find(f"{{{HP}}}ctrl") is not None:
                # 텍스트만 비움
                for t in run.iter(f"{{{HP}}}t"):
                    t.text = ""
            else:
                p.remove(run)


def _write_hwpx(out_path: str, members, blobs: Dict[str, bytes]) -> None:
    # mimetype 을 가장 먼저 STORED 로
    order = ["mimetype"] + [m.filename for m in members if m.filename != "mimetype"]
    seen = set()
    with zipfile.ZipFile(out_path, "w") as zf:
        for name in order:
            if name in seen or name not in blobs:
                continue
            seen.add(name)
            data = blobs[name]
            if name == "mimetype":
                zf.writestr(name, data, compress_type=zipfile.ZIP_STORED)
            else:
                zf.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)


def _xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
