# -*- coding: utf-8 -*-
"""PPTX 폰트 — 한글 렌더 폰트 지정(a:ea)과 폰트 임베드.

왜 필요한가
-----------
python-pptx 의 ``run.font.name`` 은 ``<a:latin>`` 만 세팅한다. PowerPoint 는 한글
글리프를 ``<a:ea>``(East Asian) 폰트로 그리므로, latin 만 지정하면 한글은 테마의
minor EA 폰트로 렌더된다 — 코드에서 지정한 폰트가 정작 한글에는 적용되지 않는다.
``apply_font`` 는 latin/ea/cs 를 함께 세팅해 이 구멍을 막는다.

``embed_fonts`` 는 assets/fonts 의 TTF 를 .pptx 안에 심어(embeddedFontLst),
폰트가 설치되지 않은 PC 에서 열어도 레이아웃이 무너지지 않게 한다. 실패해도
덱 생성 자체는 계속되도록 전부 그레이스풀.

폰트 파일(선택): assets/fonts/
    Pretendard-Regular.ttf · Pretendard-SemiBold.ttf · Pretendard-Bold.ttf
    (OFL — 상업 이용·임베드 허용). 없으면 자동으로 '맑은 고딕' 폴백.
"""
from __future__ import annotations

import os
import struct
from typing import Dict, List, NamedTuple, Optional

PREFERRED = "Pretendard"
FALLBACK = "맑은 고딕"

# assets/fonts 에서 찾는 파일 → (임베드 슬롯, 웨이트 키)
_WANTED = (
    ("Pretendard-Regular.ttf", "regular", "r"),
    ("Pretendard-SemiBold.ttf", "regular", "sb"),
    ("Pretendard-Bold.ttf", "bold", "b"),
)

# 파일명에서 웨이트를 읽는 규칙 — 어느 글꼴이든 Regular/Bold 두 슬롯만 쓴다.
# ★ 가변폰트(fvar 를 가진 VF)는 넣지 않는다. 파워포인트의 임베드(fntdata)는
#   정적 폰트를 전제해서, VF 를 심으면 받는 쪽에서 글꼴이 깨지거나 무시된다.
#   VF 밖에 없으면 무게별 정적 인스턴스로 변환해 두어야 한다
#   (tools/build_fonts.py 참고).
_BOLD_HINTS = ("bold", "-b.", "semibold", "600", "700")

_FNTDATA_CT = "application/x-fontdata"
_FONT_RT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/font")
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


class FontSet(NamedTuple):
    """덱 전체가 쓰는 폰트 3종. ``embedded`` 가 False 면 시스템 폴백 상태."""
    regular: str
    semibold: str
    bold: str
    embedded: bool

    def face(self, weight: str) -> str:
        return {"r": self.regular, "sb": self.semibold, "b": self.bold}.get(
            weight, self.regular)

    def is_bold(self, weight: str) -> bool:
        """SemiBold 전용 페이스가 없으면 '굵게' 로 흉내 낸다.

        Pretendard 처럼 SemiBold 가 **별도 패밀리**로 있을 때만 굵게 처리를 끈다.
        Noto Sans KR 처럼 Regular/Bold 두 벌만 심은 경우 sb 를 굵게로 두지 않으면
        키커·카드 라벨·표 머리글이 본문과 같은 굵기가 되어 위계가 사라진다.
        """
        if weight == "b":
            return True
        if weight == "sb":
            return (not self.embedded) or (self.semibold == self.regular)
        return False


SYSTEM_SET = FontSet(FALLBACK, FALLBACK, FALLBACK, embedded=False)


def fonts_dir(assets_dir) -> str:
    return os.path.join(str(assets_dir), "fonts")


def available(assets_dir) -> Dict[str, str]:
    """assets/fonts 에 실제로 있는 {파일명: 절대경로}."""
    d = fonts_dir(assets_dir)
    if not os.path.isdir(d):
        return {}
    return {f: os.path.join(d, f) for f, _s, _w in _WANTED
            if os.path.isfile(os.path.join(d, f))}


def _is_variable(path: str) -> bool:
    """가변폰트(fvar 테이블 보유) 여부. 임베드 대상에서 제외하려는 것."""
    try:
        with open(path, "rb") as f:
            data = f.read(12 + 16 * 64)
        n = struct.unpack(">H", data[4:6])[0]
        for i in range(min(n, 64)):
            if data[12 + i * 16: 16 + i * 16] == b"fvar":
                return True
    except Exception:  # noqa: BLE001
        pass
    return False


def families(assets_dir) -> Dict[str, Dict[str, str]]:
    """assets/fonts 안의 정적 TTF 를 {패밀리명: {슬롯: 경로}} 로 모은다.

    파일명을 정해 두지 않는다 — 회사가 어떤 글꼴을 주든 심을 수 있어야 한다.
    """
    d = fonts_dir(assets_dir)
    out: Dict[str, Dict[str, str]] = {}
    if not os.path.isdir(d):
        return out
    for fn in sorted(os.listdir(d)):
        if not fn.lower().endswith((".ttf", ".otf")):
            continue
        p = os.path.join(d, fn)
        if _is_variable(p):
            continue                       # 가변폰트는 임베드 불가
        fam = _ttf_family_name(p)
        if not fam:
            continue
        slot = "bold" if any(h in fn.lower() for h in _BOLD_HINTS) else "regular"
        out.setdefault(fam, {}).setdefault(slot, p)
    return out


def font_set(assets_dir, prefer: Optional[str] = None) -> FontSet:
    """이 덱에서 쓸 폰트. prefer(견본이 지정한 글꼴)가 assets/fonts 에 있으면 그것.

    없으면 Pretendard, 그것도 없으면 시스템 폴백(맑은 고딕).
    """
    fams = families(assets_dir)
    if prefer:
        for fam, slots in fams.items():
            if fam.lower() == prefer.strip().lower() and "regular" in slots:
                return FontSet(regular=fam, semibold=fam, bold=fam, embedded=True)
    files = available(assets_dir)
    if "Pretendard-Regular.ttf" not in files:
        return SYSTEM_SET
    base = _ttf_family_name(files["Pretendard-Regular.ttf"]) or PREFERRED
    sb = base
    if "Pretendard-SemiBold.ttf" in files:
        sb = _ttf_family_name(files["Pretendard-SemiBold.ttf"]) or base
    return FontSet(regular=base, semibold=sb, bold=base, embedded=True)


# ── 런 단위 폰트 적용 ──────────────────────────────────────────────────────
def apply_face(run, name: str, *, spc_em: float = 0.0, size_pt=None) -> None:
    """run 의 latin/ea/cs 를 모두 name 으로. spc_em 은 자간(em) → rPr@spc(1/100pt)."""
    from pptx.oxml.ns import qn

    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)
    if spc_em and size_pt:
        rPr.set("spc", str(int(round(spc_em * float(size_pt) * 100))))


def apply_face_to_cell(cell, name: str) -> None:
    """표 셀 안 모든 run 에 폰트 적용(크기·굵기는 건드리지 않음)."""
    for para in cell.text_frame.paragraphs:
        for run in para.runs:
            apply_face(run, name)


# ── TTF name 테이블에서 패밀리명 읽기 ──────────────────────────────────────
def _ttf_family_name(path: str) -> Optional[str]:
    try:
        with open(path, "rb") as f:
            data = f.read()
        num_tables = struct.unpack(">H", data[4:6])[0]
        name_off = None
        for i in range(num_tables):
            rec = 12 + i * 16
            if data[rec:rec + 4] == b"name":
                name_off = struct.unpack(">I", data[rec + 8:rec + 12])[0]
                break
        if name_off is None:
            return None
        cnt, str_off = struct.unpack(">HH", data[name_off + 2:name_off + 6])
        best = None
        for i in range(cnt):
            r = name_off + 6 + i * 12
            pid, eid, lid, nid, ln, off = struct.unpack(">HHHHHH", data[r:r + 12])
            if nid != 1:
                continue
            raw = data[name_off + str_off + off: name_off + str_off + off + ln]
            try:
                txt = raw.decode("utf-16-be") if pid == 3 else raw.decode("latin-1")
            except Exception:  # noqa: BLE001
                continue
            if not txt:
                continue
            if pid == 3 and eid == 1 and lid == 0x409:   # Windows / en-US 우선
                return txt
            best = best or txt
        return best
    except Exception:  # noqa: BLE001
        return None


# ── 폰트 임베드 ────────────────────────────────────────────────────────────
def embed_fonts(prs, assets_dir, only: Optional[str] = None) -> List[str]:
    """assets/fonts 의 TTF 를 .pptx 에 임베드. 임베드된 패밀리명 목록 반환.

    presentation.xml 에 embedTrueTypeFonts="1" 와 <p:embeddedFontLst> 를 넣고,
    폰트 바이트를 /ppt/fonts/fontN.fntdata 파트로 추가한다. 어느 단계든 실패하면
    [] 를 돌려주고 덱 생성은 그대로 진행한다(그레이스풀).
    """
    try:
        from pptx.opc.package import Part
        from pptx.opc.packuri import PackURI
        from pptx.oxml.ns import qn
    except Exception:  # noqa: BLE001
        return []

    # 패밀리별 슬롯. only 를 주면 그 패밀리만 심는다(덱이 실제로 쓰는 글꼴 하나).
    fams = families(assets_dir)
    if only:
        fams = {k: v for k, v in fams.items() if k.lower() == only.strip().lower()}
    if not fams:
        return []

    try:
        pres_part, pres_el = prs.part, prs.element
        pres_el.set("embedTrueTypeFonts", "1")
        pres_el.set("saveSubsetFonts", "0")

        old = pres_el.find(qn("p:embeddedFontLst"))
        if old is not None:
            pres_el.remove(old)
        # ★ 양식이 이미 글꼴을 품고 있으면 그 **관계까지** 끊는다.
        #   목록 원소만 지우면 옛 파트가 관계로 남아 그대로 저장되고, 새 파트가
        #   같은 /ppt/fonts/fontN.fntdata 이름으로 또 들어가 zip 에 두 벌이 된다
        #   (시안을 엔진으로 생성하면서 실제로 파일이 1.9배가 됐다).
        for rid in [rid for rid, rel in list(pres_part.rels.items())
                    if getattr(rel, "reltype", "") == _FONT_RT]:
            try:
                pres_part.drop_rel(rid)
            except Exception:  # noqa: BLE001
                pass
        lst = pres_el.makeelement(qn("p:embeddedFontLst"), {})

        n, embedded = 0, []
        for fam, slots in fams.items():
            ef = lst.makeelement(qn("p:embeddedFont"), {})
            fel = ef.makeelement(qn("p:font"), {})
            fel.set("typeface", fam)
            fel.set("pitchFamily", "34")
            fel.set("charset", "-127")           # Hangul
            ef.append(fel)
            for slot in ("regular", "bold"):
                path = slots.get(slot)
                if not path:
                    continue
                n += 1
                with open(path, "rb") as fp:
                    blob = fp.read()
                part = Part(PackURI(f"/ppt/fonts/font{n}.fntdata"), _FNTDATA_CT,
                            pres_part.package, blob)
                sel = ef.makeelement(qn(f"p:{slot}"), {})
                sel.set(_R_ID, pres_part.relate_to(part, _FONT_RT))
                ef.append(sel)
            if len(ef) > 1:                      # p:font 외 최소 1슬롯
                lst.append(ef)
                embedded.append(fam)
        if len(lst) == 0:
            return []

        # embeddedFontLst 는 sldSz/notesSz 뒤에 온다
        anchor = None
        for tag in ("p:sldIdLst", "p:sldSz", "p:notesSz"):
            found = pres_el.find(qn(tag))
            if found is not None:
                anchor = found
        if anchor is not None:
            anchor.addnext(lst)
        else:
            pres_el.append(lst)
        return embedded
    except Exception as e:  # noqa: BLE001
        print(f"[font] 임베드 실패(무시하고 진행): {e}", flush=True, file=__import__("sys").stderr)
        return []
