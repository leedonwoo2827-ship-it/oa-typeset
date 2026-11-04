# -*- coding: utf-8 -*-
"""디자인 슬라이드 빌더 — 슬라이드플랜(JSON) → 디자인된 .pptx 바이트.

네이비(#1E2761) + 앰버(#F2A900) 디자인 시스템. 회사 템플릿을 베이스로 열어
(테마·마스터·로고 상속) 예시 슬라이드를 비우고, 모든 슬라이드를 Blank 레이아웃에
도형으로 직접 배치한다.

타입: cover / section / photo / bullets / process / cards / compare / table
      / quiz / objectives / agenda / closing / stat
python-pptx 미설치 시 build_deck 은 None 을 돌려준다(그레이스풀).

디자인 규칙(docs/디자인-토큰.md 참고)
  · 색은 네이비 단일 계조 + 앰버 1점. 무지개 로테이션 금지.
  · 앰버는 채움·마커 전용. 흰 배경 위 본문 텍스트 색으로 쓰지 않는다(대비 2.0:1).
  · 그림자는 사진·카드·비교 컬럼에만. 칩·바탕 도형은 평면.
  · 모든 본문은 _fit() 을 거쳐 박스를 넘지 않는다.
"""
from __future__ import annotations

import io
import json
import math
import os
import re
import threading
from pathlib import Path
from typing import Callable, Dict, List, Optional

from . import fonts as pptx_font

# ── 캔버스 ────────────────────────────────────────────────────────────────
SLIDE_W, SLIDE_H = 13.33, 7.5
MARGIN = 0.6
CONTENT_W = 12.13                      # 전폭 콘텐츠
CONTENT_R = MARGIN + CONTENT_W         # 우측 끝 12.73

ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets"

# ── 색 — 네이비 단일 계조 + 앰버 ───────────────────────────────────────────
_HEX = {
    "navy":     (0x1E, 0x27, 0x61),   # 제목·주요 채움 (흰 글씨 OK)
    "navy_75":  (0x56, 0x5D, 0x88),   # 2단계
    "navy_55":  (0x83, 0x88, 0xA8),   # 3단계·키커·캡션
    "navy_30":  (0xBB, 0xBE, 0xD0),   # 비활성·대형 숫자
    "navy_12":  (0xE4, 0xE5, 0xEC),   # 헤어라인 링
    "chip":     (0xEE, 0xF2, 0xFB),   # 존 배경·표 짝수행·사진 자리 워시
    "amber":    (0xF2, 0xA9, 0x00),   # 마커·강조 채움 (텍스트 금지)
    "amber_dk": (0xB5, 0x7F, 0x00),   # 앰버 계열 텍스트가 꼭 필요할 때
    "amber_wash": (0xFD, 0xF3, 0xDC),  # 강조 존 배경
    "ink":      (0x2B, 0x34, 0x40),   # 본문 글자
    "ink_mid":  (0x47, 0x55, 0x69),   # 본문보다 한 단계 옅은 글자
    "muted":    (0x94, 0xA3, 0xB8),   # 라벨·캡션·푸터 (배색을 따라가지 않는다)
    "navy_dk":  (0x1E, 0x27, 0x61),   # 글자가 얹히는 밴드 배경 (기본은 navy 와 동일)
    "on_navy":  (0xA9, 0xB0, 0xCE),   # 네이비 위 보조 글자
    "white":    (0xFF, 0xFF, 0xFF),
}

# 순서 있는 것(process)·병렬(cards) 모두 단일 계조. 무지개 금지.
_RAMP = ("navy", "navy_75", "navy_55", "navy_30")

# ── 타이포 스케일 ─────────────────────────────────────────────────────────
#   size(pt) / weight(r|sb|b) / color / spc(em 자간) / line(행간)
TYPE = {
    "kicker":      dict(size=12.0, weight="sb", color="muted", spc=0.14),
    "kicker_on":   dict(size=11.0, weight="sb", color="on_navy", spc=0.14),
    "cover_title": dict(size=34.0, weight="b",  color="ink", spc=-0.025, line=1.14),
    "cover_sub":   dict(size=13.5, weight="r",  color="muted", line=1.35),
    "sec_title":   dict(size=32.0, weight="b",  color="ink", spc=-0.02, line=1.18),
    "title":       dict(size=28.0, weight="b",  color="ink", spc=-0.02, line=1.16),
    "chip":        dict(size=15.0, weight="sb", color="navy", line=1.25),
    "bullet":      dict(size=15.0, weight="r",  color="ink", line=1.42),
    "bullet2":     dict(size=13.0, weight="r",  color="muted", line=1.40),
    "card_label":  dict(size=15.5, weight="sb", color="navy", line=1.20),
    "card_desc":   dict(size=12.5, weight="r",  color="ink", line=1.38),
    "node_label":  dict(size=14.5, weight="sb", color="white", line=1.18),
    "node_desc":   dict(size=11.0, weight="r",  color="ink", line=1.30),
    "th":          dict(size=12.0, weight="sb", color="white"),
    "td":          dict(size=11.5, weight="r",  color="ink", line=1.30),
    "caption":     dict(size=9.5,  weight="r",  color="muted"),
    "footer":      dict(size=9.5,  weight="r",  color="muted"),
    "pageno":      dict(size=10.0, weight="sb", color="navy_30"),
    "big_num":     dict(size=34.0, weight="b",  color="navy_30", spc=-0.02),
    "sec_num":     dict(size=96.0, weight="b",  color="navy_12", spc=-0.03),
    "stat_num":    dict(size=52.0, weight="b",  color="navy", spc=-0.03),
    "answer":      dict(size=13.0, weight="sb", color="navy"),
    # ↓ 1차 정답지 3종(0079·0083·1ch)에서 실측해 더한 롤.
    "note":        dict(size=11.0, weight="r",  color="muted", line=1.30),
    "subhead":     dict(size=13.5, weight="sb", color="ink", line=1.25),
    "source":      dict(size=9.0,  weight="r",  color="muted"),
    "interp_lbl":  dict(size=11.25, weight="sb", color="navy_dk"),
    "interp":      dict(size=12.0, weight="r",  color="navy_dk", line=1.32),
    "kpi_label":   dict(size=12.75, weight="sb", color="ink", line=1.20),
    "kpi_num":     dict(size=48.0, weight="b",  color="navy", spc=-0.03),
    "kpi_cap":     dict(size=12.0, weight="r",  color="ink", line=1.28),
    "rail_label":  dict(size=9.75, weight="sb", color="muted"),
    "rail_num":    dict(size=27.0, weight="b",  color="navy", spc=-0.03),
    "rail_cap":    dict(size=9.0,  weight="r",  color="ink_mid", line=1.25),
    "toc_num":     dict(size=33.75, weight="b", color="navy_30", spc=-0.02),
    "toc_part":    dict(size=17.25, weight="sb", color="navy", line=1.20),
    "toc_title":   dict(size=17.25, weight="r",  color="ink", line=1.20),
    "step_label":  dict(size=13.5, weight="sb", color="navy", line=1.20),
    "step_desc":   dict(size=12.75, weight="r",  color="ink", line=1.30),
    "thanks":      dict(size=45.0, weight="b",  color="ink", spc=-0.025, line=1.14),
    "closing_line": dict(size=14.25, weight="sb", color="amber", line=1.32),
}

# ── 세로 그리드 ───────────────────────────────────────────────────────────
KICKER_Y = 0.52
TITLE_Y, TITLE_H = 0.80, 0.94
RULE_Y, RULE_W, RULE_H = 1.66, 0.62, 0.05
CHIP_Y, CHIP_H = 1.88, 0.86
BODY_TOP = 2.95          # 콘텐츠 상단 기준선
BODY_BOTTOM = 6.62       # 콘텐츠 하단 한계
FOOT_LINE_Y = 6.90       # 러닝 헤어라인
FOOT_TEXT_Y = 6.98

# 헤더 '좁은 판' — 리드 아래 보조 한 줄(note)이 붙을 때로 전환한다.
# 1차 정답지 실측(cm→in): 넉넉한 판 2.04/4.21/5.03 · 좁은 판 1.69~1.85/3.33/3.86.
# note 를 넉넉한 판에 그대로 붙이면 본문 시작선(2.95)을 침범한다.
TITLE_Y_T, RULE_Y_T, CHIP_Y_T = 0.665, 1.311, 1.520
NOTE_H = 0.26            # 리드 아래 보조 한 줄
SUBHEAD_H = 0.31         # 본문 블록 소제목 (0.78cm)
SOURCE_H = 0.22          # 단위·출처 캡션 (0.55cm)
BLOCK_GAP = 0.10         # 소제목·캡션·본문 사이 간격

# 표준 시각 반경(인치)
R_CARD = 0.10
R_PHOTO = 0.10

# ── 슬라이드에 인쇄되는 고정 문구 ─────────────────────────────────────────
# 이 엔진은 강의 도구에서 왔지만 AIM 은 사업 보고서를 만들고, 영문 덱도 낸다.
# 그래서 '학습목표'·'형성평가' 같은 강의 용어를 코드에 박아 두지 않는다.
# deck13 어댑터가 덱 언어에 맞춰 갈아끼운다(set_labels).
LABELS = {
    "answer": "정답",
    "key": "핵심",
    "thanks": "감사합니다.",
    "agenda": "목차",
    "closing": ("요약", "다음 단계", "후속 조치"),
}

_LABELS_KO = {"answer": "정답", "key": "핵심", "thanks": "감사합니다.",
              "agenda": "목차",
              "closing": ("요약", "다음 단계", "후속 조치")}
_LABELS_EN = {"answer": "Answer", "key": "KEY", "thanks": "Thank you.",
              "agenda": "CONTENTS",
              "closing": ("Summary", "Next Steps", "Follow-up")}


def set_labels(lang: str = "ko") -> None:
    """덱 언어에 맞는 고정 문구로 바꾼다. 'en' 외에는 한국어."""
    LABELS.clear()
    LABELS.update(_LABELS_EN if str(lang).lower().startswith("en") else _LABELS_KO)

# 현재 빌드의 폰트 세트(build_deck 진입 시 설정)
_FS = pptx_font.SYSTEM_SET

# 견본 PPTX 에서 읽어 온 글꼴·견본 이름 (palette.apply 가 채운다)
TEMPLATE_FONT: Optional[str] = None
TEMPLATE_NAME: str = "기본"


# ── 넘침 보고 ─────────────────────────────────────────────────────────────
# 슬라이드는 문서와 다르다. 한글·워드는 넘치면 다음 장으로 흘려보내면 그만이지만,
# 슬라이드는 **줄일지 나눌지가 내용에 따라 달라진다** — 표 한 줄이 더 붙는 게 나은
# 장이 있고, 두 장으로 갈라야 하는 장이 있다. 그건 사람이 봐야 안다.
#
# 그래서 여기서는 **아무것도 버리지 않는다.** 넘치면 넘친 채로 두고(눈에 보여야
# 사람이 판단한다) 어느 장이 얼마나 넘쳤는지만 적어 올린다. 그 목록이 화면의
# 경고로 뜨고, 사용자가 그 장 번호를 채팅에 적어 '나눠줘'/'줄여줘' 하면 된다.
#
# 스레드별로 따로 담는다 — 잡 두 개가 동시에 렌더되면 보고가 섞인다.
_TL = threading.local()


def _note(kind: str, label: str, detail: str) -> None:
    rep = getattr(_TL, "report", None)
    if rep is None:
        return
    cur = getattr(_TL, "slide", None) or {}
    rep.append({"slide": cur.get("no"), "title": cur.get("title", ""),
                "kind": kind, "label": label, "detail": detail})


def _log(msg: str) -> None:
    """콘솔 코드페이지(cp949 등)에 없는 글자로 로그가 죽지 않게 한다.

    렌더 함수 안에서 print 가 UnicodeEncodeError 를 던지면 build_deck 의
    try/except 가 이를 렌더 실패로 오인해 폴백시킨다 — 실제로 겪은 문제.
    """
    try:
        print(msg, flush=True, file=__import__("sys").stderr)
    except Exception:  # noqa: BLE001
        try:
            enc = getattr(getattr(__import__("sys"), "stdout", None), "encoding",
                          None) or "ascii"
            print(msg.encode(enc, "replace").decode(enc, "replace"), flush=True, file=__import__("sys").stderr)
        except Exception:  # noqa: BLE001
            pass


def _rgb(name):
    from pptx.dml.color import RGBColor
    return RGBColor(*_HEX[name])


def _hexstr(name: str) -> str:
    return "%02X%02X%02X" % _HEX[name]


def luminance(name: str) -> float:
    """상대 휘도(0~1). WCAG 식 — 감마 보정을 해야 사람 눈과 맞는다."""
    def lin(c: float) -> float:
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _HEX.get(name, (0, 0, 0))
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def band(name: str, min_ratio: float = 4.5) -> str:
    """글자를 얹을 밴드의 채움색 — 대비가 모자라면 `<name>_dk` 로 바꾼다.

    템플릿마다 주색의 밝기가 다르다. 기본 배색의 navy(#1E2761)는 흰글씨와
    13.8:1 이라 그대로 써도 되지만, 리디자인의 navy(#2E93D9)는 3.3:1 이고
    amber(#4CB782)는 2.5:1 로 어떤 글자색을 얹어도 안 읽힌다.

    그래서 **모자랄 때만** 진한 짝을 쓴다. 넉넉한 템플릿은 손대지 않으므로
    기존 덱의 외형이 바뀌지 않는다.
    """
    def _best(bg: str) -> float:
        lb = luminance(bg)
        def r(a, b):
            hi, lo = max(a, b), min(a, b)
            return (hi + 0.05) / (lo + 0.05)
        return max(r(luminance("white"), lb), r(luminance("ink"), lb))

    if _best(name) >= min_ratio:
        return name
    dk = f"{name}_dk"
    if dk in _HEX and _best(dk) > _best(name):
        return dk
    return name


def on_white(name: str, min_ratio: float = 4.5) -> str:
    """**흰 바탕 위 글자색** — 대비가 모자라면 `<name>_dk`, 그래도 안 되면 `ink`.

    band() 와 헷갈리면 안 된다. band() 는 *배경* 을 고르고, 이건 *글자* 를 고른다.
    칩이 배경(연한 알약)을 잃고 흰 바탕에 놓이면서 필요해졌다 —
    리디자인의 navy(#2E93D9)는 연한 칩 위에서는 괜찮았지만
    흰 바탕에서는 3.34:1 로 떨어진다.
    """
    def ratio(a: float, b: float) -> float:
        hi, lo = max(a, b), min(a, b)
        return (hi + 0.05) / (lo + 0.05)

    w = luminance("white")
    for cand in (name, f"{name}_dk", "ink"):
        try:
            if ratio(luminance(cand), w) >= min_ratio:
                return cand
        except Exception:  # noqa: BLE001 — 템플릿에 없는 키
            continue
    return "ink"


def text_on(bg: str, *cands: str, min_ratio: float = 4.5) -> str:
    """`bg` 위에서 min_ratio 를 **처음 넘는** 후보. 아무도 못 넘으면 대비가 가장 큰 것.

    on_white() 는 흰 바탕 전용이고 on_color() 는 흰/ink 둘 중 하나만 고른다.
    강조색을 먼저 써 보고 모자랄 때만 포기하려면 후보 목록이 필요하다 —
    요약줄('핵심')이 그렇다. 시안 배색의 연한 워시(#E3F3F6) 위에서
    amber_dk 는 3.97:1, navy 는 3.13:1 로 둘 다 모자랐다.
    """
    def ratio(name: str) -> float:
        la, lb = luminance(name), luminance(bg)
        hi, lo = max(la, lb), min(la, lb)
        return (hi + 0.05) / (lo + 0.05)

    ok = [c for c in cands if c in _HEX]
    if not ok:
        return "ink"
    for c in ok:
        if ratio(c) >= min_ratio:
            return c
    return max(ok, key=ratio)


def on_color(bg: str, light: str = "white", dark: str = "ink") -> str:
    """배경 위에 얹을 글자색 — 명암비가 큰 쪽을 고른다.

    색을 박아 두면 템플릿에서 배경색을 바꿀 때 대비가 무너진다. 실제로 앰버
    밴드에는 어두운 글씨(6.9:1)가 맞고 흰 글씨(2.1:1)는 안 읽히는데, 초록 밴드는
    그 반대다. 그래서 고정값 대신 계산한다.

    ★ 어두운 쪽 기본값은 `navy` 가 아니라 **`ink`(본문 글자색)** 다. `navy` 는
      템플릿에 따라 진한 남색일 수도, 중간톤 블루일 수도 있다 — 리디자인 배색에서
      navy=#2E93D9 였고, 그걸 글자색으로 쓰면 연한 배경에서 2.8:1 로 떨어졌다.
      `ink` 는 어느 템플릿에서든 '가장 읽히는 글자색' 이라는 뜻이 유지된다.
    """
    def ratio(a: float, b: float) -> float:
        hi, lo = max(a, b), min(a, b)
        return (hi + 0.05) / (lo + 0.05)
    lb = luminance(bg)
    return light if ratio(luminance(light), lb) >= ratio(luminance(dark), lb) else dark


# ── 저수준 헬퍼 ───────────────────────────────────────────────────────────
def _remove_all_slides(prs) -> None:
    from pptx.oxml.ns import qn
    part = prs.part
    id_lst = prs.slides._sldIdLst
    for sid in list(id_lst):
        rid = sid.get(qn("r:id"))
        if rid and rid in part.rels:
            try:
                part.drop_rel(rid)
            except Exception:  # noqa: BLE001
                pass
        id_lst.remove(sid)


def _blank_layout(prs):
    for lay in prs.slide_layouts:
        nm = (lay.name or "").lower()
        if "blank" in nm or "빈" in (lay.name or ""):
            return lay
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


def _strip_placeholders(slide) -> None:
    """레이아웃에서 상속된 빈 플레이스홀더 제거(마스터 로고·장식은 유지)."""
    for ph in list(slide.placeholders):
        try:
            ph._element.getparent().remove(ph._element)
        except Exception:  # noqa: BLE001
            pass


def _no_deco(shape):
    """도형 기본 그림자·테두리 제거."""
    try:
        shape.line.fill.background()
    except Exception:  # noqa: BLE001
        pass
    try:
        shape.shadow.inherit = False
    except Exception:  # noqa: BLE001
        pass


def _adj(radius_in: float, w: float, h: float) -> int:
    """시각 반경(inch)을 roundRect adj 값으로. 도형 크기가 달라도 라운드가 같아진다."""
    base = max(min(w, h), 0.01)
    return int(max(0.0, min(radius_in / base, 0.5)) * 100000)


def _set_adj(shape, val: int) -> None:
    from pptx.oxml.ns import qn
    geom = shape._element.spPr.find(qn("a:prstGeom"))
    if geom is None:
        return
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = geom.makeelement(qn("a:avLst"), {})
        geom.append(av)
    for gd in list(av):
        av.remove(gd)
    gd = av.makeelement(qn("a:gd"), {"name": "adj", "fmla": f"val {val}"})
    av.append(gd)


def _shadow(shape, *, blur=0.14, dist=0.035, alpha=12, color="navy") -> None:
    """부드러운 아래 방향 그림자. 사진·카드·비교 컬럼에만 쓴다."""
    from pptx.oxml.ns import qn
    try:
        spPr = shape._element.spPr
        for el in spPr.findall(qn("a:effectLst")):
            spPr.remove(el)
        eff = spPr.makeelement(qn("a:effectLst"), {})
        shd = eff.makeelement(qn("a:outerShdw"), {
            "blurRad": str(int(blur * 914400)),
            "dist": str(int(dist * 914400)),
            "dir": "5400000", "rotWithShape": "0"})
        clr = shd.makeelement(qn("a:srgbClr"), {"val": _hexstr(color)})
        clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 1000))}))
        shd.append(clr)
        eff.append(shd)
        spPr.append(eff)
    except Exception:  # noqa: BLE001
        pass


def _rect(slide, x, y, w, h, fill):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    _no_deco(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    return shp


def _rrect(slide, x, y, w, h, fill=None, radius=R_CARD, line=None, line_w=None,
           pill=False):
    """라운드 사각형. radius 는 '인치 단위 시각 반경'(비율 아님). pill=True 면 알약."""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    _no_deco(shp)
    if fill is not None:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
    else:
        shp.fill.background()
    if line is not None:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_w or 1.0)
    _set_adj(shp, 50000 if pill else _adj(radius, w, h))
    return shp


def _oval(slide, x, y, d, fill):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                 Inches(d), Inches(d))
    _no_deco(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    return shp


def _arrow(slide, x, y, w, h, fill="amber"):
    """단계·순서 연결자."""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y),
                                 Inches(w), Inches(h))
    _no_deco(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    return shp


def _text(slide, x, y, w, h, anchor="t"):
    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE,
                          "b": MSO_ANCHOR.BOTTOM}.get(anchor, MSO_ANCHOR.TOP)
    return tf


def _para(tf, first, align=None, space_after=4, line=1.05):
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if align:
        p.alignment = {"c": PP_ALIGN.CENTER, "l": PP_ALIGN.LEFT,
                       "r": PP_ALIGN.RIGHT}[align]
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    try:
        p.line_spacing = line
    except Exception:  # noqa: BLE001
        pass
    return p


def _run(p, text, size, color, weight="r", spc=0.0):
    from pptx.util import Pt
    r = p.add_run()
    r.text = str(text)
    r.font.size = Pt(size)
    r.font.bold = _FS.is_bold(weight)
    r.font.color.rgb = _rgb(color)
    pptx_font.apply_face(r, _FS.face(weight), spc_em=spc, size_pt=size)
    return r


def _T(p, text, role, *, size=None, color=None, weight=None, spc=None):
    """TYPE 롤로 런 추가. 개별 인자로 덮어쓸 수 있다."""
    t = TYPE[role]
    return _run(p, text,
                t["size"] if size is None else size,
                t["color"] if color is None else color,
                t["weight"] if weight is None else weight,
                t.get("spc", 0.0) if spc is None else spc)


_MARK = re.compile(r"\*([^*]+)\*")


def _T_split(p, text, role, *, size=None, accent=None):
    """제목에 **2색 강조**를 넣는다 — `*키워드*` 로 감싼 부분만 주색.

    회사 양식의 시그니처다(DESIGN.md §5): 제목 문장은 본문색이고 핵심 키워드만
    주색으로 뜬다. 마크가 없으면 통째로 한 런이라 예전 동작과 같다.
    """
    parts = _MARK.split(str(text or ""))
    if len(parts) == 1:
        return _T(p, text, role, size=size)
    base = TYPE[role]["color"]
    acc = accent or on_white("navy")
    for i, seg in enumerate(parts):
        if seg:
            _T(p, seg, role, size=size, color=(acc if i % 2 else base))


def _demark(text) -> str:
    """강조 마크를 뗀 순수 문자열 — 폭 계산·보고용."""
    return _MARK.sub(r"\1", str(text or ""))


def _hang(p, marL_in=0.24):
    """내어쓰기 — ▸ 는 왼쪽으로 튀고 줄바꿈 글자는 텍스트 아래로 정렬."""
    emu = str(int(marL_in * 914400))
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", emu)
    pPr.set("indent", "-" + emu)


# ── 텍스트 오버플로 보호 ───────────────────────────────────────────────────
def _measure(text: str, pt: float) -> float:
    """문자열의 근사 렌더 폭(inch). 한글·전각 1.0em, 라틴 0.54em, 공백 0.28em."""
    em = pt / 72.0
    w = 0.0
    for ch in str(text):
        o = ord(ch)
        if ch == " ":
            w += 0.28
        elif (0x1100 <= o <= 0x11FF or 0x2E80 <= o <= 0xA4CF
              or 0xAC00 <= o <= 0xD7A3 or 0xF900 <= o <= 0xFAFF
              or 0xFF00 <= o <= 0xFF60):
            w += 1.0
        elif ch.isdigit() or ch.isupper():
            w += 0.58
        else:
            w += 0.52
    return w * em


def _needed_h(lines: List[str], box_w: float, pt: float, line: float,
              space_after: float) -> float:
    """문단들이 차지할 높이(inch) 추정."""
    usable = max(box_w - 0.12, 0.4)
    rows = 0
    for t in lines:
        rows += max(1, math.ceil(_measure(t, pt) / usable - 1e-6))
    return rows * (pt / 72.0) * line + len(lines) * (space_after / 72.0)


def _fit(lines: List[str], box_w: float, box_h: float, base_pt: float,
         line: float = 1.4, space_after: float = 0.0,
         min_pt: float = 10.5, label: str = "") -> float:
    """박스에 들어가는 최대 폰트 크기(0.5pt 단위). 최소치에서도 넘치면 경고."""
    if not lines:
        return base_pt
    # 추정이 낙관적이라 '들어간다'고 본 것이 실제로는 마지막 줄이 잘리는 일이
    # 있었다(마무리 카드 8줄에서 실측). 4% 여유를 두어 보수적으로 맞춘다.
    box_h = box_h * 0.96
    pt = base_pt
    while pt > min_pt:
        if _needed_h(lines, box_w, pt, line, space_after) <= box_h:
            return pt
        pt -= 0.5
    need = _needed_h(lines, box_w, min_pt, line, space_after)
    if need > box_h and label:
        over = need - box_h
        rows = max(1, int(round(over / max(0.01, (min_pt / 72.0) * line))))
        _log(f"[deck] 넘침: {label} — {min_pt}pt 로도 {over:.2f}in 초과")
        _note("overflow", label, f"{rows}줄 넘침")
    return min_pt


def _fit_each(items: List[str], box_w: float, box_h: float, base_pt: float,
              line: float = 1.3, min_pt: float = 10.5, label: str = "") -> float:
    """항목마다 '자기 박스'를 갖는 레이아웃용 — 가장 빡빡한 항목에 맞춘 공통 크기.

    _fit() 은 여러 문단이 한 박스에 쌓이는 경우(불릿 목록)를 가정하므로,
    목차·보기·표 셀처럼 항목별 박스가 따로인 곳에 쓰면 높이를 과대 계산한다.
    """
    vals = [str(t) for t in items if str(t).strip()]
    if not vals:
        return base_pt
    return min(_fit([v], box_w, box_h, base_pt, line, 0.0, min_pt, label)
               for v in vals)


def _trim_to_fit(lines: List[str], box_w: float, box_h: float, pt: float,
                 line: float, space_after: float, label: str = "") -> List[str]:
    """넘쳐도 **버리지 않는다.** 몇 항목이 넘치는지만 세어 보고한다.

    예전에는 뒤에서부터 덜어냈다. 그런데 마지막 항목이 그 장의 결론인 경우가
    있어서, 조용히 사라지면 작성자가 알 길이 없다. 지금은 넘친 채로 두고
    (화면에서 눈에 띈다) 어느 장인지 보고해 사람이 나눌지 줄일지 정하게 한다.
    """
    n = len(lines)
    while n > 1 and _needed_h(lines[:n], box_w, pt, line, space_after) > box_h:
        n -= 1
    if n < len(lines):
        _log(f"[deck] 넘침: {label} — {len(lines) - n}개 항목이 상자 밖")
        _note("overflow", label, f"항목 {len(lines) - n}개 넘침")
    return list(lines)


def _normautofit(tf, scale: float = 1.0, lnspc_red: float = 0.0) -> None:
    """<a:normAutofit> 주입 — PowerPoint 가 편집 시 다시 맞추도록 하는 보조 장치."""
    from pptx.oxml.ns import qn
    try:
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        if bodyPr is None:
            return
        for tag in ("a:normAutofit", "a:spAutoFit", "a:noAutofit"):
            el = bodyPr.find(qn(tag))
            if el is not None:
                bodyPr.remove(el)
        naf = bodyPr.makeelement(qn("a:normAutofit"), {})
        if scale < 1.0:
            naf.set("fontScale", str(int(scale * 100000)))
        if lnspc_red > 0:
            naf.set("lnSpcReduction", str(int(lnspc_red * 100000)))
        bodyPr.append(naf)
    except Exception:  # noqa: BLE001
        pass


# ── 구성 요소 ─────────────────────────────────────────────────────────────
def add_kicker(slide, text, x=MARGIN, y=KICKER_Y, w=8.0, color=None,
               in_logo_band=False):
    """상단 라벨.

    ★ 좌상단 로고가 있으면 **로고 띠 안에서 세로 중앙을 맞춘다.**
      기본 y(0.52)는 로고가 없던 시절의 값이라, 공식 양식의 로고 띠
      (0.21~0.55)와 겹치거나 바로 아래에 붙어 '왼쪽 로고의 부제'처럼 읽혔다.
      로고와 같은 줄에 놓으면 [로고] 라벨 ⋯ [로고] 로 한 줄이 된다.
    """
    if not text:
        return
    if in_logo_band:
        tf = _text(slide, x, LOGO_Y, w, LOGO_H, anchor="m")
    else:
        tf = _text(slide, x, y, w, 0.24, anchor="m")
    _T(_para(tf, True), str(text).upper(), "kicker", color=color)


def add_title(slide, text, *, has_logo=False, y=TITLE_Y, size=None, h=None):
    """제목. `h` 는 **룰 윗선까지의 높이** — 안 주면 넉넉한 판 기준.

    ★ 높이를 사다리에서 받아야 한다. 예전에는 TITLE_H(0.94)를 고정으로 썼는데,
      좁은 판(제목 0.665 · 룰 1.311)에서는 상자가 룰을 덮어 룰이 **첫 낱말의
      밑줄**처럼 보였다(가운데 정렬이라 글자가 룰 위로 내려앉는다).
    """
    w = 10.6 if has_logo else CONTENT_W
    box_h = TITLE_H if h is None else h
    t = TYPE["title"]
    plain = _demark(text)
    pt = size or _fit([plain], w, box_h, t["size"], t["line"],
                      min_pt=19, label=f"title:{plain[:18]}")
    tf = _text(slide, MARGIN, y, w, box_h, anchor="m")
    _T_split(_para(tf, True, line=t["line"]), text or "", "title", size=pt)


def add_rule(slide, y=RULE_Y, x=MARGIN, w=RULE_W, color="amber"):
    _rrect(slide, x, y, w, RULE_H, fill=color, pill=True)


def add_chip(slide, text, x=MARGIN, y=CHIP_Y, w=CONTENT_W, emphasis=False):
    """한 줄 요약. **글꼴로만** 쓴다 — 반환 = 하단 y.

    예전에는 라운드 배경 + 앰버 점 + 한 문장이었다. 뺀 이유:
      · 알약이 슬라이드마다 반복되니 제목보다 먼저 눈에 들어와 위계가 뒤집혔다
      · 배경색이 사진·도해와 부딪혀 화면이 시끄러웠다
    강조는 색을 채우지 않고 **더 진한 청색**으로 한다.

    ★ 높이(CHIP_H)는 그대로 둔다. 이 함수가 돌려주는 y 를 13종 레이아웃이
      본문 시작점으로 쓰기 때문에, 줄이면 전 레이아웃의 배치가 한꺼번에 틀어진다.
      배경이 없어진 만큼 위아래 여백이 되어 소제목처럼 읽힌다.
    """
    if not text:
        return y
    t = TYPE["chip"]
    # 강조는 색이 아니라 **크기**로 준다. 기본 배색은 navy 가 이미 13.8:1 이라
    # 더 진한 단계가 없어서, 색만으로는 두 템플릿 중 한쪽에서 차이가 안 난다.
    base = t["size"] + (1.5 if emphasis else 0.0)
    pt = _fit([text], w, CHIP_H - 0.16, base, t["line"],
              min_pt=11.5, label=f"chip:{str(text)[:18]}")
    tf = _text(slide, x, y, w, CHIP_H, anchor="m")
    _T(_para(tf, True, line=t["line"]), text, "chip", size=pt,
       color=on_white("navy"))
    return y + CHIP_H


def add_note(slide, text, y, x=MARGIN, w=CONTENT_W):
    """리드 아래 **보조 한 줄**. 반환 = 하단 y(글이 없으면 y 그대로).

    1차 정답지 3종 전부가 리드(굵은 그린 한 줄) 아래에 옅은 회색 한 줄을 더 둔다.
    근거·범위·판정 유의점처럼 결론을 뒷받침하지만 결론은 아닌 문장이다. 이 자리가
    없으면 그 내용이 리드에 섞여 리드가 두 줄이 되고, 그러면 제목과 위계가 붙는다.
    """
    if not text:
        return y
    t = TYPE["note"]
    pt = _fit([str(text)], w, NOTE_H, t["size"], t["line"], min_pt=9,
              label=f"note:{str(text)[:18]}")
    tf = _text(slide, x, y, w, NOTE_H, anchor="m")
    _T(_para(tf, True, line=t["line"]), str(text), "note", size=pt)
    return y + NOTE_H


def add_subhead(slide, text, y, x=MARGIN, w=CONTENT_W):
    """본문 블록의 소제목(표·차트 위에 붙는 이름). 반환 = 하단 y.

    긴 이름은 **두 줄까지 자란다** — 한 줄에 억지로 맞추면 10.5pt 까지 내려가
    아래 캡션(9pt)과 위계가 붙는다(정답지 0079 S8 의 소제목도 두 줄이다).
    """
    if not text:
        return y
    t = TYPE["subhead"]
    h = SUBHEAD_H if _measure(str(text), t["size"]) <= (w - 0.14) else SUBHEAD_H * 2
    pt = _fit([str(text)], w, h, t["size"], t["line"], min_pt=10.5,
              label=f"subhead:{str(text)[:18]}")
    tf = _text(slide, x, y, w, h, anchor="m")
    _T(_para(tf, True, line=t["line"]), str(text), "subhead", size=pt)
    return y + h


def add_source(slide, text, y, x=MARGIN, w=CONTENT_W):
    """단위·출처 캡션. 반환 = 하단 y.

    성과관리 보고는 수치의 **출처 파일명**을 화면에 남겨야 검증이 된다.
    소제목 바로 아래, 표·차트 바로 위가 그 자리다(정답지 3종 공통).
    """
    if not text:
        return y
    tf = _text(slide, x, y, w, SOURCE_H, anchor="m")
    _T(_para(tf, True), str(text), "source")
    return y + SOURCE_H


def add_bullets(slide, bullets, x, y, w, h, *, role="bullet", marker=True,
                label="", color=None):
    if not bullets:
        return
    t = TYPE[role]
    items = [str(b) for b in bullets if str(b).strip()]
    if not items:
        return
    sa = 7.0
    # 마커 폭(▸ + 공백)만큼 텍스트 폭이 줄어든다
    eff_w = w - (0.26 if marker else 0.0)
    pt = _fit(items, eff_w, h, t["size"], t["line"], sa, min_pt=11, label=label)
    items = _trim_to_fit(items, eff_w, h, pt, t["line"], sa, label=label)
    tf = _text(slide, x, y, w, h, anchor="t")
    for i, b in enumerate(items):
        p = _para(tf, i == 0, space_after=sa, line=t["line"])
        if marker:
            _hang(p)
            _run(p, "▸ ", pt, "amber", "sb")
        _T(p, b, role, size=pt, color=color)
    _normautofit(tf, scale=min(1.0, pt / t["size"]))


def add_photo_slot(slide, x, y, w, h):
    """사진이 아직 없을 때의 **빈 액자**. 자리를 확정해 배치가 두 번 바뀌지 않게 한다.

    글자를 넣지 않는다 — 사진 수집을 건너뛴 덱이 그대로 쓰일 수도 있고, 그때
    '사진 자리' 같은 문구가 남으면 미완성으로 보인다. 대신 옅은 워시와 헤어라인,
    가운데 작은 표식으로 '여기 그림이 들어갈 자리' 를 알린다.
    """
    box = _rrect(slide, x, y, w, h, fill="chip", radius=R_PHOTO,
                 line="navy_12", line_w=1.25)
    # 가운데 작은 원 — 도형이므로 어느 배색에서도 튀지 않는다
    d = min(w, h) * 0.11
    _oval(slide, x + (w - d) / 2, y + (h - d) / 2, d, "navy_30")
    return box


def add_photo(slide, data, x, y, w, h, *, credit=None):
    from pptx.util import Inches
    from pptx.oxml.ns import qn
    try:
        pic = slide.shapes.add_picture(io.BytesIO(data), Inches(x), Inches(y),
                                       Inches(w), Inches(h))
    except Exception:  # noqa: BLE001
        return None
    # 종횡비 맞춰 센터 크롭(가능하면)
    try:
        from PIL import Image
        iw, ih = Image.open(io.BytesIO(data)).size
        tgt, src = w / h, iw / ih
        if src > tgt:
            c = (1 - tgt / src) / 2
            pic.crop_left = c
            pic.crop_right = c
        elif src < tgt:
            c = (1 - src / tgt) / 2
            pic.crop_top = c
            pic.crop_bottom = c
    except Exception:  # noqa: BLE001
        pass
    geom = pic._element.spPr.find(qn("a:prstGeom"))
    if geom is not None:
        geom.set("prst", "roundRect")
        _set_adj(pic, _adj(R_PHOTO, w, h))
    _shadow(pic, blur=0.16, dist=0.04, alpha=14)
    if credit:
        tf = _text(slide, x, y + h + 0.04, w, 0.22, anchor="t")
        _T(_para(tf, True, align="r"), str(credit)[:90], "caption")
    return pic


INTERP_ROW = 0.27          # 해석 밴드 한 줄 높이 (정답지 실측 0.68cm)
INTERP_BAR_W = 0.083       # 좌측 액센트 바 (0.21cm)


def interp_h(n_lines: int) -> float:
    """해석 밴드가 n줄을 담을 때의 높이. 배치 계산에 미리 쓴다."""
    if n_lines <= 0:
        return 0.0
    if n_lines == 1:
        return 0.60                      # 라벨을 같은 줄에 얹는 판
    return 0.34 + INTERP_ROW * n_lines + 0.08


def add_interpretation(slide, lines, *, label=None, y=None):
    """하단 **해석 밴드**. 반환 = 밴드 높이(0 이면 안 그렸다).

    한 줄이면 예전 '핵심' 띠 그대로(라벨을 같은 줄에 얹는다). 두 줄 이상이면
    라벨을 위로 올리고 불릿으로 쌓는다 — 정답지의 판정 근거는 대개 두세 문장이고,
    한 줄에 이어 붙이면 글꼴이 줄어들어 정작 결론이 안 읽혔다.

    채움은 `amber_wash`(강조 워시 슬롯) 다. 회사 양식 배색에서는 정답지가 쓴
    #E3F5EA 와 사실상 같은 색이고, 다른 배색으로 갈아도 '강조' 의미가 유지된다.
    """
    lines = [str(l) for l in (lines or []) if str(l).strip()]
    if not lines:
        return 0.0
    lbl = label or LABELS.get("key", "핵심")
    h = interp_h(len(lines))
    y = (BODY_BOTTOM - h) if y is None else y
    bg = "amber_wash"
    _rrect(slide, MARGIN, y, CONTENT_W, h, fill=bg, radius=R_CARD)
    _rrect(slide, MARGIN, y, INTERP_BAR_W, h, fill="amber", radius=0.04)
    tx = MARGIN + 0.34
    tw = CONTENT_W - 0.60
    col = text_on(bg, "navy_dk", "navy", "ink")

    if len(lines) == 1:
        tf = _text(slide, tx, y, tw, h, anchor="m")
        p = _para(tf, True)
        _run(p, lbl + "   ", 12.5, text_on(bg, "amber_dk", "navy_dk", "ink"), "sb")
        _run(p, lines[0], 13.5, col, "sb")
        return h

    tfl = _text(slide, tx, y + 0.06, tw, 0.27, anchor="m")
    _T(_para(tfl, True), lbl, "interp_lbl",
       color=text_on(bg, "amber_dk", "navy_dk", "ink"))
    t = TYPE["interp"]
    pt = _fit_each(lines, tw - 0.24, INTERP_ROW, t["size"], t["line"],
                  min_pt=9.5, label="interp")
    tf = _text(slide, tx, y + 0.33, tw, INTERP_ROW * len(lines) + 0.06, anchor="t")
    for i, ln in enumerate(lines):
        pp = _para(tf, i == 0, space_after=2.0, line=t["line"])
        _hang(pp, 0.20)
        _run(pp, "· ", pt, col, "sb")
        _T(pp, ln, "interp", size=pt, color=col)
    return h


def _body_bottom(s, gap=0.22):
    """해석 밴드를 뺀 본문 하단. 밴드가 없으면 BODY_BOTTOM.

    예전에는 렌더러마다 `BODY_BOTTOM - 0.82` 처럼 밴드 높이를 박아 두었다.
    밴드가 여러 줄로 늘어날 수 있게 되었으니 한 곳에서 계산한다.
    """
    lines = _interp_of(s)
    return BODY_BOTTOM - interp_h(len(lines)) - gap if lines else BODY_BOTTOM


RAIL_W = 4.126           # 우측 수치 레일 카드 폭 (10.48cm)
RAIL_X = CONTENT_R - RAIL_W
RAIL_GAP = 0.18          # 본문 열과 레일 사이


def add_stat_rail(slide, items, top, bottom):
    """우측 **수치 레일** — 라벨 / 큰 수치 / 캡션 카드를 세로로 쌓는다.

    1차 정답지 0083 S47 의 판이다. 왼쪽에 서술을 두고 오른쪽에 근거 수치를
    세우면 한 장에서 '무슨 말인지' 와 '숫자가 얼마인지' 를 같이 보여줄 수 있다.
    stat 슬라이드(큰 카드 3개)와는 쓰임이 다르다 — 이건 본문의 곁가지다.

    items = [{"label": 이름, "value": 수치, "desc": 캡션}]  — value 를 빼면
    label 을 수치로 읽는다(stat 항목을 그대로 넘겨도 그려진다).
    """
    items = [it for it in (items or []) if it][:4]
    if not items:
        return
    n = len(items)
    pitch = (bottom - top) / n
    card_h = min(1.272, max(0.86, pitch - 0.165))
    nt = TYPE["rail_num"]
    nums, names = [], []
    for it in items:
        v = it.get("value") if it.get("value") is not None else it.get("num")
        nums.append(str(v if v is not None else it.get("label", "")))
        names.append(str(it.get("label", "")) if v is not None else "")
    npt = min(_fit([v], RAIL_W - 0.9, 0.62, nt["size"], 1.1, min_pt=15,
                   label="rail num") for v in nums)
    for i, it in enumerate(items):
        cy = top + i * pitch
        _rrect(slide, RAIL_X, cy, RAIL_W, card_h, fill="on_navy", radius=R_CARD)
        # 카드 높이가 남는 자리에 따라 달라지므로 세로 배분을 계산한다.
        cap_h = 0.30 if it.get("desc") else 0.06
        num_h = max(0.40, card_h - 0.32 - cap_h)
        if names[i]:
            tfl = _text(slide, RAIL_X + 0.25, cy + 0.06, RAIL_W - 0.5, 0.26,
                        anchor="m")
            _T(_para(tfl, True), names[i], "rail_label")
        tfn = _text(slide, RAIL_X + 0.25, cy + 0.32, RAIL_W - 0.5, num_h,
                    anchor="m")
        _T(_para(tfn, True), nums[i], "rail_num", size=npt,
           color=on_white("navy"))
        if it.get("desc"):
            tfd = _text(slide, RAIL_X + 0.25, cy + card_h - cap_h,
                        RAIL_W - 0.5, 0.28, anchor="m")
            _T(_para(tfd, True), it["desc"], "rail_cap")


def _interp_of(s):
    """슬라이드에서 해석 밴드에 넣을 줄 목록. key/footer 한 줄도 여기로 모인다."""
    v = s.get("interpretation")
    if isinstance(v, str):
        v = [v]
    if v:
        return [str(x) for x in v if str(x).strip()]
    one = s.get("key") or s.get("footer")
    return [str(one)] if one and str(one).strip() else []


def add_footer_key(slide, text, y=None):
    """하단 '핵심' 한 줄 — 해석 밴드의 한 줄짜리 판."""
    if not text:
        return 0.0
    return add_interpretation(slide, [text], y=y)


# ── 상단 로고 띠 — 회사 공식 양식의 슬라이드마스터 실측 좌표 그대로 ──────────
# (_context/로고크기조정-회사양식PPT_마스터_260616.pptx)
# 사업측이 준 마스터를 **그대로** 따른다. 로고 이미지도 그 파일 안의 것을 쓴다.
#   왼쪽 로고    L0.59  T0.21 W1.33 H0.34   (그림)
#   UBION        L10.79 T0.15 W1.18 H0.40   (그림)
#   consortium   L11.90 T0.18 W1.17 H0.34   (텍스트 14pt #0079FF)
# ★ 그림을 자르지 않는다 — 원본의 흰 여백까지 포함해야 이 좌표에 정확히 앉는다.
LOGO_L, LOGO_Y, LOGO_LW, LOGO_H = 0.59, 0.21, 1.33, 0.34      # 왼쪽 로고
_BAND_MID = LOGO_Y + LOGO_H / 2                                # 0.38 — 띠의 세로 중심

# UBION 은 **워드마크만** 정식 로고다. 마스터의 'consortium' 은 정식 마크가 아니라
# 텍스트 글상자로 얹은 것이고, 글꼴을 Gentium Basic 으로 지정해 두었는데 그 글꼴이
# 설치돼 있지 않아 PC 마다 다른 글꼴로 대체돼 보인다. 납품물에 그 위험을 둘 이유가
# 없으므로 **워드마크는 그림, consortium 은 덱 글꼴 텍스트**로 우리가 배열한다.
LOGO_RX, LOGO_RH = 10.79, 0.28
LOGO_RW = LOGO_RH * 4.096                       # 워드마크 원본 비율(942x230)
LOGO_RY = _BAND_MID - LOGO_RH / 2               # 세로 중앙 정렬
CONSORTIUM = dict(x=LOGO_RX + LOGO_RW + 0.10,   # 워드마크 오른쪽 + 여백
                  h=0.34, w=1.03, pt=14.0, color="0079FF")
CONSORTIUM["y"] = _BAND_MID - CONSORTIUM["h"] / 2   # 워드마크와 같은 세로 중심
LOGO_R = CONSORTIUM["x"] + CONSORTIUM["w"]      # 헤더 오른쪽 끝 (마스터 13.07 근처)

# 표지 — 회사 양식 '00_표지' 레이아웃 실측
COVER_BAR_W = 0.42                # 좌측 진한 세로 바
COVER_BAR2_W = 0.09               # 그 옆 강조 세로 바
COVER_PANEL_X, COVER_PANEL_W = 9.55, 3.78   # 우측 연한 패널
COVER_TEXT_X = 1.15               # 표지 글자 시작 x
COVER_TEXT_W = COVER_PANEL_X - COVER_TEXT_X - 0.55   # 7.85 — 우측 패널을 안 침범
# 표지 세로 자리 — 1차 정답지 3종이 모두 같은 판을 쓴다(cm→in).
#   로고 1.59 · 제목 4.5 · 부제 8.47 · 룰 12.44 · 발표자 13.36 · 소속 14.23
# 예전 판(제목 2.35 · 부제 3.25 · 룰 3.88)은 제목에 두 줄밖에 못 줘서 사업명이
# 긴 덱에서 글꼴이 24pt 까지 내려갔다. 제목 자리를 위로 올려 세 줄을 준다.
COVER_LOGO_Y = 0.626
COVER_TITLE_Y = 1.772
COVER_SUB_Y = 3.335
COVER_RULE_Y = 4.898
COVER_RULE_W, COVER_RULE_H = 3.102, 0.043
COVER_PRESENTER_Y = 5.260
COVER_AFFIL_Y = 5.602
COVER_FOOT_Y = 6.550
# 우하단 로고 — 표지·마감 장에서는 상단 띠가 아니라 아래쪽에 놓인다.
COVER_LOGO_RX, COVER_LOGO_RY = 10.311, 6.665
COVER_TAB = (10.60, 1.15, 2.00, 0.18)   # 우측 패널 강조 바


def _cover_frame(slide):
    """표지·마감 장의 공통 바탕 — 좌측 세로 바 두 줄 + 우측 연한 패널 + 강조 바.

    회사 공식 양식 '00_표지' 실측이고 정답지 3종도 같다. 좌측 전면 밴드를 쓰던
    예전 판은 버렸다(밴드가 좌상단 로고를 덮어 첫 장에 헤더가 안 보였다).
    """
    _rect(slide, 0, 0, COVER_BAR_W, SLIDE_H, "navy")
    _rect(slide, COVER_BAR_W, 0, COVER_BAR2_W, SLIDE_H, "amber")
    _rect(slide, COVER_PANEL_X, 0, COVER_PANEL_W, SLIDE_H, "chip")
    _rect(slide, COVER_TAB[0], COVER_TAB[1], COVER_TAB[2], COVER_TAB[3], "amber")


def add_logos(slide, logo_path, logo2_path=None, *,
              left_x: float = LOGO_L, right_x: float = LOGO_R,
              left_y: Optional[float] = None,
              right_y: Optional[float] = None) -> float:
    """상단 로고. 로고1은 **우상단**, 로고2는 **좌상단**.

    반환 = 좌상단 로고가 차지한 폭(in). 0 이면 없다.
    ★ 이 값이 필요한 이유: 키커(0.52~0.76)가 로고 띠(0.28~0.70)와 겹친다.
      좌상단에 로고가 오면 키커를 그만큼 오른쪽으로 밀어야 글자가 안 겹친다.
      로고 폭은 그림 비율에 따라 다르므로 **넣어 보고 재는** 수밖에 없다.

    로고가 없으면 그 자리는 그냥 빈다 — 자리표시를 그리지 않는다.
    """
    from pptx.util import Inches
    left_w = 0.0
    dx = left_x - LOGO_L          # 표지 등에서 띠를 통째로 옮길 때의 이동량
    if logo2_path and os.path.isfile(logo2_path):
        try:
            slide.shapes.add_picture(logo2_path, Inches(LOGO_L + dx),
                                     Inches(LOGO_Y if left_y is None else left_y),
                                     Inches(LOGO_LW), Inches(LOGO_H))
            left_w = LOGO_LW
        except Exception:  # noqa: BLE001
            pass
    if logo_path and os.path.isfile(logo_path):
        rdx = right_x - LOGO_R
        rdy = 0.0 if right_y is None else (right_y - LOGO_RY)
        try:
            slide.shapes.add_picture(logo_path, Inches(LOGO_RX + rdx),
                                     Inches(LOGO_RY + rdy),
                                     Inches(LOGO_RW), Inches(LOGO_RH))
        except Exception:  # noqa: BLE001
            pass
        # 'consortium' 은 마스터에서도 그림이 아니라 텍스트다. 그대로 따른다.
        c = CONSORTIUM
        try:
            tf = _text(slide, c["x"] + rdx, c["y"] + rdy, c["w"], c["h"],
                       anchor="m")
            tf.margin_left = tf.margin_right = 0     # 시작점을 좌표로 직접 잡는다
            tf.word_wrap = False                     # 'consortiu/m' 로 접히던 것
            p = _para(tf, True)
            r = p.add_run()
            r.text = "consortium"
            r.font.size = Pt_(c["pt"])
            r.font.color.rgb = _rgbhex(c["color"])
            pptx_font.apply_face(r, _FS.face("r"), size_pt=c["pt"])
        except Exception:  # noqa: BLE001
            pass
    return left_w


def Pt_(v):
    from pptx.util import Pt
    return Pt(v)


def _rgbhex(h: str):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(h)


def add_running(slide, footer_text: str, page: Optional[int]):
    """러닝 요소 — 헤어라인 + 좌측 과목·주차 + 우측 페이지 번호."""
    if not footer_text and page is None:
        return
    _rect(slide, MARGIN, FOOT_LINE_Y, CONTENT_W, 0.01, "navy_12")
    if footer_text:
        tf = _text(slide, MARGIN, FOOT_TEXT_Y, 9.0, 0.28, anchor="m")
        _T(_para(tf, True), footer_text, "footer")
    if page is not None:
        tf = _text(slide, CONTENT_R - 1.2, FOOT_TEXT_Y, 1.2, 0.28, anchor="m")
        _T(_para(tf, True, align="r"), str(page), "pageno")


def _tight(s) -> bool:
    """좁은 헤더 판을 쓸지. **note 가 있으면 자동으로 좁힌다.**

    보조 한 줄을 넉넉한 판(리드 하단 2.74)에 그대로 붙이면 본문 시작선(2.95)을
    침범한다. 1차 정답지도 같은 선택을 했다 — 보조 줄이 있는 장은 제목을 위로
    올려 한 줄 몫을 만든다. 슬라이드에서 `tight` 를 명시하면 그 값이 이긴다.
    """
    if s.get("tight") is not None:
        return bool(s["tight"])
    return bool(str(s.get("note") or "").strip())


def _head(slide, s, *, has_logo=False, rule=True):
    """상단 존(키커 → 제목 → 룰) 공통 처리. 반환 = 요약줄 시작 y."""
    # 좌상단 로고가 있으면 키커를 그 오른쪽으로 밀고 로고와 같은 줄에 앉힌다.
    dx = float(s.get("_kicker_dx") or 0.0)
    add_kicker(slide, s.get("kicker"), x=MARGIN + dx, in_logo_band=bool(dx))
    t = _tight(s)
    ty, ry, cy = ((TITLE_Y_T, RULE_Y_T, CHIP_Y_T) if t
                  else (TITLE_Y, RULE_Y, CHIP_Y))
    add_title(slide, s.get("title", ""), has_logo=has_logo, y=ty,
              h=ry - ty - 0.02)          # 제목 상자는 룰 윗선에서 멈춘다
    if rule:
        add_rule(slide, y=ry)
    return cy


def _lead(slide, s, y, x=MARGIN, w=CONTENT_W):
    """리드(한 줄 요약) + 보조 한 줄을 이어 그린다. 반환 = 하단 y."""
    by = add_chip(slide, s.get("chip"), x=x, y=y, w=w,
                  emphasis=bool(s.get("emphasis")))
    return add_note(slide, s.get("note"), by, x=x, w=w)


# ── 사진 배치 프리셋 ───────────────────────────────────────────────────────
# img=(L,T,W,H), tx/tw=텍스트 열, below=이미지가 하단(본문은 위 전폭).
# 사진 슬라이드 '순번'으로 순환한다(절대 인덱스 아님) — 배치 반복 방지.
#
# 세로 기준(첨삭 2026-08-03): 옆단 사진은 **요약줄 윗선(1.88")에 top 을 맞추고**
# 콘텐츠 하단(6.57")까지 내린다. 본문 시작선(2.95")에서 시작하면 칩 옆이 비어
# 위쪽이 헐거워 보인다. 폭만 바꿔 변화를 준다(높이를 줄이면 정렬이 깨진다).
PHOTO_TOP = CHIP_Y          # 1.88 — 요약줄 윗선
PHOTO_BOT = 6.57            # 콘텐츠 하단(러닝 푸터 위)
_PH = PHOTO_BOT - PHOTO_TOP  # 4.69

# 그림 2단 — 정답지 1ch S6 실측: 흰 카드 15.13x12.7cm 를 0.215 간격으로 둘.
PHOTO2_TOP = 1.665
PHOTO2_GAP = 0.215


def _photo_2up(slide, s, by, img=None):
    """그림 두 장을 나란히. 사진이 없어도 **자리는 남긴다.**

    한 장에 그림 두 개를 놓는 판이지만 이미지 배관은 슬라이드당 한 장이다
    (build_deck 의 images[i]). 그래서 받은 한 장은 왼쪽에 넣고 오른쪽은 빈 액자로
    둔다 — 두 번째 그림 자리를 사람이 채우게 눈에 보이는 채로 남긴다.
    """
    top = max(min(by + 0.20, BODY_TOP), PHOTO2_TOP)
    bottom = _body_bottom(s)
    w = (CONTENT_W - PHOTO2_GAP) / 2
    caps = [str(c) for c in (s.get("captions") or [])]
    ch = 0.28 if caps else 0.0
    h = max(1.0, bottom - top - ch)
    for i in range(2):
        x = MARGIN + i * (w + PHOTO2_GAP)
        if i == 0 and img:
            add_photo(slide, img, x, top, w, h)
        else:
            add_photo_slot(slide, x, top, w, h)
        if i < len(caps):
            tf = _text(slide, x, top + h + 0.03, w, 0.26, anchor="m")
            _T(_para(tf, True, align="c"), caps[i], "caption")


PHOTO_PRESETS = [
    {"img": (8.55, PHOTO_TOP, 4.18, _PH), "tx": 0.60, "tw": 7.60},  # 우측 표준
    {"img": (0.60, PHOTO_TOP, 4.18, _PH), "tx": 5.13, "tw": 7.60},  # 좌측 표준
    {"img": (7.95, PHOTO_TOP, 4.78, _PH), "tx": 0.60, "tw": 7.00},  # 우측 넓게
    {"img": (0.60, PHOTO_TOP, 4.78, _PH), "tx": 5.72, "tw": 7.01},  # 좌측 넓게
    {"img": (8.20, PHOTO_TOP, 4.53, _PH), "tx": 0.60, "tw": 7.25},  # 우측 중간
]
# '하단 와이드'(2.60, 4.40, 8.13x2.22) 는 뺐다. 비율이 3.7:1 이라 1:1 로 주문한
# 그림을 넣으면 위아래가 잘리고, 본문이 위로 밀려 화면이 비어 보인다.
# below 처리 코드는 남겨 둔다 — 나중에 와이드 자리를 다시 넣을 수 있다.


# ── 타입별 렌더 ───────────────────────────────────────────────────────────
def render_cover(slide, s, img=None):
    """표지 — 회사 공식 양식 '00_표지' + 1차 정답지 3종의 세로 판.

    글 블록은 위에서 아래로 제목 → 부제 → 룰 → 발표자/소속, 그리고 맨 아래
    기관·기간 한 줄이다. 로고는 상단 띠가 아니라 **좌상단 글자열에 맞춰** 놓고,
    UBION 워드마크는 우하단으로 내린다(정답지 3종 공통).
    """
    _cover_frame(slide)
    x, w = COVER_TEXT_X, COVER_TEXT_W

    t = TYPE["cover_title"]
    title = s.get("title", "발표자료")
    # 제목 상자는 부제 윗선까지다 — 넘쳐도 부제를 밀지 않는다.
    box_h = COVER_SUB_Y - COVER_TITLE_Y - 0.13
    pt = _fit([_demark(title)], w, box_h, t["size"], t["line"], min_pt=24,
              label="cover title")
    tf = _text(slide, x, COVER_TITLE_Y, w, box_h, anchor="t")
    _T_split(_para(tf, True, line=t["line"]), title, "cover_title", size=pt)

    sub = s.get("subtitle") or s.get("chip") or ""
    if sub:
        st = TYPE["cover_sub"]
        spt = _fit([str(sub)], w, 0.95, 13.5, st["line"], min_pt=10.5,
                   label="cover sub")
        tf2 = _text(slide, x, COVER_SUB_Y, w, 0.95, anchor="t")
        _T(_para(tf2, True, line=st["line"]), sub, "cover_sub", size=spt)

    _rect(slide, x, COVER_RULE_Y, COVER_RULE_W, COVER_RULE_H, "amber")

    if s.get("presenter"):
        tf3 = _text(slide, x, COVER_PRESENTER_Y, w, 0.34, anchor="m")
        _T(_para(tf3, True), s["presenter"], "chip", size=15,
           color=on_white("navy"))
    if s.get("affiliation"):
        tf4 = _text(slide, x, COVER_AFFIL_Y, w, 0.28, anchor="m")
        _T(_para(tf4, True), s["affiliation"], "note", size=12)

    foot = s.get("footer") or s.get("band") or s.get("course") or ""
    if foot:
        tf5 = _text(slide, x, COVER_FOOT_Y, w, 0.30, anchor="m")
        _T(_para(tf5, True), foot, "caption", size=12)


def render_thanks(slide, s, img=None):
    """마감 장 — 표지 프레임에 한 줄만. 정답지 1ch S51 ('감사합니다.')."""
    _cover_frame(slide)
    x, w = COVER_TEXT_X, COVER_TEXT_W
    t = TYPE["thanks"]
    text = s.get("title") or LABELS.get("thanks", "감사합니다.")
    pt = _fit([str(text)], w, 0.95, t["size"], t["line"], min_pt=24,
              label="thanks")
    tf = _text(slide, x, 3.126, w, 0.95, anchor="t")
    _T(_para(tf, True, line=t["line"]), text, "thanks", size=pt)
    _rect(slide, x, 4.374, COVER_RULE_W, COVER_RULE_H, "amber")
    if s.get("chip"):
        tf2 = _text(slide, x, 4.62, w, 0.60, anchor="t")
        _T(_para(tf2, True, line=1.35), s["chip"], "cover_sub", size=13.5)


def render_section(slide, s, img=None):
    """구간 전환 — 대형 연번(연한 네이비) + 앰버 바 + 제목."""
    num = str(s.get("num") or "").strip()
    if num:
        tf0 = _text(slide, CONTENT_R - 4.2, 1.55, 4.2, 2.0, anchor="m")
        _T(_para(tf0, True, align="r"), num.zfill(2), "sec_num")
    _rrect(slide, MARGIN, 3.12, 0.11, 1.26, fill="amber", pill=True)
    t = TYPE["sec_title"]
    title = s.get("title", "")
    pt = _fit([_demark(title)], CONTENT_W - 0.6, 1.5, t["size"], t["line"],
              min_pt=22, label="section title")
    tf = _text(slide, MARGIN + 0.42, 2.98, CONTENT_W - 0.42, 1.55, anchor="m")
    _T_split(_para(tf, True, line=t["line"]), title, "sec_title", size=pt)
    if s.get("chip"):
        # 정답지 실측 18.75pt — 15pt 로는 96pt 연번에 눌려 안 읽혔다.
        tf2 = _text(slide, MARGIN + 0.42, 4.60, CONTENT_W - 0.42, 0.9, anchor="t")
        _T(_para(tf2, True, line=1.35), s["chip"], "cover_sub", size=18.75)


def render_bullets(slide, s, img=None):
    y = _head(slide, s)
    # 수치 레일이 있으면 본문 열을 그만큼 좁힌다(정답지 0083 S47 의 판).
    stats = [it for it in (s.get("stats") or []) if it]
    bw = (RAIL_X - MARGIN - RAIL_GAP) if stats else CONTENT_W
    by = _lead(slide, s, y, w=bw)
    top = max(by + 0.28, BODY_TOP)
    bottom = _body_bottom(s)
    if stats:
        add_stat_rail(slide, stats, top, bottom)
    add_bullets(slide, s.get("bullets", []), MARGIN + 0.04, top,
                bw - 0.08, bottom - top,
                label=f"bullets:{str(s.get('title'))[:18]}")
    add_interpretation(slide, _interp_of(s))


def render_photo(slide, s, img=None):
    """사진 슬라이드. 사진이 없어도 **자리는 남긴다.**

    예전에는 사진이 없으면 전폭 불릿으로 되돌아갔다. 그래서 초안(2단계) 덱에서는
    사진 자리가 보이지 않고, 사진을 넣는 3단계에서야 비로소 자리가 생기며 글이
    한쪽으로 밀렸다 — 배치가 두 번 바뀌는 셈이었다.

    지금은 빈 액자를 그려 자리를 확정한다. 그래서
      · 초안만 보고도 최종 배치를 알 수 있다
      · 사진 수집(3단계)을 건너뛰고 대본·영상으로 바로 가도 배치가 그대로다
      · 직접 그림을 그려 넣을 때 어느 크기·위치에 맞출지 눈으로 보인다
    """
    y = _head(slide, s)
    if str(s.get("layout") or "").lower() in ("2up", "two", "pair", "2"):
        by2 = _lead(slide, s, y)
        _photo_2up(slide, s, by2, img)
        add_interpretation(slide, _interp_of(s))
        return
    emph = bool(s.get("emphasis"))
    p = PHOTO_PRESETS[s.get("_photo_ord", 0) % len(PHOTO_PRESETS)]
    if img:
        add_photo(slide, img, *p["img"], credit=s.get("_credit_short"))
    else:
        add_photo_slot(slide, *p["img"])
    tx, tw = p["tx"], p["tw"]
    by = add_chip(slide, s.get("chip"), x=tx, y=y, w=tw, emphasis=emph)
    # 칩이 없으면 본문을 사진 윗선에 맞춘다(칩이 있을 때는 칩 아래).
    top = max(by + 0.26, BODY_TOP) if s.get("chip") else PHOTO_TOP
    bottom = (p["img"][1] - 0.22) if p.get("below") else BODY_BOTTOM
    add_bullets(slide, s.get("bullets", []), tx + 0.04, top, tw - 0.08,
                max(0.8, bottom - top),
                label=f"photo:{str(s.get('title'))[:18]}")


# 번호 밴드 스텝 — 정답지 0079 S16 실측(cm→in).
STEP_BAND_H = 0.665      # 밴드 높이 1.69cm
STEP_PITCH = 0.831       # 밴드 간격 2.11cm
STEP_DOT = 0.303         # 원형 번호 지름 0.77cm
STEP_DOT_DX, STEP_DOT_DY = 0.201, 0.177
STEP_LBL_DX, STEP_LBL_W = 0.657, 2.374
STEP_DSC_DX = 3.043
# 원형 번호는 옅은 것 → 진한 것으로 진행한다(정답지: CBD5E1·93D0B2·2A9760·1F9155).
_STEP_RAMP = ("navy_30", "navy_55", "amber", "navy")


def _process_band(slide, s, items, top, bottom):
    """번호 밴드 스텝 — 원형 번호 + 라벨 + **긴 설명 한 줄**.

    가로 박스 판(_process_boxes)은 설명이 한 문장을 넘으면 글꼴이 9pt 까지
    내려간다. 설명이 긴 단계는 밴드로 눕히는 쪽이 읽힌다(정답지 0079 S16).
    """
    n = len(items)
    pitch = min(STEP_PITCH, (bottom - top) / max(n, 1))
    band_h = max(0.42, min(STEP_BAND_H, pitch - 0.10))
    lt, dt = TYPE["step_label"], TYPE["step_desc"]
    labels = [str(it.get("label", "")) for it in items]
    lpt = min(_fit([v], STEP_LBL_W - 0.10, band_h - 0.10, lt["size"], lt["line"],
                   min_pt=10.5, label="step label") for v in labels)
    dsc_w = CONTENT_W - STEP_DSC_DX - 0.24
    descs = [str(it.get("desc") or "") for it in items]
    dpt = dt["size"]
    if any(descs):
        dpt = min(_fit([d], dsc_w, band_h - 0.10, dt["size"], dt["line"],
                       min_pt=9.5, label="step desc") for d in descs if d)
    for i, it in enumerate(items):
        by_ = top + i * pitch
        _rrect(slide, MARGIN, by_, CONTENT_W, band_h, fill="on_navy",
               radius=R_CARD)
        dot = band(_STEP_RAMP[min(i, len(_STEP_RAMP) - 1)], min_ratio=3.0)
        _oval(slide, MARGIN + STEP_DOT_DX, by_ + (band_h - STEP_DOT) / 2,
              STEP_DOT, dot)
        tfn = _text(slide, MARGIN + STEP_DOT_DX, by_, STEP_DOT, band_h,
                    anchor="m")
        _T(_para(tfn, True, align="c"), str(i + 1), "card_label", size=11.25,
           color=on_color(dot))
        tfl = _text(slide, MARGIN + STEP_LBL_DX, by_, STEP_LBL_W, band_h,
                    anchor="m")
        _T(_para(tfl, True, line=lt["line"]), it.get("label", ""), "step_label",
           size=lpt, color=on_white("navy"))
        if it.get("desc"):
            tfd = _text(slide, MARGIN + STEP_DSC_DX, by_, dsc_w, band_h,
                        anchor="m")
            _T(_para(tfd, True, line=dt["line"]), it["desc"], "step_desc",
               size=dpt)


def _band_wanted(s, items) -> bool:
    """밴드 판을 쓸지. `layout:"band"` 로 못박거나 **설명이 길면** 자동 전환."""
    lay = str(s.get("layout") or "").lower()
    if lay in ("band", "rows", "steps"):
        return True
    if lay in ("box", "boxes", "flow"):
        return False
    return any(len(str(it.get("desc") or "")) > 40 for it in items)


def render_process(slide, s, img=None):
    y = _head(slide, s)
    by = _lead(slide, s, y)
    items = s.get("items") or [{"label": b} for b in s.get("bullets", [])]
    items = items[:5] or [{"label": "항목"}]
    if _band_wanted(s, items):
        _process_band(slide, s, items, max(by + 0.24, BODY_TOP),
                      _body_bottom(s))
        add_interpretation(slide, _interp_of(s))
        return
    n = len(items)
    gap = 0.36
    node_h = 1.42
    node_w = (CONTENT_W - gap * (n - 1)) / n
    ny = BODY_TOP + 0.18
    has_desc = any(it.get("desc") for it in items)
    lt, ld = TYPE["node_label"], TYPE["node_desc"]

    labels = [str(it.get("label", "")) for it in items]
    lpt = min(_fit([l], node_w - 0.24, node_h - 0.2, lt["size"], lt["line"],
                   min_pt=10.5, label="process label") for l in labels)
    descs = [str(it.get("desc") or "") for it in items]
    dpt = ld["size"]
    if has_desc:
        dpt = min(_fit([d], node_w - 0.14, 1.45, ld["size"], ld["line"],
                       min_pt=9, label="process desc") for d in descs if d)

    for i, it in enumerate(items):
        x = MARGIN + i * (node_w + gap)
        # 단일 계조: 마지막 노드만 앰버로 도착점 표시
        last = (i == n - 1) and n > 1
        # ★ 채움은 band(), 글자색은 on_color() 로 **계산한다.** 색을 박아 두면
        #   배색을 바꿀 때 대비가 무너진다 — 시안 배색에서 amber(#1799B7) 위의
        #   navy(#0C93B8) 글씨가 1.05:1 로 아예 사라졌고(마지막 노드),
        #   navy_30(#CBD5E1) 위의 흰 글씨는 1.3:1 이었다(4번째 노드).
        #   계조가 밝은 노드는 어두운 글자, 어두운 노드는 흰 글자로 뒤집힌다.
        fill = band("amber") if last else band(_RAMP[min(i, len(_RAMP) - 1)])
        txt_color = on_color(fill)
        node = _rrect(slide, x, ny, node_w, node_h, fill=fill, radius=R_CARD)
        _shadow(node, blur=0.12, dist=0.03, alpha=10)
        tf = _text(slide, x + 0.12, ny, node_w - 0.24, node_h, anchor="m")
        _T(_para(tf, True, align="c", line=lt["line"]), it.get("label", ""),
           "node_label", size=lpt, color=txt_color)
        if it.get("desc"):
            tf2 = _text(slide, x + 0.07, ny + node_h + 0.14, node_w - 0.14, 1.45)
            _T(_para(tf2, True, align="c", line=ld["line"]), it["desc"],
               "node_desc", size=dpt)
        if i < n - 1:
            _arrow(slide, x + node_w + 0.05, ny + node_h / 2 - 0.13,
                   gap - 0.10, 0.26, "amber")
    add_interpretation(slide, _interp_of(s))


def render_cards(slide, s, img=None):
    y = _head(slide, s)
    _lead(slide, s, y)
    items = s.get("items") or [{"label": b} for b in s.get("bullets", [])]
    items = items[:4] or [{"label": "항목"}]
    numbered = bool(s.get("numbered", False))   # 비순차 분류는 숫자 없음
    n = len(items)
    gap = 0.34
    card_w = (CONTENT_W - gap * (n - 1)) / n
    cy = BODY_TOP + 0.05
    card_h = _body_bottom(s) - cy

    lt, dt = TYPE["card_label"], TYPE["card_desc"]
    labels = [str(it.get("label", "")) for it in items]
    lpt = min(_fit([l], card_w - 0.62, 0.62, lt["size"], lt["line"],
                   min_pt=11.5, label="card label") for l in labels)
    descs = [str(it.get("desc") or "") for it in items]
    label_bottom = 1.16 if numbered else 0.98
    desc_h = card_h - label_bottom - 0.26
    dpt = dt["size"]
    if any(descs):
        dpt = min(_fit([d], card_w - 0.62, desc_h, dt["size"], dt["line"],
                       min_pt=9.5, label="card desc") for d in descs if d)

    for i, it in enumerate(items):
        x = MARGIN + i * (card_w + gap)
        accent = _RAMP[min(i, len(_RAMP) - 1)]
        card = _rrect(slide, x, cy, card_w, card_h, fill="white", radius=R_CARD,
                      line="navy_12", line_w=1.0)
        _shadow(card, blur=0.13, dist=0.03, alpha=9)
        if numbered:
            # 연번 원에는 글자가 얹힌다 — 계조 뒤쪽(navy_55·navy_30)은 밝아서
            # 흰 숫자가 1.3:1 로 사라진다. 채움·글자색을 계산한다.
            dot = band(accent)
            _oval(slide, x + 0.30, cy + 0.30, 0.52, dot)
            tfb = _text(slide, x + 0.30, cy + 0.30, 0.52, 0.52, anchor="m")
            _T(_para(tfb, True, align="c"), str(i + 1), "card_label", size=15,
               color=on_color(dot))
        else:
            _rrect(slide, x, cy, card_w, 0.14, fill=accent, radius=0.07)
            _rrect(slide, x + 0.30, cy + 0.42, 0.30, 0.075, fill=accent, pill=True)
        tft = _text(slide, x + 0.30, cy + label_bottom - 0.62, card_w - 0.60, 0.62,
                    anchor="m")
        _T(_para(tft, True, line=lt["line"]), it.get("label", ""), "card_label",
           size=lpt, color="navy")
        if it.get("desc"):
            tfd = _text(slide, x + 0.30, cy + label_bottom, card_w - 0.60, desc_h)
            _T(_para(tfd, True, line=dt["line"]), it["desc"], "card_desc", size=dpt)
    add_interpretation(slide, _interp_of(s))


def render_compare(slide, s, img=None):
    y = _head(slide, s)
    _lead(slide, s, y)
    items = (s.get("items") or [])[:2]
    while len(items) < 2:
        items.append({"label": "", "lines": []})
    cy = BODY_TOP + 0.05
    col_h = _body_bottom(s) - cy
    col_w = (CONTENT_W - 0.42) / 2
    head_h = 0.60
    # 밴드에 글자가 얹히므로 대비가 모자라면 진한 짝으로 바꾼다(템플릿마다 다르다)
    fills = (band("navy"), band("amber"))
    # ★ 글자색을 고정하지 않고 **배경 밝기로 고른다.** 색을 박아 두면 템플릿에서
    #   밴드 색을 바꿨을 때 대비가 무너진다(초록 밴드에 남색 글씨가 흐릿했다).
    heads = tuple(on_color(f) for f in fills)
    td = TYPE["td"]

    for i, it in enumerate(items):
        x = MARGIN + i * (col_w + 0.42)
        col = _rrect(slide, x, cy, col_w, col_h, fill="chip", radius=R_CARD)
        _shadow(col, blur=0.12, dist=0.03, alpha=8)
        _rrect(slide, x, cy, col_w, head_h, fill=fills[i], radius=R_CARD)
        _rect(slide, x, cy + head_h - 0.10, col_w, 0.10, fills[i])
        tft = _text(slide, x + 0.2, cy, col_w - 0.4, head_h, anchor="m")
        _T(_para(tft, True, align="c"), it.get("label", ""), "card_label",
           size=15, color=heads[i])
        lines = it.get("lines") or ([it["desc"]] if it.get("desc") else [])
        lines = [str(l) for l in lines if str(l).strip()]
        if not lines:
            continue
        # 라벨만 가운데, 본문은 좌측 정렬(문장이 가운데면 읽히지 않는다)
        box_w, box_h = col_w - 0.66, col_h - head_h - 0.42
        sa = 6.0
        pt = _fit(lines, box_w, box_h, 14.0, td["line"], sa, min_pt=10,
                  label=f"compare:{it.get('label')}")
        lines = _trim_to_fit(lines, box_w, box_h, pt, td["line"], sa,
                             label=f"compare:{it.get('label')}")
        tf = _text(slide, x + 0.34, cy + head_h + 0.26, box_w, box_h, anchor="t")
        for j, ln in enumerate(lines):
            p = _para(tf, j == 0, space_after=sa, line=td["line"])
            _hang(p, 0.22)
            _run(p, "· ", pt, "amber", "sb")
            _T(p, ln, "td", size=pt)


def _cell_border(cell, color="white", w_pt=1.0):
    from pptx.oxml.ns import qn
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    # lnL, lnR, lnT, lnB 는 tcPr 의 맨 앞 순서
    for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):
        ln = tcPr.makeelement(qn(tag), {"w": str(int(w_pt * 12700)),
                                        "cap": "flat", "cmpd": "sng",
                                        "algn": "ctr"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        fill.append(fill.makeelement(qn("a:srgbClr"), {"val": _hexstr(color)}))
        ln.append(fill)
        tcPr.insert(0, ln)


# 표 — 정답지 3종 실측(cm→in). 머리행이 본문행보다 한 단계 높다.
TBL_HEAD_H = 0.382       # 0.97cm
TBL_ROW_H = 0.343        # 0.87cm
TBL_ROW_MIN = 0.28       # 이보다 낮으면 발표 화면에서 못 읽는다


def _table_widths(headers, rows_data, ncol, w):
    """열 폭을 **내용 비례**로 나눈다. 균등 분할은 문장 열이 접히고 수치 열이 놀았다."""
    lens = []
    for c in range(ncol):
        vals = [str(headers[c])] if c < len(headers) else []
        vals += [str(r[c]) for r in rows_data if c < len(r)]
        lens.append(max((_measure(v, 11.0) for v in vals), default=0.4))
    tot = sum(lens) or 1.0
    minw = min(0.58, w / ncol)
    extra = max(0.0, w - minw * ncol)
    return [minw + extra * (L / tot) for L in lens]


def _draw_table(slide, x, y, w, avail, headers, rows_data, label=""):
    """표 하나를 그린다. 반환 = 실제로 쓴 높이.

    머리행은 주색 채움 + 흰 글씨, 본문행은 흰색↔연회색(on_navy) 교차다. 짝수행에
    chip(주색의 아주 옅은 틴트) 을 쓰면 표 전체가 색을 띠어 무거워진다 — 정답지 3종도
    #F1F5F9 회색을 쓴다. 테두리는 흰 선
    1pt 로 셀 사이를 벌려 두었다 — 정답지가 셀을 0.02cm 떼어 그린 것과 같은 효과다.
    행 높이에는 하한(TBL_ROW_MIN)이 있어서, 행이 많으면 표가 화면을 넘는다.
    **줄이지 않고 넘긴 채로 두되** 몇 행이 밖으로 나갔는지 보고한다.
    """
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn

    ncol = len(headers) or max((len(r) for r in rows_data), default=1)
    nbody = len(rows_data)
    head_h = TBL_HEAD_H if headers else 0.0
    row_h = min(TBL_ROW_H, max(TBL_ROW_MIN, (avail - head_h) / max(nbody, 1)))
    used = head_h + row_h * nbody
    if used > avail + 0.01:
        fits = max(1, int((avail - head_h) / row_h))
        _note("overflow", f"table:{label}"[:40], f"표 {nbody - fits}행이 화면 밖")
    nrow = nbody + (1 if headers else 0)
    tbl_shape = slide.shapes.add_table(nrow, ncol, Inches(x), Inches(y),
                                       Inches(w), Inches(used))
    tbl = tbl_shape.table

    # PowerPoint 기본 파랑 밴딩 스타일 제거 — 셀 채움을 직접 지정한다
    try:
        tbl.first_row = bool(headers)
        tbl.horz_banding = False
        tbl.vert_banding = False
        tblPr = tbl._tbl.find(qn("a:tblPr"))
        if tblPr is not None:
            for el in tblPr.findall(qn("a:tableStyleId")):
                tblPr.remove(el)
    except Exception:  # noqa: BLE001
        pass
    try:
        for ci, cw in enumerate(_table_widths(headers, rows_data, ncol, w)):
            tbl.columns[ci].width = Emu(int(cw * 914400))
        for ri in range(nrow):
            h = head_h if (headers and ri == 0) else row_h
            tbl.rows[ri].height = Emu(int(h * 914400))
    except Exception:  # noqa: BLE001
        pass

    # 문장 열은 좌측, 짧은 값 열은 가운데
    col_len = []
    for c in range(ncol):
        vals = [str(r[c]) for r in rows_data if c < len(r)]
        col_len.append(max((len(v) for v in vals), default=0))
    aligns = [PP_ALIGN.LEFT if L > 12 else PP_ALIGN.CENTER for L in col_len]

    th, td = TYPE["th"], TYPE["td"]
    body_pt = td["size"]
    if rows_data:
        cell_w = w / ncol
        body_pt = min(
            _fit_each([str(r[c]) for r in rows_data if c < len(r)],
                      cell_w - 0.20, row_h - 0.08, td["size"], 1.16,
                      min_pt=8.5, label=f"table col{c}")
            for c in range(ncol))
    head_pt = min(th["size"], max(9.0, body_pt + 0.5))

    def _style(cell, text, *, header, align, bg):
        cell.text = str(text)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.02)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(bg)
        _cell_border(cell, "white", 1.0)
        for para in cell.text_frame.paragraphs:
            para.alignment = align
            para.line_spacing = 1.16
            for run in para.runs:
                run.font.size = Pt(head_pt if header else body_pt)
                run.font.bold = _FS.is_bold("sb" if header else "r")
                run.font.color.rgb = _rgb(on_color(bg) if header else "ink")
                pptx_font.apply_face(run, _FS.face("sb" if header else "r"))

    r0 = 0
    if headers:
        for c, htxt in enumerate(headers):
            # 머리행에 흰 글씨가 얹힌다 — 주색이 4.5:1 에 못 미치면 진한 짝으로.
            _style(tbl.cell(0, c), htxt, header=True, align=PP_ALIGN.CENTER,
                   bg=band("navy"))
        r0 = 1
    for ri, row in enumerate(rows_data):
        bg = "white" if ri % 2 == 0 else "on_navy"
        for c in range(ncol):
            _style(tbl.cell(ri + r0, c), row[c] if c < len(row) else "",
                   header=False, align=aligns[c], bg=bg)
    return used


def _table_blocks(s):
    """슬라이드에서 표 블록 목록. `tables` 배열이 있으면 그걸, 없으면 단일 표."""
    blocks = [b for b in (s.get("tables") or []) if b and b.get("rows")]
    if blocks:
        return blocks[:2]
    if not s.get("rows"):
        return []
    return [{"subhead": s.get("subhead"), "headers": s.get("headers"),
             "rows": s.get("rows"),
             "source": s.get("source") or s.get("unit")}]


def render_table(slide, s, img=None):
    """원자료 표. `tables` 로 두 개를 주면 좌우로 놓인다(정답지 0079 S12).

    각 표는 자기 **소제목**과 **단위·출처 캡션**을 갖는다 — 성과관리 보고에서
    수치는 출처 파일명이 화면에 남아야 검증이 된다.
    """
    y0 = _head(slide, s)
    by = _lead(slide, s, y0)
    blocks = _table_blocks(s)
    if not blocks:
        return render_bullets(slide, s)
    top = max(by + 0.30, BODY_TOP)
    bottom = _body_bottom(s)
    n = len(blocks)
    gap = 0.42
    w = (CONTENT_W - gap * (n - 1)) / n
    for i, b in enumerate(blocks):
        x = MARGIN + i * (w + gap)
        ty = add_subhead(slide, b.get("subhead"), top, x=x, w=w)
        ty = add_source(slide, b.get("source") or b.get("unit"), ty, x=x, w=w)
        if ty > top:
            ty += BLOCK_GAP
        _draw_table(slide, x, ty, w, max(0.5, bottom - ty),
                    b.get("headers") or [], b.get("rows") or [],
                    label=str(s.get("title") or "")[:20])
    add_interpretation(slide, _interp_of(s))


def render_quiz(slide, s, img=None):
    """확인 문항 1개 — 문제(네이비 칩) + 보기 알약 + 정답 뱃지 우하단.

    키커는 기본값을 두지 않는다 — 강의 용어가 사업 보고서 덱에 인쇄되던 문제.
    """
    s = dict(s)
    y = _head(slide, s)
    question = s.get("question") or s.get("chip") or ""
    by = add_chip(slide, question, y=y, emphasis=True) if question else y

    choices = s.get("choices") or s.get("items") or s.get("bullets") or []
    choices = [str(c).strip() for c in choices if str(c).strip()][:5]
    top = max(by + 0.34, BODY_TOP)
    answer = s.get("answer") or s.get("key")
    bottom = BODY_BOTTOM - (0.72 if answer else 0.0)

    if choices:
        marks = "①②③④⑤"
        n = len(choices)
        gap = 0.16
        row_h = min(0.82, max(0.52, (bottom - top - gap * (n - 1)) / n))
        pt = _fit_each(choices, CONTENT_W - 1.2, row_h - 0.12, 15.0, 1.2,
                       min_pt=11, label="quiz choice")
        for i, c in enumerate(choices):
            ry = top + i * (row_h + gap)
            _rrect(slide, MARGIN, ry, CONTENT_W, row_h, fill="chip", pill=True)
            # 번호 원에 글자가 얹힌다 — 채움·글자색을 계산한다(12pt 라 특히 얇다).
            dot = band("navy")
            _oval(slide, MARGIN + 0.20, ry + row_h / 2 - 0.15, 0.30, dot)
            tfm = _text(slide, MARGIN + 0.20, ry, 0.30, row_h, anchor="m")
            _T(_para(tfm, True, align="c"), marks[i], "card_label", size=12,
               color=on_color(dot))
            tf = _text(slide, MARGIN + 0.66, ry, CONTENT_W - 0.9, row_h, anchor="m")
            _T(_para(tf, True, line=1.2), c, "bullet", size=pt)
    else:
        add_bullets(slide, s.get("bullets", []), MARGIN + 0.04, top,
                    CONTENT_W - 0.08, bottom - top, label="quiz body")

    if answer:
        lbl = LABELS.get("answer", "정답")
        w = min(3.6, max(1.5, 0.30 + _measure(f"{lbl}  {answer}", 13.0)))
        ay = BODY_BOTTOM - 0.50
        # 알약에 글자가 얹힌다 — 채움·글자색을 계산한다. 시안 배색에서는
        # amber(#1799B7) 위의 amber_dk(#13819A) 가 1.05:1 로 안 보였다.
        pill_fill = band("amber")
        pill_txt = on_color(pill_fill)
        _rrect(slide, CONTENT_R - w, ay, w, 0.50, fill=pill_fill, pill=True)
        tf = _text(slide, CONTENT_R - w + 0.16, ay, w - 0.32, 0.50, anchor="m")
        p = _para(tf, True, align="c")
        _run(p, f"{lbl}  ", 12.0, pill_txt, "sb")
        _T(p, answer, "answer", color=pill_txt)


def render_objectives(slide, s, img=None):
    """대형 연번 + 한 줄 항목. 이미지 없음. (키커는 호출부가 정한다)"""
    s = dict(s)
    y = _head(slide, s)
    by = _lead(slide, s, y)
    goals = s.get("items") or s.get("bullets") or []
    goals = [(g.get("label") if isinstance(g, dict) else str(g)) for g in goals]
    goals = [str(g).strip() for g in goals if str(g).strip()][:4]
    if not goals:
        return render_bullets(slide, s)
    top = max(by + 0.34, BODY_TOP)
    n = len(goals)
    gap = 0.20
    row_h = min(1.30, (BODY_BOTTOM - top - gap * (n - 1)) / n)
    pt = _fit_each(goals, CONTENT_W - 1.72, row_h - 0.14, 17.0, 1.34,
                   min_pt=12, label="objective")
    for i, g in enumerate(goals):
        ry = top + i * (row_h + gap)
        tfn = _text(slide, MARGIN, ry, 1.30, row_h, anchor="m")
        _T(_para(tfn, True), str(i + 1).zfill(2), "big_num")
        _rect(slide, MARGIN + 1.16, ry + 0.10, 0.02, row_h - 0.20, "navy_12")
        tf = _text(slide, MARGIN + 1.42, ry, CONTENT_W - 1.42, row_h, anchor="m")
        _T(_para(tf, True, line=1.34), g, "bullet", size=pt)


# 목차 행 — 정답지 0079 S2 실측(cm→in).
TOC_NUM_W = 1.386        # 대형 연번 열 3.52cm
TOC_DIV_DX = 1.16        # 세로 구분선 4.47cm
TOC_TXT_DX = 1.42        # 글자 시작 5.13cm
TOC_PITCH = 0.917        # 행 간격 2.33cm


def render_agenda(slide, s, img=None):
    """목차 — 대형 연번 + 세로 구분선 + (구간명) + 제목. 정답지 0079 S2 의 판.

    예전 판(원형 번호 2열)은 항목이 7개 이상일 때만 쓴다 — 5개 안팎에서는
    줄이 짧아 좌우가 헐렁하고, 구간명(Part I)을 놓을 자리가 없었다.
    """
    s = dict(s)
    s.setdefault("kicker", LABELS.get("agenda", "목차"))
    y = _head(slide, s)
    by = _lead(slide, s, y)
    rows = []
    for it in (s.get("items") or s.get("bullets") or []):
        if isinstance(it, dict):
            rows.append((str(it.get("part") or "").strip(),
                         str(it.get("label") or it.get("title") or "").strip()))
        else:
            rows.append(("", str(it).strip()))
    rows = [r for r in rows if r[1]][:10]
    if not rows:
        return render_bullets(slide, s)
    top = max(by + 0.34, BODY_TOP)
    bottom = _body_bottom(s)
    if len(rows) > 6:
        _agenda_grid(slide, s, [r[1] for r in rows], top, bottom)
        add_interpretation(slide, _interp_of(s))
        return

    n = len(rows)
    pitch = min(TOC_PITCH, (bottom - top) / n)
    active = s.get("active")
    tp, tt = TYPE["toc_part"], TYPE["toc_title"]
    parts = [r[0] for r in rows if r[0]]
    ppt = tp["size"]
    pw = 0.0
    if parts:
        ppt = min(_fit([v], 2.40, pitch - 0.12, tp["size"], tp["line"],
                       min_pt=11.5, label="toc part") for v in parts)
        pw = min(2.40, max(_measure(v, ppt) for v in parts) + 0.20)
    tx = MARGIN + TOC_TXT_DX + pw
    tpt = _fit_each([r[1] for r in rows], CONTENT_R - tx, pitch - 0.12,
                    tt["size"], tt["line"], min_pt=11.5, label="toc title")
    for i, (part, title) in enumerate(rows):
        ry = top + i * pitch
        on = (active is not None and str(active).isdigit()
              and int(active) == i + 1)
        tfn = _text(slide, MARGIN, ry, TOC_NUM_W, pitch, anchor="m")
        _T(_para(tfn, True), str(i + 1).zfill(2), "toc_num",
           color=(on_white("navy") if on else None))
        _rect(slide, MARGIN + TOC_DIV_DX, ry + 0.11, 0.02, pitch - 0.22,
              "navy_12")
        if part:
            tfp = _text(slide, MARGIN + TOC_TXT_DX, ry, pw, pitch, anchor="m")
            _T(_para(tfp, True), part, "toc_part", size=ppt,
               color=on_white("navy"))
        tf = _text(slide, tx, ry, CONTENT_R - tx, pitch, anchor="m")
        _T(_para(tf, True, line=tt["line"]), title, "toc_title", size=tpt,
           weight=("sb" if on else None))
    add_interpretation(slide, _interp_of(s))


def _agenda_grid(slide, s, items, top, bottom):
    """목차 — 2열 번호 목록. active 로 현재 구간 강조."""
    s = dict(s)
    s.setdefault("kicker", "목차")
    y = _head(slide, s)
    by = _lead(slide, s, y)
    items = s.get("items") or s.get("bullets") or []
    items = [(i.get("label") if isinstance(i, dict) else str(i)) for i in items]
    items = [str(i).strip() for i in items if str(i).strip()][:10]
    if not items:
        return render_bullets(slide, s)
    active = s.get("active")
    top = max(by + 0.34, BODY_TOP)
    ncol = 2 if len(items) > 4 else 1
    per = math.ceil(len(items) / ncol)
    col_w = (CONTENT_W - 0.5) / ncol
    row_h = min(0.80, (BODY_BOTTOM - top) / per)
    pt = _fit_each(items, col_w - 0.85, row_h - 0.12, 15.5, 1.25, min_pt=11,
                   label="agenda")
    for i, it in enumerate(items):
        c, r = divmod(i, per)
        x = MARGIN + c * (col_w + 0.5)
        ry = top + r * row_h
        on = (active is not None and int(active) == i + 1)
        if on:
            _rrect(slide, x - 0.10, ry + 0.04, col_w + 0.10, row_h - 0.08,
                   fill="chip", radius=R_CARD)
        # 비활성 원은 navy_30(연한 회색)이라 흰 숫자가 1.3:1 로 안 보였다.
        dot = band("navy") if on else "navy_30"
        _oval(slide, x, ry + row_h / 2 - 0.15, 0.30, dot)
        tfn = _text(slide, x, ry, 0.30, row_h, anchor="m")
        _T(_para(tfn, True, align="c"), str(i + 1), "card_label", size=11.5,
           color=on_color(dot))
        tf = _text(slide, x + 0.46, ry, col_w - 0.5, row_h, anchor="m")
        _T(_para(tf, True, line=1.25), it, "bullet", size=pt,
           color=("navy" if on else "ink"), weight=("sb" if on else "r"))


def _closing_flow(slide, s, by):
    """마무리 — ▸ 요약 불릿 + 하단 '확인 필요' 패널. 정답지 0079 S20 의 판.

    예전에는 불릿만 주면 그것을 셋으로 쪼개 '요약/다음 단계/후속 조치' 라벨을
    붙여 3열 카드로 만들었다. 그런데 결론 문장 넷을 억지로 셋으로 나누면 라벨과
    내용이 안 맞는다. 지금은 결론을 한 흐름으로 두고, 미확인 항목만 아래 패널로
    내린다(그 자리가 정답지의 '[Missing Attachments]' 다).
    """
    bl = [str(b) for b in (s.get("bullets") or []) if str(b).strip()]
    panel = s.get("panel") or {}
    plines = [str(l) for l in (panel.get("lines") or []) if str(l).strip()]
    top = max(by + 0.28, BODY_TOP)
    bottom = _body_bottom(s)
    ph = 0.0
    if plines:
        ph = min(0.62 + 0.40 * len(plines), max(0.90, (bottom - top) * 0.55))
    if bl:
        add_bullets(slide, bl, MARGIN + 0.04, top, CONTENT_W - 0.08,
                    max(0.6, bottom - (ph + 0.24 if ph else 0.0) - top),
                    role="closing_line", color=on_white("amber"),
                    label="closing")
    if plines:
        py = bottom - ph
        _rrect(slide, MARGIN, py, CONTENT_W, ph, fill="on_navy", radius=R_CARD)
        ly = py + 0.16
        if panel.get("label"):
            tfl = _text(slide, MARGIN + 0.33, ly, CONTENT_W - 0.66, 0.30,
                        anchor="m")
            _T(_para(tfl, True), panel["label"], "kpi_label")
            ly += 0.34
        pt = _fit_each(plines, CONTENT_W - 0.90, 0.34, 12.0, 1.30, min_pt=9.5,
                       label="closing panel")
        tf = _text(slide, MARGIN + 0.33, ly, CONTENT_W - 0.66,
                   max(0.30, py + ph - 0.10 - ly), anchor="t")
        for j, ln in enumerate(plines):
            pp = _para(tf, j == 0, space_after=3.0, line=1.30)
            _hang(pp, 0.20)
            _run(pp, "· ", pt, "navy_55", "sb")
            _T(pp, ln, "note", size=pt)
    add_interpretation(slide, _interp_of(s))


def render_closing(slide, s, img=None):
    """마무리 — `items` 3블록 카드, 없으면 불릿 + 하단 패널 흐름판."""
    s = dict(s)
    y = _head(slide, s)
    by = _lead(slide, s, y)
    items = s.get("items") or []
    if not items:
        return _closing_flow(slide, s, by)
    items = items[:3]
    n = len(items)
    gap = 0.34
    col_w = (CONTENT_W - gap * (n - 1)) / n
    cy = BODY_TOP + 0.05
    col_h = _body_bottom(s) - cy
    td = TYPE["td"]
    for i, it in enumerate(items):
        x = MARGIN + i * (col_w + gap)
        box = _rrect(slide, x, cy, col_w, col_h, fill="white", radius=R_CARD,
                     line="navy_12", line_w=1.0)
        _shadow(box, blur=0.13, dist=0.03, alpha=9)
        _rrect(slide, x, cy, col_w, 0.14, fill=_RAMP[min(i, 3)], radius=0.07)
        tfl = _text(slide, x + 0.28, cy + 0.34, col_w - 0.56, 0.5, anchor="m")
        _T(_para(tfl, True), it.get("label", ""), "card_label")
        lines = it.get("lines") or ([it["desc"]] if it.get("desc") else [])
        lines = [str(l) for l in lines if str(l).strip()]
        if not lines:
            continue
        box_w, box_h = col_w - 0.56, col_h - 1.10
        sa = 6.0
        pt = _fit(lines, box_w, box_h, 13.0, td["line"], sa, min_pt=9.5,
                  label=f"closing:{it.get('label')}")
        lines = _trim_to_fit(lines, box_w, box_h, pt, td["line"], sa,
                             label=f"closing:{it.get('label')}")
        tf = _text(slide, x + 0.28, cy + 0.94, box_w, box_h, anchor="t")
        for j, ln in enumerate(lines):
            p = _para(tf, j == 0, space_after=sa, line=td["line"])
            _hang(p, 0.22)
            _run(p, "· ", pt, "amber", "sb")
            _T(p, ln, "td", size=pt)


# 목표 대비 실적 카드 — 정답지 0079 S7 실측(cm→in).
KPI_BAR_H = 0.134        # 상단 액센트 바 0.34cm
KPI_LBL_DY = 0.386       # 카드 상단→라벨 0.98cm
KPI_NUM_DY, KPI_NUM_H = 0.803, 1.008     # 2.04cm / 2.56cm
KPI_CAP_DY = 1.811       # 4.60cm
_KPI_BAR = ("amber", "navy", "navy_75")


def render_stat(slide, s, img=None):
    """목표 대비 실적 — 큰 수치 카드 2~3개. 정답지 0079 S7·S14·S18 의 판.

    상단 얇은 액센트 바 → 카드 이름(caption) → 48pt 수치 → 한 줄 근거.
    `kind:"target"` 항목은 **목표**이므로 수치를 본문색으로 눕히고 실적만 주색으로
    세운다 — 목표와 결과가 같은 색이면 어느 쪽이 성과인지 한눈에 갈리지 않는다.

    items = [{"label": 수치, "caption": 카드 이름, "desc": 근거 한 줄,
              "kind": "target"|"result"}]
    """
    y = _head(slide, s)
    _lead(slide, s, y)
    items = [it for it in (s.get("items") or []) if it]
    if not items:
        return render_bullets(slide, s)
    items = items[:3]
    n = len(items)
    gap = 0.34
    col_w = (CONTENT_W - gap * (n - 1)) / n
    cy = BODY_TOP + 0.15
    col_h = max(1.60, _body_bottom(s) - cy)
    nt, ct = TYPE["kpi_num"], TYPE["kpi_cap"]
    nums = [str(it.get("label", "")) for it in items]
    npt = min(_fit([v], col_w - 0.30, KPI_NUM_H, nt["size"], 1.06, min_pt=24,
                   label="stat num") for v in nums)
    caps = [str(it.get("desc") or "") for it in items]
    cap_h = max(0.28, col_h - KPI_CAP_DY - 0.18)
    cpt = ct["size"]
    if any(caps):
        cpt = min(_fit([c], col_w - 0.70, cap_h, ct["size"], ct["line"],
                       min_pt=9.5, label="stat desc") for c in caps if c)
    ri = 0
    for i, it in enumerate(items):
        x = MARGIN + i * (col_w + gap)
        target = str(it.get("kind") or "").lower().startswith("t")
        if target:
            bar = "navy_55"
        else:
            bar = _KPI_BAR[ri % len(_KPI_BAR)]
            ri += 1
        _rrect(slide, x, cy, col_w, col_h, fill="on_navy", radius=R_CARD)
        _rrect(slide, x, cy, col_w, KPI_BAR_H, fill=bar, radius=0.07)
        lbl = it.get("caption") or it.get("name") or ""
        if lbl:
            tfl = _text(slide, x + 0.27, cy + KPI_LBL_DY, col_w - 0.54, 0.30,
                        anchor="m")
            _T(_para(tfl, True), lbl, "kpi_label")
        tfn = _text(slide, x + 0.15, cy + KPI_NUM_DY, col_w - 0.30, KPI_NUM_H,
                    anchor="m")
        _T(_para(tfn, True, align="c"), nums[i], "kpi_num", size=npt,
           color=("ink" if target else on_white("navy")))
        if it.get("desc"):
            tfd = _text(slide, x + 0.35, cy + KPI_CAP_DY, col_w - 0.70, cap_h,
                        anchor="t")
            _T(_para(tfd, True, align="c", line=ct["line"]), it["desc"],
               "kpi_cap", size=cpt)
    add_interpretation(slide, _interp_of(s))


# ── 차트 ───────────────────────────────────────────────────────────────────
# 이 엔진은 강의 덱에서 왔고 강의 덱엔 차트가 거의 없었다. 그런데 AIM 이 만드는
# 성과관리(M&E) 보고 덱은 정반대다 — 1차 정답지 2종(0083·0079)을 세어 보면
# 51장 중 차트 33개(막대 26 · 꺾은선 7)·표 33개로, 차트가 사실상 본문이다.
# 그래서 여기에 네이티브 차트를 더한다. 그림이 아니라 진짜 차트라서 발표자가
# PowerPoint 에서 값을 고칠 수 있다.
#
# 정답지의 지배적 형태를 그대로 따른다: **한 장에 차트 2개를 좌우로**, 각 차트
# 위에 설명 제목, 아래에 데이터 표(c:dTable), 범례는 하단.
# 계열 순서대로 진한 것 → 옅은 것. 단일 색조 배색에서도 계열이 갈린다.
_CHART_RAMP = ("navy", "navy_55", "navy_75", "navy_30")

_CHART_KINDS = {
    "column": "COLUMN_CLUSTERED", "bar": "BAR_CLUSTERED",
    "line": "LINE_MARKERS", "line_plain": "LINE",
    "pie": "PIE", "doughnut": "DOUGHNUT",
    "column_stacked": "COLUMN_STACKED", "bar_stacked": "BAR_STACKED",
}


def _to_num(v):
    """차트 값 → float. 비숫자('[확인 필요]' 등)·빈값은 None.

    ★ 이 가드를 빼면 안 된다. 비숫자 값을 차트에 그대로 넣으면 PowerPoint 가
      '복구할 수 없는 내용' 이라며 파일을 못 연다. AIM 은 근거가 없을 때
      '[확인 필요]' 를 남기는 도구라서 비숫자가 실제로 자주 들어온다.
    """
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).strip().replace(",", "").replace("%", "").replace(" ", "")
        return float(s) if s else None
    except (TypeError, ValueError):
        return None


def _num(v):
    """차트 값은 **숫자만** 받는다 — 못 읽으면 빈 칸(None).

    모델이 모르는 값을 '[확인 필요]' 로 적어 보내는 일이 있고, 그게 계열 값으로
    들어가면 PowerPoint 가 파일을 거부한다. 한 칸 때문에 덱 전체를 잃지 않도록
    조용히 빈 칸으로 둔다(화면에서 끊긴 자리로 보인다). '12.3%' · '1,234' 는 읽는다.
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "").rstrip("%").strip()
    try:
        return float(t)
    except ValueError:
        return None


def _norm_series(raw):
    """[{name, values}] / {name: values} / [[name, values]] 를 모두 받는다."""
    out = []
    if isinstance(raw, dict):
        for k, v in raw.items():
            out.append((str(k), [_num(x) for x in (v or [])]))
    elif isinstance(raw, list):
        for s in raw:
            if isinstance(s, dict):
                out.append((str(s.get("name", "")),
                            [_num(x) for x in (s.get("values") or [])]))
            elif isinstance(s, (list, tuple)) and len(s) == 2:
                out.append((str(s[0]), [_num(x) for x in (s[1] or [])]))
    return out


def _chart_data_table(chart) -> None:
    """차트 아래 데이터 표(c:dTable). python-pptx 에 API 가 없어 XML 로 넣는다.

    스키마 순서상 dTable 은 plotArea 안에서 축(valAx/catAx) **뒤**, spPr 앞이다.
    순서를 어기면 PowerPoint 가 파일을 거부한다.
    """
    from pptx.oxml.ns import qn
    plot_area = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    if plot_area is None or plot_area.find(qn("c:dTable")) is not None:
        return
    dt = plot_area.makeelement(qn("c:dTable"), {})
    for tag, val in (("c:showHorzBorder", "1"), ("c:showVertBorder", "1"),
                     ("c:showOutline", "1"), ("c:showKeys", "1")):
        el = dt.makeelement(qn(tag), {})
        el.set("val", val)
        dt.append(el)
    spPr = plot_area.find(qn("c:spPr"))
    if spPr is not None:
        spPr.addprevious(dt)
    else:
        plot_area.append(dt)


def _style_chart(chart, series_n: int, number_format: str, *, labels=True,
                 legend=True) -> None:
    from pptx.util import Pt
    from pptx.enum.chart import XL_LEGEND_POSITION, XL_TICK_MARK

    try:
        chart.font.size = Pt(10)
        chart.font.color.rgb = _rgb("ink")
        chart.font.name = _FS.face("r")
    except Exception:  # noqa: BLE001
        pass

    chart.has_title = False
    chart.has_legend = bool(legend and series_n > 1)
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        try:
            chart.legend.font.size = Pt(10)
        except Exception:  # noqa: BLE001
            pass

    # 계열색 — **명도 대비**로 가른다. 색상(hue)으로 가르면 안 된다:
    # 이 배색은 단일 색조라서 navy(#1F9155)와 amber(#2A9760)가 사실상 같은 초록이고,
    # 그 둘로 2계열 막대를 칠했더니 Math/BS 가 구분되지 않았다(실측).
    # 진한 것 → 옅은 것 순으로 두면 흑백 인쇄에서도 읽힌다.
    ramp = _CHART_RAMP
    for i, s in enumerate(chart.series):
        try:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = _rgb(ramp[i % len(ramp)])
            s.format.line.color.rgb = _rgb(ramp[i % len(ramp)])
        except Exception:  # noqa: BLE001
            pass

    if labels:
        try:
            plot = chart.plots[0]
            plot.has_data_labels = True
            plot.data_labels.number_format = number_format
            plot.data_labels.number_format_is_linked = False
            plot.data_labels.font.size = Pt(9)
            plot.data_labels.font.color.rgb = _rgb("ink")
        except Exception:  # noqa: BLE001
            pass
        try:
            chart.plots[0].gap_width = 60
        except Exception:  # noqa: BLE001
            pass

    for ax_name in ("value_axis", "category_axis"):
        try:
            ax = getattr(chart, ax_name)
            ax.has_major_gridlines = (ax_name == "value_axis")
            if ax.has_major_gridlines:
                ax.major_gridlines.format.line.color.rgb = _rgb("navy_12")
            ax.major_tick_mark = XL_TICK_MARK.NONE
            ax.minor_tick_mark = XL_TICK_MARK.NONE
            ax.format.line.color.rgb = _rgb("navy_30")
            ax.tick_labels.font.size = Pt(9)
            ax.tick_labels.font.color.rgb = _rgb("muted")
            if ax_name == "value_axis":
                ax.tick_labels.number_format = number_format
                ax.tick_labels.number_format_is_linked = False
        except Exception:  # noqa: BLE001
            pass


def _one_chart(slide, spec, x, y, w, h) -> None:
    """차트 한 개 — 위에 설명 제목, 아래에 차트."""
    from pptx.util import Inches
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    # 소제목·출처는 표와 **같은 판**을 쓴다 — 한 장에 표와 차트가 같이 놓이는
    # 일이 잦고, 그때 캡션 정렬·크기가 다르면 두 블록이 서로 다른 자료처럼 보인다.
    cap = str(spec.get("caption") or spec.get("header") or "").strip()
    src_line = " · ".join(t for t in (str(spec.get("unit") or "").strip(),
                                     str(spec.get("source") or "").strip()) if t)
    cy, ch = y, h
    ny = add_subhead(slide, cap, cy, x=x, w=w)
    ny = add_source(slide, src_line, ny, x=x, w=w)
    if ny > cy:
        ny += BLOCK_GAP * 0.6
    ch -= (ny - cy)
    cy = ny

    cats = [str(c) for c in (spec.get("categories") or [])] or [""]
    series = _norm_series(spec.get("series"))
    num = [(n, [_to_num(v) for v in vals]) for n, vals in series]
    has_num = any(any(v is not None for v in vals) for _, vals in num)

    if not has_num:
        # 값이 없으면 깨진 차트 대신 '무엇을 채워야 하는지' 를 적는다.
        _rrect(slide, x, cy, w, ch, fill="chip", radius=R_CARD)
        tf = _text(slide, x + 0.24, cy + 0.18, w - 0.48, ch - 0.36, anchor="t")
        first = True
        for c in cats[:8]:
            p = _para(tf, first, line=1.4)
            first = False
            _run(p, f"· {c}  —  [확인 필요]", 11.5, "ink")
        if first:
            _run(_para(tf, True), "데이터 [확인 필요]", 11.5, "ink")
        return

    cd = CategoryChartData()
    cd.categories = cats
    for name, vals in num:
        # 부분 결측은 0 이 아니라 None 으로 둔다 — 0 은 '값이 0' 이라는 거짓말이다.
        cd.add_series(name or "값", vals)

    kind = _CHART_KINDS.get(str(spec.get("kind") or "column").lower(), "COLUMN_CLUSTERED")
    gf = slide.shapes.add_chart(getattr(XL_CHART_TYPE, kind), Inches(x), Inches(cy),
                                Inches(w), Inches(max(0.8, ch)), cd)
    chart = gf.chart
    nfmt = str(spec.get("number_format") or "0.00")
    _style_chart(chart, len(num), nfmt,
                 labels=bool(spec.get("data_labels", True)),
                 legend=bool(spec.get("legend", True)))
    if spec.get("data_table"):
        try:
            _chart_data_table(chart)
        except Exception as e:  # noqa: BLE001
            _log(f"[deck] 데이터 표 삽입 실패(무시): {e}")


def render_chart(slide, s, img=None):
    """네이티브 차트 1~2개. 두 개면 좌우로 나눈다(정답지의 지배적 형태)."""
    y = _head(slide, s)
    # 칩이 없으면 요약줄 자리(1.88)부터 쓴다. 다른 타입은 본문 기준선(2.95)에
    # 맞춰 정렬하지만 차트는 높이가 곧 가독성이라, 빈 띠를 남기지 않는다.
    top = _lead(slide, s, y)
    bottom = BODY_BOTTOM - (0.74 if s.get("key") else 0.0)

    charts = s.get("charts")
    if not isinstance(charts, list) or not charts:
        # 슬라이드에 바로 categories/series 를 준 축약형도 받아 준다.
        charts = [s] if s.get("series") or s.get("categories") else []
    if not charts:
        return render_bullets(slide, s)
    charts = charts[:2]

    if len(charts) == 1:
        _one_chart(slide, charts[0], MARGIN, top, CONTENT_W, bottom - top)
    else:
        gap = 0.34
        cw = (CONTENT_W - gap) / 2
        for i, c in enumerate(charts):
            _one_chart(slide, c, MARGIN + i * (cw + gap), top, cw, bottom - top)

    add_interpretation(slide, _interp_of(s))


_RENDER = {
    "cover": render_cover,
    "thanks": render_thanks,
    "chart": render_chart,
    "section": render_section,
    "photo": render_photo,
    "process": render_process,
    "cards": render_cards,
    "compare": render_compare,
    "table": render_table,
    "bullets": render_bullets,
    "quiz": render_quiz,
    "objectives": render_objectives,
    "agenda": render_agenda,
    "closing": render_closing,
    "stat": render_stat,
}

# 러닝 푸터·페이지 번호를 넣지 않는 타입
_NO_RUNNING = ("cover", "thanks", "section")


def build_deck(plan: List[Dict], template_path: Optional[str] = None,
               images: Optional[Dict[int, bytes]] = None,
               deck_title: str = "강의 슬라이드",
               logo_path: Optional[str] = None, *,
               logo2_path: Optional[str] = None,
               footer: str = "", assets_dir=None,
               embed_font: bool = True,
               report: Optional[List[Dict]] = None) -> Optional[bytes]:
    """슬라이드플랜 → 디자인된 .pptx 바이트. python-pptx 미설치 시 None.

    embed_font=False 면 Pretendard 를 '이름으로만' 지정한다(파일 ~3MB 작아지지만
    폰트가 설치되지 않은 PC 에서는 대체 폰트로 열린다).

    배색·글꼴은 **template_path(견본 PPTX)에서 뽑는다** — 고르는 화면은 없다.
    견본이 없으면 palette.BASELINE 으로 떨어진다.
    """
    global _FS
    try:
        from pptx import Presentation
        from pptx.util import Inches
    except Exception:  # noqa: BLE001
        return None

    # ★ 색·글꼴을 먼저 갈아끼운다. 도형을 그리기 시작한 뒤에 바꾸면 앞뒤 장의
    #   색이 섞인다.
    from . import palette as _pal
    _pal.apply(template_path)

    adir = assets_dir or ASSETS_DIR
    # 견본이 지정한 글꼴을 우선 찾는다 — assets/fonts 에 그 정적 TTF 가 있으면
    # 파일째 덱에 심어(임베드) 어느 PC 에서 열어도 같은 모양이 나온다.
    _FS = pptx_font.font_set(adir, prefer=(TEMPLATE_FONT or "").strip() or None)
    # 템플릿이 **Pretendard 아닌** 글꼴을 지정했을 때만 갈아끼운다.
    #   · Pretendard 는 font_set() 이 이미 다룬다 — assets 에 파일이 있으면 임베드하고,
    #     없으면 맑은 고딕으로 폴백한다. 여기서 덮어쓰면 그 폴백이 깨져서
    #     설치도 안 된 Pretendard 를 이름으로만 지정하게 된다(실측: 회귀했다).
    #   · 임베드는 Pretendard 만 가능하므로 다른 글꼴은 이름만 넣는다.
    _tpl_font = (TEMPLATE_FONT or "").strip()
    if _tpl_font and _FS.regular.lower() != _tpl_font.lower():
        # 견본 글꼴의 TTF 가 assets/fonts 에 없다 → 이름만 지정한다.
        # 받는 PC 에 그 글꼴이 없으면 대체돼 줄바꿈까지 달라진다.
        _FS = pptx_font.FontSet(regular=_tpl_font, semibold=_tpl_font,
                                bold=_tpl_font, embedded=False)
        _log(f"[font] 견본 글꼴: {_tpl_font} — assets/fonts 에 없어 이름만 지정")

    # 넘침 보고를 담을 곳. 호출부가 리스트를 주면 거기에 쌓인다.
    _TL.report = report if report is not None else []
    _TL.slide = None

    images = images or {}
    use_tpl = bool(template_path and os.path.isfile(template_path))
    if use_tpl:
        try:
            prs = Presentation(template_path)
            _remove_all_slides(prs)
        except Exception:  # noqa: BLE001
            prs = Presentation()
            use_tpl = False
    else:
        prs = Presentation()
    if not use_tpl:
        prs.slide_width = Inches(SLIDE_W)
        prs.slide_height = Inches(SLIDE_H)

    if _FS.embedded and embed_font:
        got = pptx_font.embed_fonts(prs, adir, only=_FS.regular)
        if not got:
            _FS = pptx_font.SYSTEM_SET     # 임베드 실패 → 안전하게 시스템 폰트로
            _log("[font] 임베드 불가 — 맑은 고딕으로 폴백")
        else:
            _log(f"[font] 임베드: {', '.join(got)}")
    elif _FS.embedded:
        _log("[font] 임베드 끄기 — Pretendard 이름만 지정")

    has_logo = bool(logo_path and os.path.isfile(logo_path))
    has_logo2 = bool(logo2_path and os.path.isfile(logo2_path))
    layout = _blank_layout(prs)
    photo_ord = 0
    sec_num = 0
    fallbacks: List[int] = []

    for i, s in enumerate(plan):
        slide = prs.slides.add_slide(layout)
        _strip_placeholders(slide)
        typ = (s.get("type") or "bullets").lower()
        # 넘침 보고에 붙일 '지금 그리는 장' — 사람이 그 번호로 채팅에 적는다.
        _TL.slide = {"no": i + 1, "title": str(s.get("title") or ""), "type": typ}
        s["_idx"] = i
        if typ == "photo":
            s["_photo_ord"] = photo_ord
            photo_ord += 1
        if typ == "section":
            sec_num += 1
            s.setdefault("num", sec_num)
        if typ == "cover":
            s.setdefault("footer", footer)
        s["_has_logo"] = has_logo
        # ★ 로고를 **먼저** 놓는다. 좌상단 로고의 실제 폭을 알아야 키커를 밀 수 있고,
        #   폭은 그림을 넣어 봐야 안다. 표지는 자체 레이아웃이라 넣지 않는다.
        s["_kicker_dx"] = 0.0
        if typ not in ("cover", "thanks") and (has_logo or has_logo2):
            lw = add_logos(slide, logo_path if has_logo else None,
                           logo2_path if has_logo2 else None)
            s["_kicker_dx"] = (lw + 0.22) if lw else 0.0
        fn = _RENDER.get(typ, render_bullets)
        try:
            fn(slide, s, images.get(i))
            # 표지는 **그린 뒤에** 로고를 얹는다 — 먼저 놓으면 우측 패널이
            # 그 위를 덮어 로고가 사라진다(그리는 순서 = 쌓이는 순서).
            # 자리는 다른 장과 같다(공식 양식은 표지에도 같은 헤더를 쓴다).
            if typ in ("cover", "thanks") and (has_logo or has_logo2):
                add_logos(slide, logo_path if has_logo else None,
                          logo2_path if has_logo2 else None,
                          left_x=COVER_TEXT_X, left_y=COVER_LOGO_Y,
                          right_x=LOGO_R + (COVER_LOGO_RX - LOGO_RX),
                          right_y=COVER_LOGO_RY)
        except Exception as e:  # noqa: BLE001
            _log(f"[deck] slide {i + 1} ({typ}) 렌더 오류: {e}")
            fallbacks.append(i + 1)
            try:
                render_bullets(slide, s)
            except Exception:  # noqa: BLE001
                pass
        if typ not in _NO_RUNNING:
            add_running(slide, footer, i + 1)

    if fallbacks:
        _log(f"[deck] 폴백 렌더된 슬라이드: {fallbacks}")

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ── 아트디렉터(LLM) 패스 — 개요 md → 슬라이드플랜 JSON ─────────────────────
_ART_SYS = (
    "너는 강의 슬라이드 아트디렉터다. 주어진 슬라이드 개요를 슬라이드플랜 JSON 배열로만 변환한다. "
    "설명·머리말·코드펜스 없이 JSON 배열만 출력한다."
)

_ART_RULES = """개요의 '### 슬라이드 N' 블록 하나당 JSON 객체 하나를 순서대로 만든다(개수 유지).
각 객체 스키마:
{
  "type": "section|photo|process|cards|compare|table|quiz|objectives|agenda|closing|stat|bullets",
  "title": "간결·정확한 제목",
  "kicker": "상단 초소형 라벨(선택, 6자 이내. 예 '개념'·'사례'·'이론')",
  "chip": "요약 한 문장(명사형). 없으면 생략 가능",
  "bullets": ["불릿 3~5개(명사형 어미)"],
  "items": [{"label":"짧은 이름","desc":"한 줄 설명"}],
  "image_query": "photo형일 때만, 영어 검색어",
  "key": "하단 핵심 한 줄(선택, 명사형)",
  "level": "기억|이해|적용|분석|평가|창조 — 이 슬라이드가 지지하는 차시 학습목표의 인지수준",
  "numbered": false,
  "emphasis": false
}
인지수준(level) — 반드시 채운다:
- 개요의 각 블록에 있는 **'지지 목표: (인지수준) 목표N'** 줄을 그대로 옮긴다.
  그 줄이 없으면 개요 맨 앞 '차시 학습목표'의 인지수준 태그([기억]…[창조])를 보고
  이 슬라이드가 어느 목표를 지지하는지 판단해 적는다. 새 목표를 만들지 않는다.
- '지지 목표' 줄은 설계용 메타이므로 bullets 에 넣지 않는다.
- level 은 슬라이드에 인쇄되지 않는다(학습자용 화면에 설계 태그를 노출하지 않는다).
  교수자가 목표–자료 정렬을 점검하는 데 쓴다.
**인지수준 → 레이아웃 정렬(구성적 정렬 원칙)**:
- 기억: cards(용어·항목 병렬) · bullets · table(정의 목록)
- 이해: photo(예시·사례로 의미 부여) · bullets · process(흐름 설명)
- 적용: process(절차·단계) · cards(적용 조건) · quiz
- 분석: compare(두 개념 대비) · table(기준별 비교) · stat(수치 근거)
- 평가: compare(장단·타당성) · table(판단 기준) · quiz
- 창조: process(산출 절차) · cards(설계 요소) · closing(과제로 연결)
목표가 분석·평가인데 레이아웃이 단순 나열(bullets)이면 정렬 오류다 — 비교·표로 바꾼다.
문체·제목:
- 불릿·라벨은 **명사형 어미**(예 "즉시 피드백 제공" → "즉시 피드백"). 존댓말·서술체·교수자 나레이션 문장 금지.
- 제목은 간결·일관("정의와 과정"·"개요"·"논의" 같은 군더더기 제거).
- **불릿 한 줄은 32자 이내.** 길면 두 항목으로 쪼갠다(슬라이드가 넘친다).
**사진 비중 — 반드시 지킬 것**:
- 이 덱은 강의 **영상**으로 렌더된다. 도형만 이어지면 화면이 단조로워진다.
- **본문 슬라이드의 약 절반을 "photo" 로 배정한다.** 개념 설명·사례·인물·현장
  장면처럼 시각화가 도움이 되는 것은 photo 를 우선한다.
- photo 로 하면 `image_query`(영어 검색어)를 **반드시** 넣는다. 없으면 사진을 못 찾는다.
- 표·수치·순서·2단 대비처럼 **구조가 정보 자체인 것만** table/stat/process/compare 로 둔다.
- 학습목표·목차·퀴즈·마무리에는 사진을 넣지 않는다(objectives/agenda/quiz/closing).
타입 선택 규칙:
- "cover"는 만들지 마라(표지는 시스템이 맨 앞에 자동 추가).
- **학습목표 슬라이드 → "objectives"** (bullets 에 목표 2~3개). 이미지 없음.
- **목차·구성 안내 → "agenda"** (items[].label 또는 bullets).
- **퀴즈·형성평가 1문항 → "quiz"**: {"question":"문제", "choices":["보기1",…4개], "answer":"③"}.
  한 슬라이드에 한 문제. 정답은 시스템이 우하단 뱃지로 배치한다.
- **마무리·요약 → "closing"**: items 3개 {"label":"오늘의 요약|다음 차시|과제","lines":[…]}.
- 순서·단계·절차가 핵심: "process" (items[].label 2~5개). 화살표로 연결됨.
- 병렬·분류·구성요소: "cards" (items[].label + desc, 2~4개). **비순차이므로 numbered=false(기본).**
- 두 개념 대비/비교: "compare" (items 2개, 각 {"label","lines":[...]}).
- 표가 자연스러운 것: "table" (headers:[...], rows:[[...]]). 5열·8행 이내.
- 수치가 메시지인 것: "stat" (items 2~3개, {"label":"85%","desc":"설명"}).
- 도입·구간 전환: "section".
- 사진이 이해를 돕는 개념/사례/인물: "photo" + 구체적 영어 image_query(막연한 'education' 금지).
  예) 행동주의→'Pavlov conditioning experiment', 인지주의→'human brain memory diagram',
  매체→'classroom projector lesson', 인물→'portrait of a scholar'. 전체의 약 40~55%.
이미지 배치 규칙(중요):
- **objectives·quiz·closing·table·순수 정리 슬라이드는 photo 로 하지 말 것(이미지 없음).**
- 사례·실물·인물·대표 이론 슬라이드에는 photo 지정.
강조(emphasis):
- **emphasis=true 는 문제/질문 칩에만.** quiz 는 자동 강조되므로 따로 켜지 않는다.
공통:
- 같은 타입이 3장 이상 연속되지 않게 섞는다. 동일 내용이 반복되면 한 슬라이드로 합친다.
- process/cards/stat 엔 key(하단 핵심 한 줄, 명사형)를 넣어도 좋다.
"""


def _extract_json_array(text: str):
    """JSON 배열 파싱. 잘린 응답도 완결된 {…} 객체만 골라 살려낸다."""
    if not text:
        return None
    t = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", text.strip())
    i = t.find("[")
    if i == -1:
        return None
    j = t.rfind("]")
    if j > i:
        try:
            return json.loads(t[i:j + 1])
        except Exception:  # noqa: BLE001
            pass
    # 살리기: 배열 안의 완결된 최상위 { … } 객체들을 순서대로 파싱
    objs, depth, start, instr, esc = [], 0, None, False, False
    for k in range(i + 1, len(t)):
        c = t[k]
        if instr:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                instr = False
            continue
        if c == '"':
            instr = True
        elif c == "{":
            if depth == 0:
                start = k
            depth += 1
        elif c == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and start is not None:
                    try:
                        objs.append(json.loads(t[start:k + 1]))
                    except Exception:  # noqa: BLE001
                        pass
                    start = None
    return objs or None


def _outline_blocks(outline_md: str) -> List[str]:
    """개요 md 를 '### 슬라이드 …' 단위 블록 리스트로 분할."""
    parts = re.split(r"(?m)^(?=#{2,3}\s*슬라이드)", outline_md or "")
    return [p.strip() for p in parts if re.match(r"#{2,3}\s*슬라이드", p.strip())]


# 개요 블록의 '설계 메타' 키 — 슬라이드 본문이 아니므로 불릿으로 넣지 않는다.
# 이걸 걸러내지 않으면 '발표자 노트' 의 나레이션 문단이 통째로 불릿이 된다(실제로 그랬다).
_META_KEYS = ("레이아웃", "시각자료", "강조 신호", "발표자", "노트", "지지",
              "구간", "시간", "슬라이드 수", "이미지", "출처", "비고")
# '- **키**: 값' 형태의 라벨 줄
_KEY_LINE = re.compile(r"^[-*+•·]?\s*\*{0,2}([^*:：]{1,24})\*{0,2}\s*[:：]\s*(.*)$")


def _block_to_bullets(block: str) -> Dict:
    """개요 블록 하나 → bullets 슬라이드 dict(아트디렉터 실패 시 폴백).

    실제 개요는 이렇게 생겼다 — '본문 개요' 아래에 들여쓴 하위 불릿이 오고,
    그 뒤에 설계 메타 줄이 더 붙는다:

        - **레이아웃 제안**: 콘텐츠
        - **핵심 메시지(1개)**: …
        - **본문 개요**:
            - 정의: AECT …
            - '학습 촉진'과 '성과 개선'이 최종 목표
        - **시각자료 제안**: …
        - **발표자 노트(대화체)**: "자, 그럼 …"

    그래서 '본문 개요' 이후의 하위 불릿만 본문으로 삼고, 다음 라벨 줄이 나오면 멈춘다.
    """
    raw = [l for l in block.splitlines() if l.strip()]
    title = re.sub(r"^#{2,3}\s*", "", raw[0].strip()) if raw else "슬라이드"
    title = re.sub(r"^슬라이드\s*\d+\s*[—\-:：]\s*", "", title).strip()

    chip, level, bullets = "", "", []
    in_body = False
    for ln in raw[1:]:
        s = ln.strip()
        m = _KEY_LINE.match(s)
        key = m.group(1).strip() if m else ""
        val = m.group(2).strip() if m else ""

        if key.startswith("본문"):
            in_body = True
            if val:                     # 같은 줄에 인라인으로 쓴 경우
                bullets += [x.strip() for x in re.split(r"\s*[/·]\s*", val) if x.strip()]
            continue
        if m and any(k in key for k in _META_KEYS):
            if "지지" in key:           # '지지 목표: (분석) 목표2' → level 만 회수
                level = next((b for b in BLOOM_LEVELS if b in val), "")
            in_body = False             # 메타 줄이 나오면 본문 구간이 끝난 것
            continue
        if m and "핵심" in key:
            chip = val
            in_body = False
            continue

        # 라벨 없는 줄: 본문 구간 안이면 불릿, 밖이면 버린다
        if in_body:
            v = re.sub(r"^[-*+•·]\s*", "", s)
            if v:
                bullets.append(v)

    # 본문 구간을 못 찾았으면(형식이 다르면) 라벨 없는 줄을 본문으로 본다
    if not bullets:
        for ln in raw[1:]:
            s = ln.strip()
            m = _KEY_LINE.match(s)
            if m and (any(k in m.group(1) for k in _META_KEYS) or "핵심" in m.group(1)):
                continue
            v = re.sub(r"^[-*+•·]\s*", "", s)
            if v and not v.startswith("**"):
                bullets.append(v)

    return {"type": "bullets", "title": title, "chip": chip,
            "bullets": bullets[:5], "level": level}


def _fallback_plan(outline_md: str, deck_title: str) -> List[Dict]:
    """개요 md 를 단순 파싱해 bullets 위주 플랜으로(아트디렉터 실패 시)."""
    plan = [{"type": "cover", "title": deck_title, "chip": ""}]
    plan += [_block_to_bullets(b) for b in _outline_blocks(outline_md)]
    return plan


BLOOM_LEVELS = ("기억", "이해", "적용", "분석", "평가", "창조")


def _normalize_slide(s: Dict) -> Optional[Dict]:
    if not isinstance(s, dict):
        return None
    s.setdefault("type", "bullets")
    s.setdefault("title", "")
    s.setdefault("bullets", [])
    if s["type"] not in _RENDER:
        s["type"] = "bullets"
    if s["type"] == "cover":   # LLM이 표지를 만들면 내용형으로 강등
        s["type"] = "bullets"
    # 인지수준은 6수준 중 하나이거나 없음. 대괄호·설명이 붙어 와도 살려낸다.
    lv = str(s.get("level") or "")
    s["level"] = next((b for b in BLOOM_LEVELS if b in lv), "")
    return s


def level_counts(plan: List[Dict]) -> Dict[str, int]:
    """플랜의 인지수준 분포 — 교수자가 목표–자료 정렬을 눈으로 확인하는 값."""
    out: Dict[str, int] = {}
    for s in plan or []:
        lv = (s or {}).get("level") if isinstance(s, dict) else None
        if lv:
            out[lv] = out.get(lv, 0) + 1
    return {b: out[b] for b in BLOOM_LEVELS if b in out}


_CHUNK = 14  # 청크당 슬라이드 수(토큰 잘림 방지)


def plan_from_outline(generate_fn: Callable[[str, str, int], str],
                      outline_md: str, deck_title: str,
                      subtitle: str = "", course: str = "") -> List[Dict]:
    """개요를 슬라이드플랜으로 변환. 개요를 청크로 나눠 변환해 개수 잘림을 막는다.

    표지(cover)는 LLM이 아니라 여기서 맨 앞에 자동 추가한다. 청크가 실패하면
    그 청크만 블록 파싱(bullets)으로 폴백 → 전체 슬라이드 수는 항상 보존.
    """
    blocks = _outline_blocks(outline_md)
    if not blocks:
        plan = _fallback_plan(outline_md, deck_title)
        plan[0]["chip"] = subtitle
        plan[0]["band"] = course
        return plan

    chunks = [blocks[i:i + _CHUNK] for i in range(0, len(blocks), _CHUNK)]
    out: List[Dict] = []
    for ci, ch in enumerate(chunks):
        chunk_md = "\n\n".join(ch)
        user = (f"{_ART_RULES}\n\n[덱 제목]\n{deck_title}\n"
                f"[부분 {ci + 1}/{len(chunks)} · 이 부분의 슬라이드 {len(ch)}개를 빠짐없이 변환]\n\n"
                f"[슬라이드 개요]\n{chunk_md}")
        try:
            raw = generate_fn(_ART_SYS, user, 12000)
        except Exception as e:  # noqa: BLE001
            _log(f"[art] 청크 {ci + 1} 생성 오류: {e}")
            raw = ""
        arr = _extract_json_array(raw)
        got = [x for x in (_normalize_slide(s) for s in arr) if x] if isinstance(arr, list) else []
        # 개수가 모자라면 부족분을 블록 파싱으로 보충(개수 보존)
        if len(got) < len(ch):
            _log(f"[art] 청크 {ci + 1}: {len(got)}/{len(ch)} → 부족분 블록 폴백")
            got += [_block_to_bullets(b) for b in ch[len(got):]]
        out += got[:len(ch)]

    out.insert(0, {"type": "cover", "title": deck_title, "chip": subtitle,
                   "band": course})
    return out


def image_queries(plan: List[Dict]) -> Dict[int, str]:
    """photo 타입 슬라이드의 {인덱스: 영어 검색어}."""
    out = {}
    for i, s in enumerate(plan):
        if (s.get("type") == "photo") and s.get("image_query"):
            out[i] = s["image_query"]
    return out


# ── 이미지 생성 프롬프트 번들 (codex-prompt-img-studio 인풋 JSON) ──────────
# ★ 비율은 **정사각에 가깝게** 지시한다. 사진은 화면 전체가 아니라 좌우 세로 패널에
#   들어가고(4.18~4.78 × 4.69 in → 0.89~1.02), add_photo() 가 센터 크롭을 한다.
#   16:9 로 만들면 좌우가 32~41% 잘려 좌우 대조 구도가 통째로 날아간다(실측).
DEFAULT_STYLE_HINT = (
    "clean modern educational illustration, flat vector with subtle depth, "
    "consistent line weight, no gradient mesh, soft navy (#1E2761) and "
    "amber (#F2A900) accents, generous negative space, uncluttered, "
    "square 1:1 composition with the subject centered and margins on all sides "
    "(the image is center-cropped into a near-square panel)")

_IMGPROMPT_SYS = ("너는 강의 슬라이드용 이미지 생성 프롬프트 작가다. 각 슬라이드 제목/요약을 보고 "
                  "이미지 생성 모델(diffusion)용 **영어** 프롬프트 한 줄을 만든다. 그림 안에 글자·워터마크·"
                  "로고가 들어가지 않게 하고, 주제를 상징하는 구체적 장면/오브젝트로 묘사한다. "
                  "구도는 **정사각(1:1)에 가깝게** 잡고 주요 오브젝트를 가운데 모은다 — "
                  "좌우로 넓게 벌린 구도는 슬라이드에서 양옆이 잘린다. "
                  "**공통 스타일 문구는 반복하지 마라**(시스템이 뒤에 한 번 붙인다). "
                  "JSON 배열 [{\"n\":정수,\"prompt\":\"...\"}] 만 출력한다.")

# 이미지를 실제로 '올릴 자리'가 있는 타입.
#   photo   — 원래부터 사진 슬라이드
#   bullets — 이미지가 들어오면 photo 로 승격(칩·불릿이 그대로 옆으로 간다)
# cards/process/compare/table/quiz/objectives/agenda/closing/stat/section/cover 는
# 도형 레이아웃이 자리를 다 쓰므로 이미지를 넣지 않는다.
PLACEABLE_TYPES = ("photo", "bullets")


def _prompt_place(s: Dict) -> bool:
    """이 슬라이드에 이미지를 실제로 배치할 것인가(프롬프트 JSON 의 place 값).

    ★ 예전에는 emphasis 슬라이드를 빼고 있었다. 그런데 emphasis 는 **칩 색만**
      바꾼다(연한 칩 → 진한 네이비 알약). 이미지 자리와 아무 상관이 없고,
      render_photo 는 emphasis 여도 사진 자리를 그대로 예약한다.
      그래서 '자리는 예약됐는데 프롬프트가 없는' 슬라이드가 생겼고,
      받는 사람은 만들 수 없는 이미지를 기다리다 빈 액자로 인쇄했다.
      (3주차 7번·76번이 그랬다)

    지켜야 할 불변식: **자리를 예약하는 슬라이드에는 반드시 프롬프트가 있다.**
    자리는 type=photo 일 때 예약되고, bullets 는 이미지가 오면 photo 로 승격되므로
    둘 다 프롬프트 대상이다 = PLACEABLE_TYPES.
    """
    return (s.get("type") or "bullets") in PLACEABLE_TYPES


def photo_slot_slides(plan: List[Dict]) -> List[int]:
    """사진 **자리가 예약되는** 슬라이드의 1-based 번호.

    빈 액자로 남을 수 있는 자리를 세는 데 쓴다 — 조용히 비면 인쇄까지 간다.
    """
    return [i + 1 for i, s in enumerate(plan or [])
            if (s.get("type") or "bullets") == "photo"]


def apply_images(plan: List[Dict], images: Dict[int, bytes]):
    """외부 이미지를 플랜에 붙인다. bullets 슬라이드는 photo 로 승격.

    반환: (새 플랜, 사용된 {인덱스: 바이트}, 자리 없어 건너뛴 1-based 번호 목록)
    이미지를 조용히 버리지 않고 무엇이 안 들어갔는지 항상 돌려준다.
    """
    out, used, skipped = [], {}, []
    for i, s in enumerate(plan or []):
        s = dict(s) if isinstance(s, dict) else {"type": "bullets", "title": str(s)}
        if i in (images or {}):
            typ = (s.get("type") or "bullets").lower()
            if typ in PLACEABLE_TYPES:
                s["type"] = "photo"
                used[i] = images[i]
            else:
                skipped.append(i + 1)
        out.append(s)
    return out, used, skipped


def image_prompt_bundle(plan: List[Dict], deck_title: str, *,
                        generate_fn: Optional[Callable[[str, str, int], str]] = None,
                        style_hint: str = DEFAULT_STYLE_HINT,
                        have_photos: Optional[set] = None) -> Dict:
    """**이미지를 실제로 올릴 수 있는 슬라이드만** 프롬프트로 뽑아 1개 번들(dict)로.

    전 슬라이드를 뽑던 것을 좁힌 이유: 배치 가능한 타입은 photo·bullets 뿐이고
    (나머지는 도형 레이아웃이 자리를 다 쓴다), 42장을 뽑으면 34장이 '자리 없어
    건너뜀' 으로 버려졌다. 만들지도 못할 프롬프트에 LLM 호출을 3배 쓰고,
    받는 사람은 42장을 그리고 34장을 헛수고했다.

    generate_fn 이 있으면 LLM으로 영어 프롬프트를 짓고(청크·폴백), 없으면
    image_query/제목 기반으로 결정론적으로 만든다. 'prompts' 배열은 위→아래
    순차 실행용이고, "n" 은 **1-based 슬라이드 번호** = 이미지 파일명 규칙의
    번호다(003.png → 3번 슬라이드). 그래서 걸러내도 번호는 원래 것을 유지한다.
    """
    total = len(plan)
    # 자동 사진이 이미 붙은 슬롯은 뺀다 — 그건 그릴 필요가 없다.
    # 남는 것 = "자리는 있는데 사진을 못 찾은 슬라이드" 뿐이고, 그게 그려야 할 목록이다.
    got = set(int(i) for i in (have_photos or ()))
    idx = [i for i, s_ in enumerate(plan) if _prompt_place(s_) and i not in got]
    n = len(idx)
    en: Dict[int, str] = {}
    if generate_fn and idx:
        for a in range(0, n, _CHUNK):
            ch = idx[a:a + _CHUNK]
            lines = "\n".join(
                f'{i + 1}. {plan[i].get("title", "")} | {plan[i].get("chip", "")}' for i in ch)
            user = (f"[공통 스타일] {style_hint}\n"
                    f"아래 슬라이드 {len(ch)}개 각각에 대해 n(슬라이드 번호)과 영어 이미지 프롬프트를 만들어라. "
                    f"프롬프트에 공통 스타일을 녹이고, 그림 안 텍스트 금지.\n\n{lines}")
            try:
                arr = _extract_json_array(generate_fn(_IMGPROMPT_SYS, user, 6000))
            except Exception as e:  # noqa: BLE001
                _log(f"[imgprompt] 청크 오류: {e}")
                arr = None
            if isinstance(arr, list):
                for o in arr:
                    if isinstance(o, dict) and o.get("n") and o.get("prompt"):
                        en[int(o["n"]) - 1] = str(o["prompt"]).strip()

    prompts = []
    for i in idx:
        s = plan[i]
        subj = (en.get(i) or s.get("image_query")
                or f"conceptual illustration for '{s.get('title', '')}'")
        subj = str(subj).strip().rstrip(".").strip()       # 마침표 중복 방지
        # LLM 이 공통 스타일을 이미 녹여 넣는 경우가 있다 — 그대로 또 붙이면
        # 같은 문구가 두 번 들어가 프롬프트가 두 배로 길어지고 주제가 묻힌다.
        tail = "" if "flat vector" in subj.lower() else f" {style_hint}."
        prompt = f"{subj}.{tail} No text, no watermark, no logo."
        prompts.append({
            "n": i + 1,
            "title": s.get("title", ""),
            "type": s.get("type", "bullets"),
            # 인지수준을 함께 흘린다 — 이 번들이 슬라이드→비디오 단계의 입력이라,
            # 씬별 나레이션 길이·속도를 정할 근거가 된다(기억<이해<분석 순으로 체류).
            "level": s.get("level", ""),
            "prompt": prompt,
            "negative": "text, letters, watermark, logo, low quality, distorted",
            "keywords": ([s["image_query"]] if s.get("image_query") else []),
            "place": _prompt_place(s),
        })
    return {
        "deck": deck_title,
        "style_hint": style_hint,
        # 슬라이드의 사진 자리는 세로 패널(≈1:1)이다. landscape 로 두면
        # 생성 도구가 3:2 를 골라 좌우 30~40% 가 잘린다.
        "aspect": "square",
        "target_box": "near-square vertical panel (~1:1), center-cropped",
        "count": n,                 # 뽑은 프롬프트 수 = 그려야 할 장수
        "deck_slides": total,       # 덱 전체 장수 (번호 n 은 이 기준이다)
        "photos_found": len(got),   # 자동 사진으로 이미 채워진 장수(프롬프트에서 제외됨)
        "file_naming": ("생성한 이미지는 images/ 폴더에 '슬라이드번호'로 저장하세요. "
                        "예: 003.png → 3번 슬라이드. 파일명 앞 숫자만 맞으면 되고 "
                        "확장자·뒤 설명은 자유입니다(003_뇌구조.png)."),
        "prompts": prompts,
    }
