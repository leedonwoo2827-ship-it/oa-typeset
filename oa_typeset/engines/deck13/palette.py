# -*- coding: utf-8 -*-
r"""덱 배색 — **견본 PPTX 한 장에서 색과 글꼴을 뽑아** 렌더러에 주입한다.

AIM 은 포맷마다 양식 파일을 **하나씩만** 둔다(assets/hwpx_template.hwpx ·
docx_template.docx · pptx_template.pptx). 덱도 같다 — 배색을 고르는 화면은 없고,
디자인팀이 `assets/pptx_template.pptx` 를 갈아 끼우면 색이 따라온다.

배자(좌표)는 읽지 않는다. 내용 길이가 견본과 다르면 넘치거나 빈 공간이 생기고,
그건 builder 의 _fit() 이 이미 처리한다. 여기서 가져오는 건 **색과 글꼴** 뿐이다.

뽑는 법 ① — 이름표(권장)
    시안에 이름이 `pal:navy` · `pal:amber` · `pal:chip` · `pal:navy_55` 인 도형을
    두면 **그 채움색을 그대로 읽는다.** `pal:ink` 처럼 회색 계열도 선언할 수 있다.
    tools/build_pptx_template.py 가 만드는 마지막 '배색' 장이 그것이다.

    면적 서명(아래)만으로는 못 버티는 경우가 실제로 나왔다 — 렌더러가 대비 보정으로
    표 머리행을 `navy_dk`(#1D854E) 로 채우는데, 그 색이 주색(#1F9155) 보다 어두워서
    되읽을 때 navy 자리를 빼앗았다(그리고 주색이 amber 로 밀렸다). 시안을 엔진으로
    생성하는 이상 파생 다크가 화면에 섞이는 건 정상이므로, 선언을 우선한다.

뽑는 법 ② — 면적 서명
    견본의 도형 채움색을 **총 면적** 순으로 세우고, 무채색(회색 계열)을 뺀 뒤
    위에서부터 navy · chip · amber · navy_55 자리에 앉힌다. 실측으로 검증했다:

        레이아웃견본_green   1F9155(45.0) · E3F5EA(21.8) · 2A9760(15.5) · 93D0B2(5.2)
        레이아웃견본_cyan    0C93B8(45.0) · E1F2F7(21.8) · 1799B7(15.5) · 93D2E2(5.2)

    골격이 같고 색만 다르므로 면적 서명이 그대로 자리를 알려 준다.
    나머지(navy_75·navy_dk·amber_dk·amber_wash)는 주색에서 계산한다.
    회색 계열(navy_30·navy_12·ink·on_navy)은 배색과 무관하게 고정이다.
"""
from __future__ import annotations

import os
import re
import threading
from collections import defaultdict
from typing import Dict, Optional, Tuple

# 견본을 못 읽을 때 쓰는 기본값 — 현재 견본(레이아웃견본_green)에서 뽑은 값이다.
BASELINE: Dict[str, str] = {
    "navy": "1F9155", "navy_75": "57AC80", "navy_55": "93D0B2",
    "chip": "E3F5EA", "amber": "2A9760", "amber_dk": "258554",
    "amber_wash": "E5F3EC", "navy_dk": "176A3C",
    "navy_30": "CBD5E1", "navy_12": "E2E8F0", "on_navy": "F1F5F9",
    "white": "FFFFFF", "ink": "334155", "muted": "94A3B8", "ink_mid": "475569",
}
# 배색이 바뀌어도 따라 움직이지 않는 회색 계열.
#
# ★ muted(#94A3B8) 를 여기 둔 이유 — 라벨·캡션·푸터는 예전에 navy_55 를 썼다.
#   navy_55 는 **주색의 옅은 틴트**라 회사 양식(그린) 에서 #93D0B2 가 되고,
#   흰 바탕 대비가 1.9:1 로 떨어져 글자가 사실상 안 보였다. 보조 글자는 배색을
#   따라가면 안 된다 — 슬레이트 회색으로 고정한다(정답지 3종도 #94A3B8 이다).
NEUTRAL: Dict[str, str] = {
    "navy_30": "CBD5E1", "navy_12": "E2E8F0", "on_navy": "F1F5F9",
    "white": "FFFFFF", "ink": "334155", "muted": "94A3B8", "ink_mid": "475569",
}
BASELINE_FONT = "Noto Sans KR"

# _HEX 를 제자리에서 바꾸므로 빌드가 겹치면 색이 섞인다. 빌드는 순간이지만
# 두 요청이 동시에 들어올 수 있어 락을 둔다.
_LOCK = threading.RLock()
_QN = None                          # pptx.oxml.ns.qn — _scan 에서 채운다
# 견본 파싱은 (경로, mtime, 크기) 로 캐시한다 — 덱 한 장마다 다시 읽을 이유가 없다.
_CACHE: Dict[Tuple[str, float, int], Tuple[Dict[str, str], Optional[str]]] = {}


# ── 색 계산 ──────────────────────────────────────────────────────────────
def _rgb(h: str) -> Tuple[int, int, int]:
    h = (h or "").lstrip("#").strip()
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _hex(t) -> str:
    return "%02X%02X%02X" % tuple(max(0, min(255, int(round(c)))) for c in t)


def _lum(t) -> float:
    def lin(c):
        c /= 255.0
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (lin(x) for x in t)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _ratio_white(t) -> float:
    return 1.05 / (_lum(t) + 0.05)


def _mix(t, other, k: float):
    """t 를 other 쪽으로 k 만큼 섞는다."""
    return tuple(a + (b - a) * k for a, b in zip(t, other))


def _chroma(t) -> int:
    """최대·최소 채널 차."""
    return max(t) - min(t)


def _sat(t) -> float:
    """상대 채도(HSV 의 S). 절대 채도로 거르면 안 된다 — 회색 계열 #475569 도
    채널 차가 34라 절대 기준을 통과해 브랜드색 행세를 한다(실측: navy 를 뺏겼다)."""
    m = max(t)
    return 0.0 if m == 0 else (m - min(t)) / m


def _hue(t) -> float:
    """색상(0~360). 브랜드색과 회색 계열을 가르는 데 쓴다 —
    슬레이트 회색은 213° 근처에 모이고 견본의 주색은 그와 떨어져 있다."""
    r, g, b = (x / 255.0 for x in t)
    mx, mn = max(r, g, b), min(r, g, b)
    d = mx - mn
    if d == 0:
        return 0.0
    if mx == r:
        h = ((g - b) / d) % 6
    elif mx == g:
        h = (b - r) / d + 2
    else:
        h = (r - g) / d + 4
    return h * 60.0


def _hue_near(a: float, b: float, tol: float = 15.0) -> bool:
    d = abs(a - b) % 360
    return min(d, 360 - d) <= tol


def _darken_to(t, min_ratio: float = 4.5):
    """흰 글씨를 얹을 수 있을 때까지 어둡게. 밴드 배경용."""
    cur = tuple(t)
    for _ in range(24):
        if _ratio_white(cur) >= min_ratio:
            break
        cur = _mix(cur, (0, 0, 0), 0.08)
    return cur


def _derive(navy_h: str, chip_h: str, amber_h: str, navy55_h: str) -> Dict[str, str]:
    """네 자리에서 나머지 토큰을 계산한다."""
    navy, amber = _rgb(navy_h), _rgb(amber_h)
    white = (255, 255, 255)
    out = dict(NEUTRAL)
    out.update({
        "navy": navy_h,
        "navy_75": _hex(_mix(navy, white, 0.25)),
        "navy_55": navy55_h or _hex(_mix(navy, white, 0.55)),
        "chip": chip_h or _hex(_mix(navy, white, 0.90)),
        "amber": amber_h,
        "amber_dk": _hex(_darken_to(amber)),
        "amber_wash": _hex(_mix(amber, white, 0.90)),
        "navy_dk": _hex(_darken_to(navy)),
    })
    return out


# 시안이 배색을 직접 선언하는 도형 이름 접두어.
_PAL_PREFIX = "pal:"


def _declared(prs, srgb) -> Dict[str, str]:
    """이름이 `pal:<토큰>` 인 도형의 채움색 = 시안이 **선언한** 배색."""
    out: Dict[str, str] = {}
    for slide in prs.slides:
        for sh in slide.shapes:
            nm = (sh.name or "").strip().lower()
            if not nm.startswith(_PAL_PREFIX):
                continue
            key = nm[len(_PAL_PREFIX):].strip()
            if not key:
                continue
            try:
                spPr = sh._element.find(_QN("p:spPr"))
                solid = None if spPr is None else spPr.find(_QN("a:solidFill"))
                v = None if solid is None else srgb(solid)
            except Exception:  # noqa: BLE001
                v = None
            if v:
                out[key] = v
    return out


# ── 견본 PPTX 읽기 ───────────────────────────────────────────────────────
def _scan(path: str) -> Tuple[Dict[str, str], Optional[str]]:
    """견본에서 (색 토큰, 글꼴). 못 읽으면 (BASELINE, BASELINE_FONT)."""
    global _QN
    try:
        from pptx import Presentation
        from pptx.oxml.ns import qn
        _QN = qn
    except Exception:
        return dict(BASELINE), BASELINE_FONT

    def srgb(el):
        for c in el.iter(qn("a:srgbClr")):
            v = c.get("val")
            if v and re.fullmatch(r"[0-9A-Fa-f]{6}", v):
                return v.upper()
        return None

    try:
        prs = Presentation(path)
    except Exception:
        return dict(BASELINE), BASELINE_FONT

    area: Dict[str, float] = defaultdict(float)
    fonts: Dict[str, int] = defaultdict(int)
    for slide in prs.slides:
        for sh in slide.shapes:
            spPr = sh._element.find(qn("p:spPr"))
            if spPr is not None:
                solid = spPr.find(qn("a:solidFill"))
                if solid is not None:
                    v = srgb(solid)
                    if v:
                        area[v] += ((sh.width or 0) / 914400) * ((sh.height or 0) / 914400)
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        rPr = r._r.find(qn("a:rPr"))
                        if rPr is not None:
                            latin = rPr.find(qn("a:latin"))
                            if latin is not None and latin.get("typeface"):
                                fonts[latin.get("typeface")] += 1

    font = max(fonts, key=fonts.get) if fonts else BASELINE_FONT

    # ★ 선언이 있으면 그것이 이긴다. navy·amber 두 자리만 있으면 나머지는 계산한다.
    decl = _declared(prs, srgb)
    if "navy" in decl and "amber" in decl:
        cols = _derive(decl["navy"], decl.get("chip", ""), decl["amber"],
                       decl.get("navy_55", ""))
        cols.update(decl)          # chip·navy_55·회색 계열까지 선언대로
        return cols, font

    by_area = [c for c, _a in sorted(area.items(), key=lambda t: -t[1])]

    # ① 진한 브랜드색만 먼저 고른다(상대 채도 0.45 이상).
    core = [c for c in by_area if _sat(_rgb(c)) >= 0.45]
    if not core:
        return dict(BASELINE), font

    # ② 브랜드 색상(hue) — 가장 진한 색의 것. 이후 판정의 기준이 된다.
    brand = _hue(_rgb(max(core, key=lambda c: _sat(_rgb(c)))))

    # ③ 자리 배정. **밝기**로 정한다 — 면적 순위로 정하면 견본의 표지 도형
    #    하나만 바꿔도 navy 와 amber 가 뒤바뀐다(표지 전면 밴드 35.6in² 를
    #    회사 양식대로 얇은 바로 줄이자 실제로 그랬다).
    navy_h = min(core, key=lambda c: _lum(_rgb(c)))          # 가장 어두운 것
    rest = [c for c in core if c != navy_h]                  # 면적 순 유지
    amber_h = rest[0] if rest else navy_h                    # 남은 것 중 면적 1위

    # ④ chip·navy_55 는 옅어서 채도 기준을 못 넘는다. 색상으로 찾는다.
    tinted = [c for c in by_area
              if _sat(_rgb(c)) >= 0.07 and _hue_near(_hue(_rgb(c)), brand)]
    chip_cands = [c for c in tinted if _lum(_rgb(c)) > 0.60]
    chip_h = max(chip_cands, key=lambda c: _lum(_rgb(c))) if chip_cands else ""
    navy55_h = next((c for c in tinted if 0.30 < _lum(_rgb(c)) < 0.60), "")

    return _derive(navy_h, chip_h, amber_h, navy55_h), font


def read(path: Optional[str]) -> Tuple[Dict[str, str], Optional[str]]:
    """견본 경로 → (색 토큰, 글꼴). 경로가 없으면 기본값."""
    if not path or not os.path.isfile(path):
        return dict(BASELINE), BASELINE_FONT
    try:
        st = os.stat(path)
        key = (os.path.abspath(path), st.st_mtime, st.st_size)
    except OSError:
        return dict(BASELINE), BASELINE_FONT
    with _LOCK:
        if key not in _CACHE:
            _CACHE.clear()                  # 견본은 한 개다 — 캐시도 한 칸이면 된다
            _CACHE[key] = _scan(path)
        return _CACHE[key]


# ── 렌더러에 주입 ────────────────────────────────────────────────────────
def apply(sample_path: Optional[str]) -> Dict[str, str]:
    """builder 의 색·글꼴을 견본에 맞춘다. 적용된 색 토큰을 돌려준다.

    ★ 기본 배색을 **한 번만** 스냅샷해 둔다. 그래야 견본을 바꿔 가며 여러 번
      빌드해도 빠진 키가 '직전 견본의 색' 이 아니라 항상 기본색으로 채워진다.
    """
    from . import builder as db

    with _LOCK:
        if not hasattr(db, "_HEX_BASE"):
            db._HEX_BASE = dict(db._HEX)

        colors, font = read(sample_path)
        merged = dict(db._HEX_BASE)
        for k, v in colors.items():
            if k in db._HEX_BASE:
                try:
                    merged[k] = _rgb(v)
                except (ValueError, IndexError):
                    pass
        db._HEX.clear()
        db._HEX.update(merged)

        db.TEMPLATE_FONT = font or None
        db.TEMPLATE_NAME = os.path.basename(sample_path) if sample_path else "기본"
        return colors
