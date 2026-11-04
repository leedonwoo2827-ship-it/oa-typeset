# -*- coding: utf-8 -*-
r"""assets/pptx_template.pptx (발표자료 시안) 를 다시 만든다.

왜 스크립트인가
    시안은 .pptx 바이너리라 손으로 고치면 무엇이 바뀌었는지 리뷰가 안 된다.
    그래서 **덱 엔진(deck13.builder) 의 상수·헬퍼를 그대로 불러** 견본 슬라이드를
    그린다. 시안과 엔진이 같은 격자를 쓰므로, 엔진을 고치고 이걸 다시 돌리면
    시안이 저절로 따라온다 — 둘이 어긋나 있던 문제가 재발하지 않는다.

엔진이 시안에서 읽는 것 (palette.py)
    **색과 글꼴 이름뿐이다.** 좌표는 읽지 않는다. 마지막 '배색' 장의 사각형은
    이름이 `pal:navy` · `pal:amber` 처럼 붙어 있고, palette._declared() 가 그
    이름을 보고 색을 그대로 가져간다. 디자이너는 그 장의 사각형 색만 갈면 된다.

    이름표가 필요한 이유 — 예전에는 면적·휘도 서명으로 자리를 추측했다. 그런데
    렌더러가 대비 보정으로 표 머리행을 navy_dk(#1D854E) 로 채우고, 그 색이 주색
    (#1F9155) 보다 어두워서 되읽을 때 navy 자리를 빼앗았다(주색은 amber 로 밀렸다).
    시안을 엔진으로 생성하는 이상 파생 다크가 화면에 섞이는 건 정상이므로,
    추측을 버리고 선언을 읽는다. 이름표가 없는 견본은 예전 서명 방식으로 폴백한다.

쓰는 법
    python tools/build_pptx_template.py            # assets/pptx_template.pptx 갱신
    python tools/build_pptx_template.py --out x.pptx --check-only
"""
from __future__ import annotations

import argparse
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

ASSETS = os.path.join(ROOT, "oa_typeset", "assets")          # 글꼴(fonts/) 자리 — 없으면 이름만 지정
TEMPLATE = os.path.join(ROOT, "oa_typeset", "forms", "보고서형-A4", "form.pptx")
LOGO_R = os.environ.get("OA_LOGO_RIGHT", "")   # 로고를 넣고 싶을 때만
LOGO_L = os.environ.get("OA_LOGO_LEFT", "")

# ── 견본 슬라이드플랜 ──────────────────────────────────────────────────────
# 13종 레이아웃 + 1차 정답지 3종에서 확인한 컴포넌트를 한 장씩 보여 준다.
# 문구는 '…을 입력해 주세요' 로 두어 시안임을 분명히 한다.
_L = "본문 항목을 입력해 주세요"
_S = "한 줄 요약을 입력해 주세요"
_N = "보조 설명 한 줄을 입력해 주세요 · 근거와 범위를 적습니다"

PLAN = [
    {"type": "cover",
     "title": "표지 제목과 *핵심 낱말*을 입력해 주세요",
     "subtitle": "부제 한 줄을 입력해 주세요 · 문서 전체 이름을 여기에 적습니다",
     "presenter": "발표자 이름", "affiliation": "(소속 기관)",
     "footer": "기관명"},

    {"type": "agenda", "kicker": "OVERVIEW", "title": "목차",
     "items": [{"part": "Part I", "label": "첫 번째 대영역 제목"},
               {"part": "Part II", "label": "두 번째 대영역 제목"},
               {"part": "Part III", "label": "세 번째 대영역 제목"},
               {"part": "Part IV", "label": "네 번째 대영역 제목"},
               {"part": "Part V", "label": "다섯 번째 대영역 제목"}]},

    {"type": "section", "title": "간지 제목을 입력해 주세요", "chip": "PART I"},

    {"type": "bullets", "kicker": "PART I · 라벨",
     "title": "서술 슬라이드 — 제목에 *결론*을 적습니다",
     "chip": _S, "note": _N,
     "bullets": [_L, _L, _L, _L], "key": "핵심 문장을 입력해 주세요"},

    {"type": "bullets", "kicker": "PART I · 라벨",
     "title": "서술 + *수치 레일*",
     "chip": _S, "note": "우측 레일에 근거 수치를 세로로 세웁니다",
     "bullets": [_L, _L, _L, _L, _L],
     "stats": [{"label": "지표 이름", "value": "00.00", "desc": "기준 시점 · 목표 대비"},
               {"label": "지표 이름", "value": "00.00", "desc": "기준 시점 · 목표 대비"},
               {"label": "지표 이름", "value": "+0.00", "desc": "기초선 대비 변화폭"}]},

    {"type": "cards", "kicker": "PART II · 라벨", "title": "병렬 항목 카드",
     "chip": _S,
     "items": [{"label": "항목 %d" % i, "desc": "설명 한 줄을 입력해 주세요"}
               for i in (1, 2, 3, 4)],
     "key": "핵심 문장을 입력해 주세요"},

    {"type": "compare", "kicker": "PART II · 라벨", "title": "두 대상 비교",
     "chip": _S,
     "items": [{"label": "비교 대상 1", "lines": [_L, _L, _L]},
               {"label": "비교 대상 2", "lines": [_L, _L, _L]}]},

    {"type": "process", "kicker": "PART III · 라벨", "title": "순서·단계 흐름",
     "chip": _S, "layout": "box",
     "items": [{"label": "단계 %d" % i, "desc": "한 줄"} for i in (1, 2, 3, 4)],
     "key": "핵심 문장을 입력해 주세요"},

    {"type": "process", "kicker": "PART III · 라벨", "title": "번호 밴드 스텝",
     "chip": _S, "layout": "band",
     "items": [{"label": "기준 시점 %d" % i,
                "desc": "설명이 한 문장을 넘어가면 가로 박스에 안 들어가므로 이 판을 씁니다"}
               for i in (1, 2, 3, 4)]},

    {"type": "table", "kicker": "PART IV · 라벨", "title": "표 — 기본",
     "chip": _S, "note": _N,
     "subhead": "표 이름을 입력해 주세요",
     "source": "단위: 표기값 기준 · 출처: 00_원자료.xlsx",
     "headers": ["머리글 1", "머리글 2", "머리글 3", "설명이 들어가는 열"],
     "rows": [["항목 %s" % c, "00.0", "00.0", "설명 문장을 입력해 주세요"]
              for c in "ABCDE"]},

    {"type": "table", "kicker": "PART IV · 라벨", "title": "표 — 좌우 2개",
     "chip": _S,
     "tables": [
         {"subhead": "왼쪽 표 이름", "source": "기준 시점 · 출처: 00_원자료.xlsx",
          "headers": ["구분", "값 1", "값 2"],
          "rows": [["항목 %s" % c, "000", "0"] for c in "ABCD"]},
         {"subhead": "오른쪽 표 이름", "source": "기준 시점 · 출처: 00_원자료.xlsx",
          "headers": ["구분", "값 1", "값 2"],
          "rows": [["항목 %s" % c, "000", "0"] for c in "ABCD"]}]},

    {"type": "table", "kicker": "PART IV · 라벨",
     "title": "표 — 밀집(글꼴 하한 시연)",
     "chip": "열·행이 많으면 최소 글꼴에서 멈추고 장을 나누라고 알립니다",
     "headers": ["구분", "값1", "값2", "값3", "값4", "값5", "값6", "합계", "비율"],
     "rows": [["항목 %02d" % i, "80", "70", "150", "72", "61", "133", "283",
               "0.0%"] for i in range(1, 11)]},

    {"type": "stat", "kicker": "PART V · 라벨", "title": "목표 대비 *실적*",
     "chip": _S, "note": _N,
     "items": [{"label": "00%", "kind": "target", "caption": "목표치",
                "desc": "목표 산정 근거 한 줄"},
               {"label": "00.0%", "caption": "관측 최소", "desc": "관측 대상 한 줄"},
               {"label": "00.0%", "caption": "관측 최대", "desc": "관측 대상 한 줄"}],
     "interpretation": ["판정 근거 문장을 입력해 주세요.",
                        "달성·미달성을 나누어 적습니다."]},

    {"type": "chart", "kicker": "PART V · 라벨", "title": "차트 — 1개",
     "chip": _S,
     "charts": [{"caption": "차트 이름을 입력해 주세요", "kind": "column",
                 "categories": ["기준 1", "기준 2", "기준 3", "기준 4"],
                 "series": [{"name": "계열 1", "values": [30.0, 38.5, 44.2, 52.1]},
                            {"name": "계열 2", "values": [29.4, 31.0, 32.8, 34.6]}],
                 "unit": "단위: 표기값 기준",
                 "source": "출처: 00_원자료.xlsx"}]},

    {"type": "chart", "kicker": "PART V · 라벨", "title": "차트 — 좌우 2개",
     "chip": _S,
     "charts": [{"caption": "왼쪽 차트 이름", "kind": "column",
                 "categories": ["구분 1", "구분 2"],
                 "series": [{"name": "계열 1", "values": [52.1, 47.3]},
                            {"name": "계열 2", "values": [49.8, 45.1]}],
                 "unit": "단위: 표기값 기준"},
                {"caption": "오른쪽 차트 이름", "kind": "line",
                 "categories": ["기준 1", "기준 2", "기준 3", "기준 4"],
                 "series": [{"name": "계열 1", "values": [58.2, 62.4, 65.9, 70.9]}],
                 "unit": "단위: 표기값 기준"}]},

    {"type": "photo", "kicker": "PART V · 라벨", "title": "그림 — 1장",
     "chip": _S, "bullets": [_L, _L, _L]},

    {"type": "photo", "kicker": "PART V · 라벨", "title": "그림 — 2장 나란히",
     "layout": "2up", "chip": _S,
     "captions": ["그림 설명을 입력해 주세요", "그림 설명을 입력해 주세요"]},

    {"type": "closing", "kicker": "CLOSING", "title": "마무리",
     "chip": _S, "note": _N,
     "bullets": ["결론 문장을 입력해 주세요"] * 4,
     "panel": {"label": "[확인 필요 항목]",
               "lines": ["확인이 필요한 항목을 입력해 주세요"] * 2}},

    {"type": "thanks", "title": "감사합니다."},
]

# ── 팔레트 장 ──────────────────────────────────────────────────────────────
# 채움으로 그리는 색 = 추출 서명을 만드는 색. 파생색은 **글자로만** 적는다.
SWATCH = [
    ("navy",     "주색 — 제목 강조·표 머리행·마지막 단계", 1.75),
    ("amber",    "강조색 — 룰·액센트 바·밴드 좌측 바",     4.20),
    ("navy_55",  "주색 옅은 단계 — 카드·스텝 진행",        1.75),
    ("chip",     "아주 옅은 면 — 표지 우측 패널·짝수행",   4.20),
]
GREYS = [("ink", "본문 글자"), ("ink_mid", "보조 글자"),
         ("muted", "라벨·캡션·푸터"), ("navy_30", "대형 연번·페이지"),
         ("navy_12", "구분선·헤어라인"), ("on_navy", "카드·밴드 바탕")]
DERIVED = [("navy_75", "주색에서 계산 — 카드·계열 2단계"),
           ("navy_dk", "주색을 어둡게 — 글자 전용"),
           ("amber_dk", "강조색을 어둡게 — 글자 전용"),
           ("amber_wash", "강조 워시 — 해석 밴드 바탕")]


def draw_palette(db, slide) -> None:
    """마지막 장 — 배색 스와치. 이 장이 추출 서명을 못박는다."""
    x0, y = db.MARGIN, 0.62
    tf = db._text(slide, x0, y, db.CONTENT_W, 0.52, anchor="m")
    db._T(db._para(tf, True), "배색 — 이 장의 사각형 색을 갈면 산출물이 따라옵니다",
          "title", size=20)
    tf = db._text(slide, x0, y + 0.52, db.CONTENT_W, 0.62, anchor="t")
    db._T(db._para(tf, True, line=1.32),
          "엔진은 시안에서 색과 글꼴만 읽습니다(좌표는 읽지 않습니다). "
          "사각형의 도형 이름이 pal:navy · pal:amber … 이고 그 이름으로 색을 가져갑니다.",
          "note")
    db._T(db._para(tf, False, line=1.32),
          "이름은 그대로 두고 채움색만 바꾸세요. 그다음 "
          "tools/build_pptx_template.py 를 다시 돌리면 견본도 새 색으로 그려집니다.",
          "note")

    # 추출되는 네 자리
    cy = 1.62
    cx = x0
    for name, use, w in SWATCH:
        # 도형 이름이 배색 선언이다 — palette._declared() 가 이 이름을 읽는다.
        db._rect(slide, cx, cy, w, 1.05, name).name = "pal:%s" % name
        tf = db._text(slide, cx, cy + 1.10, w, 0.28, anchor="m")
        db._run(db._para(tf, True), "%s  #%s" % (name, db._hexstr(name)),
                11.0, "ink", "sb")
        tf = db._text(slide, cx, cy + 1.36, w, 0.30, anchor="t")
        db._T(db._para(tf, True, line=1.25), use, "note", size=9.5)
        cx += w + 0.24

    # 고정 회색 — 배색을 바꿔도 따라 움직이지 않는다
    cy = 3.55
    tf = db._text(slide, x0, cy, db.CONTENT_W, 0.30, anchor="m")
    db._run(db._para(tf, True), "고정 회색 — 배색과 무관하게 유지됩니다",
            11.5, "ink", "sb")
    cy += 0.34
    w = (db.CONTENT_W - 0.20 * (len(GREYS) - 1)) / len(GREYS)
    for i, (name, use) in enumerate(GREYS):
        cx = x0 + i * (w + 0.20)
        db._rect(slide, cx, cy, w, 0.62, name).name = "pal:%s" % name
        tf = db._text(slide, cx, cy + 0.66, w, 0.26, anchor="m")
        db._run(db._para(tf, True), "%s  #%s" % (name, db._hexstr(name)),
                9.5, "ink", "sb")
        tf = db._text(slide, cx, cy + 0.90, w, 0.26, anchor="t")
        db._T(db._para(tf, True), use, "note", size=9.0)

    # 파생색 — 채움으로 그리지 않는다(navy 자리를 빼앗는다)
    cy = 5.55
    tf = db._text(slide, x0, cy, db.CONTENT_W, 0.30, anchor="m")
    db._run(db._para(tf, True),
            "파생색 — 주색에서 계산됩니다. 어두운 그린은 도형 채움으로 쓰지 마세요"
            " (navy 자리를 빼앗습니다)", 11.5, "ink", "sb")
    cy += 0.34
    # 파생색은 **작은 견본 사각형 대신 색점**으로만 보여 준다 — 채움으로 크게 깔면
    # 어두운 그린이 navy 자리를 빼앗는다(이름표가 없는 견본으로 폴백했을 때).
    w = (db.CONTENT_W - 0.24 * (len(DERIVED) - 1)) / len(DERIVED)
    for i, (name, use) in enumerate(DERIVED):
        cx = x0 + i * (w + 0.24)
        db._oval(slide, cx, cy + 0.04, 0.20, name)
        tf = db._text(slide, cx + 0.28, cy, w - 0.28, 0.28, anchor="m")
        db._run(db._para(tf, True), "%s  #%s" % (name, db._hexstr(name)),
                9.5, "ink", "sb")
        tf = db._text(slide, cx, cy + 0.30, w, 0.30, anchor="t")
        db._T(db._para(tf, True, line=1.25), use, "note", size=9.0)


# ── 빌드 ───────────────────────────────────────────────────────────────────
def build(out_path: str, sample: str) -> bytes:
    from pptx import Presentation
    from oa_typeset.engines.deck13 import builder as db

    data = db.build_deck(
        PLAN, template_path=(sample if os.path.isfile(sample) else None),
        deck_title="발표자료 시안", footer="문서명 · 장 제목",
        logo_path=LOGO_R or None,
        logo2_path=LOGO_L or None,
        assets_dir=ASSETS, embed_font=True, report=[])
    if not data:
        raise SystemExit("build_deck 실패 — python-pptx 가 없나?")

    # 팔레트 장은 시안에만 있는 장이다(빌드 때 어차피 전부 지워진다).
    prs = Presentation(io.BytesIO(data))
    slide = prs.slides.add_slide(db._blank_layout(prs))
    db._strip_placeholders(slide)
    draw_palette(db, slide)
    if LOGO_R or LOGO_L:
        db.add_logos(slide, LOGO_R or None, LOGO_L or None)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def check(path: str, expect_colors, expect_font) -> int:
    """새 시안에서 색·글꼴을 되읽어 기존과 같은지 확인한다."""
    from oa_typeset.engines.deck13 import palette as pal
    pal._CACHE.clear()
    got, font = pal.read(path)
    bad = [(k, v, got.get(k)) for k, v in expect_colors.items() if got.get(k) != v]
    extra = sorted(set(got) - set(expect_colors))
    print("  글꼴: %s%s" % (font, "" if font == expect_font
                            else "  ← 달라졌다! (기대 %s)" % expect_font))
    if extra:
        print("  새 키: %s" % ", ".join(extra))
    if bad or font != expect_font:
        for k, want, got_v in bad:
            print("  %-11s 기대 %s → 실제 %s" % (k, want, got_v))
        print("  ✗ 배색 서명이 깨졌다 — 팔레트 장의 사각형 면적을 확인하라")
        return 1
    print("  ✓ 색 %d개·글꼴 그대로" % len(expect_colors))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=TEMPLATE, help="쓸 경로 (기본: 양식 자리)")
    ap.add_argument("--sample", default=TEMPLATE,
                    help="색·글꼴을 물려받을 견본 (기본: 지금 양식)")
    ap.add_argument("--check-only", action="store_true",
                    help="쓰지 않고 서명만 확인")
    a = ap.parse_args()

    from oa_typeset.engines.deck13 import palette as pal
    before, font_before = pal.read(a.sample)
    print("견본: %s" % a.sample)
    print("  navy=%s amber=%s chip=%s navy_55=%s / %s"
          % (before.get("navy"), before.get("amber"), before.get("chip"),
             before.get("navy_55"), font_before))

    data = build(a.out, a.sample)
    tmp = a.out + ".new"
    with open(tmp, "wb") as f:
        f.write(data)
    print("생성: %d slides, %d bytes" % (len(PLAN) + 1, len(data)))

    rc = check(tmp, before, font_before)
    if rc or a.check_only:
        os.remove(tmp)
        if a.check_only and not rc:
            print("(--check-only — 파일을 쓰지 않았다)")
        return rc
    os.replace(tmp, a.out)
    print("→ %s" % a.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
