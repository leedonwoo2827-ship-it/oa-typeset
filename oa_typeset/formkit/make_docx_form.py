"""Word 양식(form.docx) 만들기 — 판형·글꼴·제목 스타일만 정한 빈 문서.

    python -m oa_typeset.formkit.make_docx_form forms/<이름>/form.docx --page 188x257 --margin 15 --font "맑은 고딕" --size 10

docx 생성기는 스타일 이름으로 찾는다: Title / Heading 1~3 / Normal / List Bullet / Quote
(한국어 Word 에서는 '제목', '제목 1~3', '표준', '글머리 기호 목록', '인용' 으로 보인다).
회사 양식이 따로 있으면 그 .docx 를 form.docx 로 두면 된다 — 이 스크립트는 기본값을 만드는 것뿐.
"""
from __future__ import annotations

import argparse


def make(out: str, page=(210, 297), margin=20.0, font="맑은 고딕", size=10.5, color="1F3A5F") -> str:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(page[0]), Mm(page[1])
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Mm(margin))

    def face(style, pt, bold=False, rgb=None, before=0, after=4):
        f = style.font
        f.name = font
        f.size = Pt(pt)
        f.bold = bold
        if rgb:
            f.color.rgb = RGBColor.from_string(rgb)
        rpr = style.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = rpr.makeelement(qn("w:rFonts"), {})
            rpr.append(fonts)
        for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            fonts.set(qn(k), font)
        pf = style.paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.line_spacing = 1.5 if not bold else 1.2
        pf.widow_control = True

    st = doc.styles
    face(st["Normal"], size)
    face(st["Title"], size + 12, True, color, 0, 12)
    face(st["Heading 1"], size + 6, True, color, 18, 8)
    face(st["Heading 2"], size + 3, True, color, 12, 6)
    face(st["Heading 3"], size + 1, True, "333333", 8, 4)
    face(st["List Bullet"], size)
    face(st["Quote"], size, False, "555555")
    for h in ("Heading 1", "Heading 2", "Heading 3"):
        st[h].paragraph_format.keep_with_next = True
    doc.save(out)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--page", default="210x297")
    ap.add_argument("--margin", type=float, default=20.0)
    ap.add_argument("--font", default="맑은 고딕")
    ap.add_argument("--size", type=float, default=10.5)
    ap.add_argument("--color", default="1F3A5F")
    a = ap.parse_args(argv)
    w, h = (float(x) for x in a.page.lower().split("x"))
    print("→", make(a.out, (w, h), a.margin, a.font, a.size, a.color))


if __name__ == "__main__":
    main()
