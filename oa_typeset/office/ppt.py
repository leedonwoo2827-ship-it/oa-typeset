"""PowerPoint 마무리 — **실제 PowerPoint 를 띄워** 줄바꿈을 보고 고친다(Windows + PowerPoint + pywin32).

python-pptx 는 글자가 몇 줄로 접히는지 모른다. 줄바꿈은 PowerPoint 가 글꼴을 실제로 놓아 봐야
정해진다. 그래서 끝줄 정리·넘침 확인은 PowerPoint 자동화(COM)로 한다.

하는 일
  1) 끝줄 정리 — 문단의 마지막 줄이 너무 짧으면(기본 3자 이하, 또는 윗줄의 20% 이하)
     그 문단의 자간을 0.25pt 씩 좁혀(최대 -1.0pt) 윗줄로 끌어올린다. 그래도 안 되면 그대로 두고 보고한다.
  2) 넘침 확인 — 글상자 안 글의 실제 높이가 상자보다 크면 보고한다(자동으로 자르지 않는다).
  3) (선택) 슬라이드 PNG 내보내기 — 눈으로 보는 검증용.

원본을 덮어쓰지 않으려면 out 경로를 따로 준다. PowerPoint 가 이미 열려 있으면 그 창은 건드리지 않는다.
"""
from __future__ import annotations

import os
from typing import Dict, List, Optional

SHORT_CHARS = 3          # 마지막 줄이 이 글자 수 이하면 '짧은 끝줄'
SHORT_RATIO = 0.20       # 또는 윗줄 길이의 이 비율 이하
STEP_PT = 0.25
MIN_SPACING_PT = -1.0
OVERFLOW_TOL_PT = 2.0


def available() -> bool:
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application").Close()
        return True
    except Exception:
        return False


def _lines(tr) -> List[str]:
    """TextRange2 의 줄 목록(PowerPoint 가 실제로 접은 대로)."""
    out = []
    n = tr.GetLines(-1, -1).Count
    for j in range(1, n + 1):
        out.append(tr.GetLines(j, 1).Text.rstrip("\r\n\x0b"))
    return out


def _short_last(lines: List[str]) -> bool:
    if len(lines) < 2:
        return False
    last = lines[-1].strip()
    prev = lines[-2].strip()
    if not last:
        return False
    return len(last) <= SHORT_CHARS or (prev and len(last) <= len(prev) * SHORT_RATIO)


def _fix_para(para) -> Optional[str]:
    """짧은 끝줄이면 자간을 좁혀 본다. 결과: None(문제 없음) | 'fixed' | 'left'."""
    lines = _lines(para)
    if not _short_last(lines):
        return None
    n0 = len(lines)
    base = float(para.Font.Spacing or 0)
    sp = base
    while sp - STEP_PT >= MIN_SPACING_PT - 1e-6:
        sp -= STEP_PT
        para.Font.Spacing = sp
        lines = _lines(para)
        if len(lines) < n0 or not _short_last(lines):
            return "fixed"
    para.Font.Spacing = base           # 못 고쳤으면 원래대로
    return "left"


def _frames(shape):
    """도형 안의 글상자들(표는 칸마다, 그룹은 안쪽까지)."""
    try:
        if shape.HasTable:
            tbl = shape.Table
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    yield tbl.Cell(r, c).Shape, True
            return
    except Exception:
        pass
    try:
        if shape.Type == 6:  # msoGroup
            for k in range(1, shape.GroupItems.Count + 1):
                yield from _frames(shape.GroupItems(k))
            return
    except Exception:
        pass
    try:
        if shape.HasTextFrame and shape.TextFrame2.HasText:
            yield shape, False
    except Exception:
        return


def finish(pptx: str, out: Optional[str] = None, *, tidy: bool = True, png_dir: Optional[str] = None,
           png_width: int = 1600) -> Dict:
    """pptx 를 PowerPoint 로 열어 끝줄 정리·넘침 확인. 결과 보고(dict)를 돌려준다."""
    import pythoncom
    from win32com.client import gencache

    src = os.path.abspath(pptx)
    dst = os.path.abspath(out or pptx)
    rep: Dict = {"file": dst, "fixed": [], "short_left": [], "overflow": [], "png": []}
    pythoncom.CoInitialize()
    # 이른 바인딩(형식 라이브러리) — 늦은 바인딩으로는 TextRange2.Paragraphs(Start, Length) 같은
    # 인자 있는 속성을 부를 수 없다. makepy 가 GetParagraphs/GetLines 로 만들어 준다.
    app = gencache.EnsureDispatch("PowerPoint.Application")
    pres = app.Presentations.Open(src, False, False, False)   # ReadOnly=False, Untitled=False, WithWindow=False
    try:
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for k in range(1, slide.Shapes.Count + 1):
                for shp, in_table in _frames(slide.Shapes(k)):
                    tf = shp.TextFrame2
                    tr = tf.TextRange
                    label = (tr.Text or "").strip().replace("\r", " ")[:24]
                    if tidy:
                        for pi in range(1, tr.GetParagraphs(-1, -1).Count + 1):
                            para = tr.GetParagraphs(pi, 1)
                            if not (para.Text or "").strip():
                                continue
                            try:
                                r = _fix_para(para)
                            except Exception:
                                r = None
                            if r == "fixed":
                                rep["fixed"].append({"slide": si, "text": para.Text.strip()[:30]})
                            elif r == "left":
                                rep["short_left"].append({"slide": si, "text": para.Text.strip()[:30]})
                    if not in_table:
                        try:
                            need = tr.BoundHeight + tf.MarginTop + tf.MarginBottom
                            if need > shp.Height + OVERFLOW_TOL_PT and tf.AutoSize == 0:
                                rep["overflow"].append({"slide": si, "text": label,
                                                       "over_pt": round(need - shp.Height, 1)})
                        except Exception:
                            pass
        if png_dir:
            os.makedirs(png_dir, exist_ok=True)
            h = int(png_width * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth)
            for si in range(1, pres.Slides.Count + 1):
                p = os.path.join(os.path.abspath(png_dir), f"s{si:03d}.png")
                pres.Slides(si).Export(p, "PNG", png_width, h)
                rep["png"].append(p)
        # 고친 게 있을 때만 저장한다. PowerPoint 가 저장하면 임베드 글꼴을 통째로 다시 넣어
        # 파일이 3~4배 커진다(1.9MB → 7.6MB 실측) — 바뀐 것도 없는데 키울 이유가 없다.
        rep["saved"] = bool(rep["fixed"])
        if rep["saved"]:
            if dst == src:
                pres.Save()
            else:
                pres.SaveCopyAs(dst)
        elif dst != src:
            import shutil
            shutil.copyfile(src, dst)
    finally:
        pres.Close()
        if app.Presentations.Count == 0:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return rep


def summary(rep: Dict) -> str:
    s = f"끝줄 정리 {len(rep['fixed'])}곳"
    if rep["short_left"]:
        s += f" · 못 고친 짧은 끝줄 {len(rep['short_left'])}곳"
    if rep["overflow"]:
        s += " · 넘친 상자: " + ", ".join(f"{o['slide']}장" for o in rep["overflow"])
    return s
