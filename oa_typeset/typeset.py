"""조판 — 원고(HTML 또는 Markdown) + 양식 → hwpx / docx / pptx.

    from oa_typeset import typeset
    r = typeset("book.html", form="교재형-46배판", formats=["hwpx", "docx", "pptx"], out_dir="out")
    r["hwpx"]["path"], r["pptx"]["notes"] ...

엔진
  hwpx : engines/html2hwpx  (HTML 을 직접 읽어 진짜 한글 표·그림·쪽 나눔까지. 양식의 글꼴·스타일·바탕쪽 상속)
         .md 원고는 engines/hwpx_md (aim-local 의 양식 ZIP 복제 방식)
  docx : HTML → Markdown 블록 → engines/docx_gen (양식의 명명 스타일 상속, 표·그림)
  pptx : HTML → 덱 spec → engines/deck13 (13종 레이아웃·배색은 form.pptx 시안에서)
         deck_spec 를 주면 그대로, provider 를 주면 LLM 이 spec 을 짠다, 둘 다 없으면 목차 구조로 자동 생성.
마무리(finish=True, Windows+Office): pptx 끝줄 정리·넘침 확인 / hwpx 한글로 다시 앉히기·PDF·쪽수.
"""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Union

from .forms import Form, get_form

FORMATS = ("hwpx", "docx", "pptx", "md")


def typeset(src: str, form: Union[str, Form, None] = None, formats: Iterable[str] = ("hwpx",),
            out_dir: Optional[str] = None, *, forms_dir: Optional[str] = None, finish: bool = False,
            deck_spec: Optional[Union[str, Dict]] = None, provider: Optional[str] = None,
            head: Optional[str] = None, stem: Optional[str] = None) -> Dict[str, Any]:
    """원고 파일 하나를 양식대로 조판한다. 결과: {형식: {"path", "notes":[...], ...}}"""
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        raise FileNotFoundError(src)
    f = form if isinstance(form, Form) else get_form(form, forms_dir)
    out_dir = os.path.abspath(out_dir or os.path.dirname(src))
    os.makedirs(out_dir, exist_ok=True)
    stem = stem or Path(src).stem
    is_md = src.lower().endswith((".md", ".markdown", ".txt"))
    md_cache: Dict[str, str] = {}

    def md() -> str:
        if "md" not in md_cache:
            if is_md:
                md_cache["md"] = Path(src).read_text(encoding="utf-8")
            else:
                from .htmlblocks import file_to_markdown
                md_cache["md"] = file_to_markdown(src)
        return md_cache["md"]

    result: Dict[str, Any] = {"form": f.name, "src": src}
    for fmt in [x.strip().lower() for x in formats if x and x.strip()]:
        if fmt not in FORMATS:
            raise ValueError(f"모르는 형식: {fmt} (가능: {', '.join(FORMATS)})")
        out = os.path.join(out_dir, f"{stem}.{fmt}")
        t0 = time.time()
        r: Dict[str, Any] = {"path": out, "notes": []}
        if fmt in ("hwpx", "docx", "pptx") and not f.file(fmt):
            # 양식에 이 형식 파일이 없다 — 엔진 기본 모양으로 만든다는 것을 알린다
            r["notes"].append(f"양식 '{f.name}' 에 form.{fmt} 가 없어 기본 모양으로 만들었습니다")
        if fmt == "hwpx":
            h = _hwpx(src, out, f, is_md, md, head)
            r["notes"] += h.pop("notes", [])
            r.update(h)
            if finish:
                r["notes"] += _finish_hwp(out, r)
        elif fmt == "docx":
            from .engines import docx_gen
            docx_gen.render(_svg_to_png_md(md(), out_dir), out, template=f.file("docx"))
            import shutil
            shutil.rmtree(os.path.join(out_dir, "_svgpng"), ignore_errors=True)
        elif fmt == "pptx":
            spec = _deck_spec(md(), deck_spec, provider, f)
            from .engines import deck13
            report: List[Dict] = []
            deck13.render(spec, out, template=f.file("pptx"), report=report,
                          logo_right=f.logo("right"), logo_left=f.logo("left"))
            r["slides"] = len(spec.get("slides") or [])
            r["notes"] += deck13.overflow_notes(report)
            if finish:
                r["notes"] += _finish_ppt(out, r)
        elif fmt == "md":
            Path(out).write_text(md(), encoding="utf-8")
        r["seconds"] = round(time.time() - t0, 1)
        result[fmt] = r
    return result


def _svg_to_png_md(md: str, out_dir: str) -> str:
    """Markdown 의 SVG 그림 → PNG (Word 는 SVG 를 못 싣는다). 옆에 같은 이름 .png 가 있으면 그것,
    없으면 html2hwpx 와 같은 방법(Edge/Chrome 헤드리스)으로 굽는다. 못 구우면 그대로 둔다(생성기가 건너뜀)."""
    from .engines import html2hwpx
    tool, tmp = None, None

    def rep(m):
        nonlocal tool, tmp
        path = m.group(2).strip()
        if not path.lower().endswith(".svg") or not os.path.isfile(path):
            return m.group(0)
        side = os.path.splitext(path)[0] + ".png"
        if os.path.isfile(side):
            return f"![{m.group(1)}]({side})"
        if tmp is None:
            tool = html2hwpx._browser()
            tmp = os.path.join(out_dir, "_svgpng")
            os.makedirs(tmp, exist_ok=True)
        png = html2hwpx._svg_to_png(path, tmp, tool)
        return f"![{m.group(1)}]({png})" if png else m.group(0)

    return re.sub(r"^!\[([^\]]*)\]\((.+)\)\s*$", rep, md, flags=re.M)


# ── hwpx ──────────────────────────────────────────────────────────────
def _hwpx(src: str, out: str, f: Form, is_md: bool, md, head: Optional[str]) -> Dict:
    tpl = f.file("hwpx")
    if is_md:
        if not tpl:
            raise FileNotFoundError(f"양식 '{f.name}' 에 form.hwpx 가 없습니다 (.md → hwpx 는 양식이 필요)")
        from .engines import hwpx_md
        hwpx_md._render_zip(md(), out, tpl)
        return {"engine": "hwpx_md"}
    from .engines import html2hwpx
    page = f.page if tpl else (f.page if f.page != "keep" else (188, 257))
    keep = tpl and page == "keep"
    r = html2hwpx.convert(src, out, page=page, margin_mm=None if keep else float(f.meta.get("margin_mm", 15)),
                          template=tpl, font=(f.meta.get("font") or None) if tpl else (f.meta.get("font") or "맑은 고딕"),
                          styles=f.styles, head=head)
    notes = []
    if not r.get("svg_tool"):
        notes.append("Edge/Chrome 을 못 찾아 SVG 그림은 옆의 .png 가 있을 때만 들어갔습니다")
    return {"engine": "html2hwpx", "paragraphs": r["paragraphs"], "tables": r["tables"], "images": r["images"],
            "notes": notes}


def _finish_hwp(path: str, r: Dict) -> List[str]:
    from .office import hwp
    if not hwp.available():
        return ["한컴 한글이 없어 마무리(다시 앉히기·PDF)를 건너뜀"]
    try:
        rep = hwp.finish(path)
        r["pages"], r["pdf"] = rep.get("pages"), rep.get("pdf")
        return [f"한글 마무리: {rep.get('pages')}쪽" + (" · PDF" if rep.get("pdf") else "")]
    except Exception as e:  # noqa: BLE001
        return [f"한글 마무리 실패: {e}"]


def _finish_ppt(path: str, r: Dict) -> List[str]:
    from .office import ppt
    if not ppt.available():
        return ["PowerPoint 가 없어 끝줄 정리를 건너뜀"]
    try:
        rep = ppt.finish(path)
        r["finish"] = rep
        return ["PowerPoint 마무리: " + ppt.summary(rep)]
    except Exception as e:  # noqa: BLE001
        return [f"PowerPoint 마무리 실패: {e}"]


# ── pptx: 덱 spec ─────────────────────────────────────────────────────
def _deck_spec(md: str, deck_spec, provider: Optional[str], f: Form) -> Dict:
    if deck_spec:
        if isinstance(deck_spec, dict):
            return deck_spec
        if os.path.isfile(str(deck_spec)):
            return json.loads(Path(deck_spec).read_text(encoding="utf-8"))
        return json.loads(deck_spec)
    if provider:
        from .writer import deck_from_markdown
        try:
            return deck_from_markdown(md, provider=provider, form=f)
        except Exception as e:  # noqa: BLE001
            print(f"[pptx] LLM 덱 구성 실패 → 목차 구조로 자동 생성: {e}", flush=True, file=__import__("sys").stderr)
    return md_to_deck(md)


MAX_BULLETS = 5
MAX_ROWS = 8


def md_to_deck(md: str) -> Dict:
    """LLM 없이 Markdown 구조만으로 덱 spec 을 만든다.

    첫 # → 표지(바로 뒤 문단은 부제). 나머지 제목 중 가장 높은 단계 → 간지(section),
    그 아래 단계 → 슬라이드 제목(상단 라벨 = 속한 간지). 문단은 문장 단위 불릿으로 쪼개 한 장에 5개씩,
    표 → table(8행씩 나눔). 마지막에 thanks.
    """
    from .engines import mdblocks
    blocks = mdblocks.parse(md)
    heads = [b for b in blocks if b["type"] == "heading"]
    title = heads[0]["text"] if heads and heads[0].get("level") == 1 else ""
    rest = heads[1:] if title else heads
    top = min((h.get("level", 2) for h in rest), default=2)
    subtitle = ""
    slides: List[Dict] = []
    part, sub = "", ""
    cur: Optional[Dict] = None
    seen_title = not title

    def new(t: str) -> Dict:
        d = {"type": "bullets", "title": t, "kicker": part, "bullets": []}
        slides.append(d)
        return d

    for b in blocks:
        t, txt = b["type"], b.get("text", "")
        if t == "heading":
            lv = b.get("level", 1)
            if not seen_title:
                seen_title = True
                continue
            if lv <= top:
                part, sub = txt, ""
                slides.append({"type": "section", "title": txt})
                cur = None
            else:
                sub = txt
                cur = new(txt)
        elif t in ("para", "bullet", "quote"):
            if title and not slides and not subtitle and t == "para":
                subtitle = txt
                continue
            for piece in _bullets(txt):
                if cur is None or len(cur["bullets"]) >= MAX_BULLETS:
                    base = (cur or {}).get("title") or sub or part or title or "내용"
                    cur = new(base if cur is None else base.removesuffix(" (계속)") + " (계속)")
                cur["bullets"].append(piece)
        elif t == "table" and (b.get("rows") or b.get("header")):
            head = b.get("header") or []
            rows = b.get("rows") or []
            if not head and rows:
                head, rows = rows[0], rows[1:]
            base = sub or part or "표"
            for i in range(0, max(1, len(rows)), MAX_ROWS):
                slides.append({"type": "table", "title": base if i == 0 else base + " (계속)", "kicker": part,
                               "headers": head[:7], "rows": [r[:7] for r in rows[i:i + MAX_ROWS]]})
            cur = None
    slides = [s for s in slides if not (s["type"] == "bullets" and not s["bullets"])]
    # 내용 없는 간지(바로 뒤가 또 간지이거나 끝)는 뺀다
    slides = [s for i, s in enumerate(slides)
              if not (s["type"] == "section" and (i + 1 >= len(slides) or slides[i + 1]["type"] == "section"))]
    slides.insert(0, {"type": "cover", "title": title or "발표자료", "chip": subtitle})
    slides.append({"type": "thanks", "title": ""})
    return {"title": title or "발표자료", "subtitle": subtitle, "footer": title, "slides": slides}


def _bullets(text: str, n: int = 70) -> List[str]:
    """문단 → 문장 단위 불릿(긴 문장만 줄인다). 슬라이드 글은 문장 하나가 한 항목이다."""
    text = re.sub(r"\s+", " ", text).strip()
    parts = [p.strip() for p in re.split(r"(?<=[.!?。])\s+", text) if p.strip()]
    return [_shorten(p, n) for p in parts] or [text]


def _shorten(s: str, n: int = 60) -> str:
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) <= n:
        return s
    cut = s.rfind(" ", 0, n)
    return s[: cut if cut > n * 0.6 else n].rstrip(",·;: ") + "…"
