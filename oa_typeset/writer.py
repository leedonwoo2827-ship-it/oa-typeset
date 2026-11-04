"""집필 — 자료 + 지시문 + 양식 → LLM 이 HTML 원고를 쓴다(목차 → 장별 HTML → book.html).

    from oa_typeset.writer import write_book
    html_path = write_book(["자료.pdf", "기존교재.hwpx"], "변액보험 판매관리사 요약집, 수험생 대상",
                           form="교재형-46배판", out_dir="work/변액보험", pages=70, provider="codex")

- 장마다 work/<…>/chapters/chNN.html 로 저장한다(체크포인트). 중간에 멈춰도 다시 부르면 끝난 장은 건너뛴다.
- 목차는 outline.json — 사람이 고쳐 두고 다시 부르면 고친 목차로 쓴다.
- HTML 은 html2hwpx 가 옮길 수 있는 태그로만 쓰게 한다(docs/HTML_작성규칙.md 와 같은 규칙).
- LLM: provider="codex"(기본, ChatGPT 로그인) | "claude"(구독 로그인) | "auto"(codex 실패 시 claude).
"""
from __future__ import annotations

import html as _html
import json
import os
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Union

from . import llm
from .forms import Form, get_form

HTML_RULES = """\
[HTML 작성 규칙 — 이 규칙 밖의 태그·속성은 쓰지 않는다]
- 제목: <h1>(장) <h2>(절) <h3>(항) <h4>(세부). 번호는 제목 글에 직접 쓴다(예: <h2>1. 셀 주소</h2>).
- 장(h1)은 새 쪽에서 시작: <h1 style="break-before:page">.
- 본문: <p>. 강조는 <b>만. 작은 회색 주석은 <small>.
- 목록: <ul><li>, 순서 있으면 <ol><li>. 항목 앞에 •·1. 같은 기호를 직접 쓰지 않는다(자동으로 붙는다).
- 표: <table><tr><th>머리</th>…</tr><tr><td>…</td></tr></table>. 칸 병합은 colspan 만(rowspan 금지). 5열 이내.
- 그림이 필요하면 <p><img src="파일명.svg" alt="설명" style="width:120mm"></p> 를 쓰고, 그 SVG 를 따로 요청받았을 때만 만든다.
  지금은 그림 자리에 <p><small>[그림 제안: 무엇을 그릴지 한 줄]</small></p> 만 남긴다.
- <div>, <span>, class, 색·글꼴 지정, <br> 연속, 마크다운 기호(**, ##)는 쓰지 않는다.
- 근거 자료에 없는 수치·사실은 지어내지 말고 [확인 필요] 라고 쓴다.
"""

SYSTEM = ("너는 한국어 출판 원고를 HTML 로 쓰는 편집자다. 요청한 내용만 돌려준다 — 설명·인사·코드펜스 없이.")


# ── 자료 읽기 ─────────────────────────────────────────────────────────
def read_sources(sources: Iterable[str], limit: int = 120_000) -> str:
    from .extractors import extract
    parts, total = [], 0
    for p in sources:
        if not p:
            continue
        txt = extract(p) if os.path.isfile(p) else str(p)
        name = os.path.basename(p) if os.path.isfile(p) else "지시 메모"
        chunk = f"\n\n=== 자료: {name} ===\n{txt.strip()}"
        if total + len(chunk) > limit:
            chunk = chunk[: max(0, limit - total)] + "\n…(자료가 길어 이하 생략)"
        parts.append(chunk)
        total += len(chunk)
        if total >= limit:
            break
    return "".join(parts).strip()


# ── 목차 ──────────────────────────────────────────────────────────────
def make_outline(brief: str, sources_text: str, form: Form, pages: int, provider: Optional[str] = None) -> Dict:
    cpp = int(form.meta.get("chars_per_page") or 1100)
    prompt = f"""다음 자료로 약 {pages}쪽(쪽당 약 {cpp}자) 분량 문서의 목차를 짜라.

[문서 요청]
{brief}

[양식: {form.title}]
{form.guide}

[출력 — JSON 하나만]
{{"title": "문서 제목", "subtitle": "부제(선택)",
  "chapters": [{{"title": "1장. …", "pages": 쪽수, "sections": ["1. …", "2. …"], "notes": "이 장에서 꼭 다룰 것"}}]}}
- 장 쪽수 합이 {pages} 안팎이 되게. 장은 4~10개, 장마다 절 2~6개.
- 자료에 없는 주제를 지어내지 않는다.

[자료]
{sources_text[:60000]}
"""
    return _json(llm.ask(prompt, provider, system=SYSTEM))


# ── 장 집필 ───────────────────────────────────────────────────────────
def write_chapter(ch: Dict, idx: int, outline: Dict, brief: str, sources_text: str, form: Form,
                  provider: Optional[str] = None) -> str:
    cpp = int(form.meta.get("chars_per_page") or 1100)
    target = int(ch.get("pages") or 6) * cpp
    toc = "\n".join(f"{i + 1}. {c.get('title')}" for i, c in enumerate(outline.get("chapters", [])))
    prompt = f"""문서 「{outline.get('title', '')}」 의 {idx}번째 장을 HTML 조각으로 써라.

[이 장]
제목: {ch.get('title')}
절: {', '.join(ch.get('sections') or [])}
꼭 다룰 것: {ch.get('notes') or '-'}
분량: 본문 글자 약 {target}자 (표 포함). 모자라면 설명·예시·표를 더한다.

[전체 목차 — 다른 장과 겹치지 않게]
{toc}

[문서 요청]
{brief}

[양식 지침: {form.title}]
{form.guide}

{HTML_RULES}
[출력] <h1 style="break-before:page">{_html.escape(str(ch.get('title') or ''))}</h1> 로 시작하는 HTML 조각만. <html>/<body> 없이.

[자료]
{sources_text[:80000]}
"""
    return clean_fragment(llm.ask(prompt, provider, system=SYSTEM))


def clean_fragment(text: str) -> str:
    t = text.strip()
    m = re.search(r"```(?:html)?\s*(.*?)```", t, re.S)
    if m:
        t = m.group(1).strip()
    t = re.sub(r"(?is)^.*?<body[^>]*>|</body>.*$", "", t)
    t = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", "", t)
    return t.strip()


# ── 책 한 권 ──────────────────────────────────────────────────────────
def write_book(sources: Iterable[str], brief: str, form: Union[str, Form, None] = None, out_dir: str = "work",
               pages: Optional[int] = None, provider: Optional[str] = None, *, forms_dir: Optional[str] = None,
               only: Optional[List[int]] = None, redo: bool = False, log=print) -> str:
    """자료 → (목차) → 장별 HTML → book.html 경로. only=[3] 이면 3장만, redo=True 면 이미 쓴 장도 다시."""
    f = form if isinstance(form, Form) else get_form(form, forms_dir)
    out = Path(out_dir)
    (out / "chapters").mkdir(parents=True, exist_ok=True)
    pages = int(pages or f.meta.get("pages_target") or 30)
    src_cache = out / "sources.txt"
    if src_cache.is_file() and not redo:
        sources_text = src_cache.read_text(encoding="utf-8")
    else:
        sources_text = read_sources(list(sources))
        src_cache.write_text(sources_text, encoding="utf-8")
    (out / "brief.txt").write_text(brief, encoding="utf-8")

    ol_path = out / "outline.json"
    if ol_path.is_file():
        outline = json.loads(ol_path.read_text(encoding="utf-8"))
        log(f"목차: 저장된 것 사용 ({ol_path})")
    else:
        log(f"목차 짜는 중… ({provider or llm.DEFAULT})")
        outline = make_outline(brief, sources_text, f, pages, provider)
        ol_path.write_text(json.dumps(outline, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"목차: {len(outline.get('chapters', []))}장 → {ol_path}")

    chapters = outline.get("chapters", [])
    for i, ch in enumerate(chapters, 1):
        p = out / "chapters" / f"ch{i:02d}.html"
        if only and i not in only:
            continue
        if p.is_file() and not redo:
            log(f"  {i}장 건너뜀(이미 있음)")
            continue
        log(f"  {i}/{len(chapters)}장 쓰는 중: {ch.get('title')}")
        frag = write_chapter(ch, i, outline, brief, sources_text, f, provider)
        p.write_text(frag, encoding="utf-8")
        log(f"  {i}장 완료 ({len(re.sub(r'<[^>]+>', '', frag))}자)")
    return assemble(out, outline)


def assemble(out_dir: Union[str, Path], outline: Optional[Dict] = None) -> str:
    """chapters/*.html 을 이어 book.html 을 만든다(장 파일을 손으로 고친 뒤에도 이것만 다시 부르면 된다)."""
    out = Path(out_dir)
    if outline is None:
        ol = out / "outline.json"
        outline = json.loads(ol.read_text(encoding="utf-8")) if ol.is_file() else {}
    title = outline.get("title") or out.name
    sub = outline.get("subtitle") or ""
    body = [f"<h1>{_html.escape(title)}</h1>"]
    if sub:
        body.append(f"<p><small>{_html.escape(sub)}</small></p>")
    for p in sorted((out / "chapters").glob("ch*.html")):
        body.append(p.read_text(encoding="utf-8"))
    doc = (f'<!doctype html>\n<html lang="ko"><head><meta charset="utf-8"><title>{_html.escape(title)}</title>\n'
           '<style>body{font-family:"Malgun Gothic",sans-serif;max-width:780px;margin:40px auto;line-height:1.7}'
           'table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:4px 8px}th{background:#f2f2f2}'
           '</style></head>\n<body>\n' + "\n\n".join(body) + "\n</body></html>\n")
    path = out / "book.html"
    path.write_text(doc, encoding="utf-8")
    return str(path)


# ── 발표자료(덱) spec ─────────────────────────────────────────────────
DECK_SPEC = """슬라이드 타입과 필드(필요한 것만):
- "cover"   표지. "title", "chip"(부제)
- "agenda"  목차. "items":[{"part":"Part I","label":"대영역 제목"}] 5개 안팎
- "section" 간지. "title", "chip"(한 줄 요약)
- "bullets" 서술. "bullets":["요점",…] 3~5개, 한 줄 32자 이내. 근거 수치는 "stats":[{"label","value","desc"}] 2~3개
- "cards"   병렬 2~4개. "items":[{"label":"짧은 이름","desc":"한 줄"}]
- "process" 단계 2~5개. "items":[{"label":"단계","desc":"한 줄"}]
- "compare" 두 대상 비교. "items":[{"label":"대상","lines":["…"]}] 2개
- "table"   표. "headers":[…], "rows":[[…]] 7열·8행 이내, "subhead"(표 이름), "source"(출처)
- "stat"    수치 카드 2~3개. "items":[{"label":"70.9","caption":"이름","desc":"근거"}]
- "chart"   "charts":[{"caption","kind":"column|bar|line|pie","categories":[…],"series":[{"name","values":[숫자…]}],"unit"}]
- "closing" 마무리. "bullets":[결론…]
- "thanks"  마지막 인사
공통: "kicker"(상단 라벨=대영역), "title"(결론 문장, 핵심 낱말은 *별표* 로 감싸면 주색), "chip"(한 줄 요약),
"interpretation"(하단 해석 2~3줄)
"""


def deck_from_markdown(md: str, provider: Optional[str] = None, form: Optional[Form] = None,
                       slides: int = 0) -> Dict:
    n = slides or max(8, min(30, len(re.findall(r"^#{2,3} ", md, re.M)) + 4))
    prompt = f"""다음 원고를 발표자료 약 {n}장으로 재구성한 덱 spec JSON 을 만들어라.
원고에 있는 사실·수치만 쓴다. 표는 table, 수치 비교는 chart/stat, 순서는 process 를 적극 쓴다.

{DECK_SPEC}
[출력 — JSON 하나만]
{{"title": "...", "subtitle": "...", "footer": "...", "slides": [{{"type": "cover", ...}}, ...]}}

[원고]
{md[:60000]}
"""
    return _json(llm.ask(prompt, provider, system="너는 발표자료 구성 편집자다. JSON 만 돌려준다."))


def _json(text: str) -> Dict:
    m = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.S)
    body = m.group(1) if m else text[text.find("{"): text.rfind("}") + 1]
    return json.loads(body)
