"""디자인 덱 PPTX — 13종 레이아웃 렌더 엔진의 AIM 어댑터.

렌더 엔진(builder/fonts/palette)은 aim-local 에서 가져왔다. 이 파일이 oa_typeset 규약에 맞춘다. 형제 생성기(hwpx_gen·docx_gen)와 같은 모양이다:

    render(payload, out_path, template=None) -> out_path

payload 는 pipeline 이 LLM 에게서 받은 덱 spec(JSON):
  {"title":.., "subtitle":.., "footer":.., "palette":"cyan",
   "slides":[{"type":"cover|section|agenda|bullets|cards|process|compare|table|
                      stat|closing|photo|quiz|objectives",
              "title":.., "kicker":.., "chip":.., "bullets":[..],
              "items":[{"label":..,"desc":..}], "headers":[..], "rows":[[..]],
              "key":".."}]}

AIM 고유 규약 — **[확인 필요]**
  근거에 없는 값은 지어내지 않고 '[확인 필요]' 로 남긴다는 것이 이 도구의 원칙이다.
  그 글자가 든 런은 렌더가 끝난 뒤 앰버(amber_dk)로 물들여 작성자가 채울 자리를
  한눈에 보게 한다. 앰버를 **텍스트에 쓸 때는 반드시 amber_dk** — 순앰버는 흰 배경
  대비가 2.0:1 이라 본문 색으로 쓰지 않는다(docs 디자인 토큰 규칙).
"""
from __future__ import annotations

import io
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple

from . import builder

# 이 엔진이 실제로 그릴 수 있는 타입. 모르는 타입은 bullets 로 떨어진다.
KNOWN_TYPES = tuple(builder._RENDER.keys())

# '확인 필요' 를 알아보는 규칙 — 프롬프트가 시키는 표기를 모두 받아 준다.
_PENDING = re.compile(r"확인\s*필요|\[미확인\]|\bTBD\b|\bTBC\b", re.IGNORECASE)

_HANGUL = re.compile(r"[가-힣]")


def _lang_of(data: Dict[str, Any]) -> str:
    """덱 언어 판정 — 슬라이드에 인쇄되는 고정 문구(정답·요약)를 맞추려는 것.

    레시피가 lang 을 주면 그대로 쓰고, 없으면 제목·불릿에 한글이 있는지로 본다.
    083·079 처럼 '전체 영문' 을 지시하는 덱에 '정답'·'요약' 이 섞이면 안 된다.
    """
    declared = str(data.get("lang") or "").strip().lower()
    if declared:
        return "en" if declared.startswith("en") else "ko"
    sample = []
    for s in (data.get("slides") or [])[:12]:
        if isinstance(s, dict):
            sample.append(str(s.get("title") or ""))
            sample += [str(b) for b in (s.get("bullets") or [])[:3]]
    return "ko" if _HANGUL.search(" ".join(sample)) else "en"


# ── payload 정규화 ──────────────────────────────────────────────────────
def _coerce(payload: Any) -> Dict[str, Any]:
    """문자열/이중 인코딩된 JSON 도 받아 준다(LLM 출력이 늘 깨끗하지 않다)."""
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except Exception:
            return {"slides": []}
    if isinstance(payload, list):          # 슬라이드 배열만 온 경우
        return {"slides": payload}
    if not isinstance(payload, dict):
        return {"slides": []}
    return payload


def _norm_items(items: Any) -> List[Dict[str, Any]]:
    """items 를 [{label, desc, lines}] 로 통일. 문자열 배열도 허용."""
    out: List[Dict[str, Any]] = []
    for it in (items or []):
        if isinstance(it, str):
            out.append({"label": it, "desc": ""})
        elif isinstance(it, dict):
            d = dict(it)
            d["label"] = str(d.get("label") or d.get("title") or d.get("no") or "")
            if d.get("desc") is not None:
                d["desc"] = str(d["desc"])
            lines = d.get("lines")
            if isinstance(lines, str):
                d["lines"] = [lines]
            elif isinstance(lines, list):
                d["lines"] = [str(x) for x in lines]
            out.append(d)
    return out


def _norm_rows(rows: Any) -> List[List[str]]:
    return [[("" if c is None else str(c)) for c in (r or [])]
            for r in (rows or [])]


def _norm_table(b: Any) -> Optional[Dict[str, Any]]:
    """표 블록 하나 — {subhead, source, headers, rows}."""
    if not isinstance(b, dict) or not b.get("rows"):
        return None
    d = dict(b)
    for k in ("subhead", "source", "unit"):
        if d.get(k) is not None:
            d[k] = str(d[k])
    d["headers"] = [str(h) for h in (d.get("headers") or [])]
    d["rows"] = _norm_rows(d.get("rows"))
    return d


def _norm_lines(v: Any) -> List[str]:
    if isinstance(v, str):
        v = [v]
    return [str(x) for x in (v or []) if str(x).strip()]


def _norm_slide(s: Any) -> Optional[Dict[str, Any]]:
    if not isinstance(s, dict):
        return None
    t = str(s.get("type") or "bullets").strip().lower()
    # 흔한 별칭 — LLM 이 스키마를 살짝 벗어나도 살려서 쓴다.
    t = {"divider": "section", "toc": "agenda", "content": "bullets",
         "kpi": "stat", "kpi_cards": "stat", "timeline": "process",
         "summary": "closing", "conclusion": "closing"}.get(t, t)
    if t not in KNOWN_TYPES:
        t = "bullets"
    d = dict(s)
    d["type"] = t
    d["title"] = str(d.get("title") or "")
    for k in ("kicker", "chip", "key", "subtitle"):
        if d.get(k) is not None:
            d[k] = str(d[k])
    b = d.get("bullets")
    if isinstance(b, str):
        d["bullets"] = [b]
    elif isinstance(b, list):
        d["bullets"] = [str(x) for x in b if str(x).strip()]
    if d.get("items") is not None:
        d["items"] = _norm_items(d["items"])
    if d.get("rows") is not None:
        d["rows"] = [[("" if c is None else str(c)) for c in (r or [])]
                     for r in (d["rows"] or [])]
    if d.get("headers") is not None:
        d["headers"] = [str(h) for h in (d["headers"] or [])]

    # ↓ 1차 정답지에서 확인한 컴포넌트들. 모델 출력이 늘 깨끗하지 않으니
    #   여기서 모양을 맞춘다 — 렌더러가 방어 코드를 겹쳐 들지 않게 한다.
    for k in ("note", "subhead", "source", "layout", "presenter",
              "affiliation"):
        if d.get(k) is not None:
            d[k] = str(d[k])
    if d.get("interpretation") is not None:
        d["interpretation"] = _norm_lines(d["interpretation"])
    if d.get("captions") is not None:
        d["captions"] = [str(x) for x in (d["captions"] or [])]
    if d.get("stats") is not None:
        d["stats"] = _norm_items(d["stats"])
    if d.get("tables") is not None:
        d["tables"] = [t for t in (_norm_table(b) for b in (d["tables"] or []))
                       if t]
    if isinstance(d.get("panel"), dict):
        d["panel"] = {"label": str(d["panel"].get("label") or ""),
                      "lines": _norm_lines(d["panel"].get("lines"))}
    return d


def _plan_of(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw = data.get("slides")
    if not isinstance(raw, list):
        return []
    plan = [x for x in (_norm_slide(s) for s in raw) if x]
    # 표지가 없으면 메타로 한 장 세운다 — 덱은 표지로 시작해야 한다.
    if plan and plan[0].get("type") != "cover" and (data.get("title") or "").strip():
        plan.insert(0, {"type": "cover", "title": str(data["title"]),
                        "chip": str(data.get("subtitle") or "")})
    return plan


# ── [확인 필요] 앰버 표시 ───────────────────────────────────────────────
def _mark_pending(blob: bytes) -> bytes:
    """'확인 필요' 가 든 런을 앰버(amber_dk)로. 실패해도 덱은 그대로 낸다."""
    try:
        from pptx import Presentation
        from pptx.dml.color import RGBColor

        # ★ 배색에서 뽑지 않고 **고정 경고색**을 쓴다. 이건 브랜드색이 아니라
        #   상태색이다 — 견본이 단일 색조(녹색)면 배색의 강조색도 초록이라
        #   본문과 구분이 안 된다. 디자인 문서도 상태색은 브랜드 색조와
        #   분리하라고 못박고 있다.
        amber = RGBColor(0xB5, 0x7F, 0x00)
        prs = Presentation(io.BytesIO(blob))
        hit = 0

        def paint(tf):
            nonlocal hit
            for para in tf.paragraphs:
                for run in para.runs:
                    if _PENDING.search(run.text or ""):
                        run.font.color.rgb = amber
                        hit += 1

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    paint(shape.text_frame)
                if getattr(shape, "has_table", False) and shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            paint(cell.text_frame)
        if not hit:
            return blob
        out = io.BytesIO()
        prs.save(out)
        return out.getvalue()
    except Exception:
        return blob


# ── 메인 ────────────────────────────────────────────────────────────────
def _suggest(details: List[str]) -> str:
    """넘친 정도를 보고 '나눠줘' 냐 '줄여줘' 냐를 고른다.

    조금 넘쳤으면 문장을 줄이는 편이 낫고, 많이 넘쳤거나 **항목·표 행이 화면 밖으로
    나갔으면** 줄여서 해결될 일이 아니라 장을 쪼개야 한다.
    """
    total, hard = 0, False
    for d in details:
        m = re.search(r"(\d+)", d or "")
        total += int(m.group(1)) if m else 1
        if ("항목" in d) or ("행" in d):
            hard = True
    return "나눠줘" if (hard or total >= 3) else "줄여줘"


def overflow_notes(report: List[Dict[str, Any]]) -> List[str]:
    """넘침 보고 → 화면에 띄울 문장. 사람이 그대로 복사해 채팅에 쓸 수 있게 쓴다.

    슬라이드는 '줄일지 나눌지'가 장마다 다르다. 그래서 자동으로 손대지 않고
    **어느 장이 넘쳤는지만** 알려 주고 판단은 사람에게 넘긴다.
    """
    if not report:
        return []
    per: Dict[int, List[str]] = {}
    for r in report:
        no = r.get("slide")
        if not no:
            continue
        per.setdefault(no, []).append(str(r.get("detail") or ""))
    if not per:
        return []
    # **박스는 하나다.** 통째로 긁어 채팅에 붙이면 바로 먹히는 명령 한 줄.
    #
    # 처음엔 줄마다 박스를 만들어 다섯 칸이 쌓였고, 다음엔 안내 칸과 명령 칸 둘로
    # 줄였다. 그래도 "어느 칸을 복사하나" 를 고르게 했다. 복사할 것이 하나면 고를
    # 일이 없다. 그래서 다른 글자는 한 자도 섞지 않는다 — 안내 문구와 장별 사유는
    # 화면이 박스 **밖에** 적는다(긁어도 안 딸려오는 자리).
    cmds = ["%d장 %s" % (no, _suggest(list(dict.fromkeys(per[no]))))
            for no in sorted(per)]
    return [", ".join(cmds)]


def render(payload: Any, out_path: str, template: Optional[str] = None,
           report: Optional[List[Dict[str, Any]]] = None,
           logo_right: Optional[str] = None, logo_left: Optional[str] = None) -> str:
    """logo_right/logo_left — 양식 폴더의 logo_right.png / logo_left.png (없으면 로고 없이)."""
    data = _coerce(payload)
    plan = _plan_of(data)
    if not plan:
        plan = [{"type": "cover",
                 "title": str(data.get("title") or "내용 없음"),
                 "chip": "생성된 슬라이드가 없습니다 — 입력 자료를 확인해 주세요."}]

    footer = str(data.get("footer") or data.get("subtitle") or "")
    title = str(data.get("title") or (plan[0].get("title") if plan else "") or "발표자료")

    # 슬라이드에 인쇄되는 고정 문구를 덱 언어에 맞춘다(영문 덱에 '정답'이 섞이지 않게).
    builder.set_labels(_lang_of(data))

    # 배색·글꼴은 견본(template)에서 뽑는다 — 고르는 값이 아니다.
    # 로고는 builder 규약상 logo_path=우상단 / logo2_path=좌상단.
    blob = builder.build_deck(
        plan,
        template_path=template,
        deck_title=title,
        footer=footer,
        logo_path=logo_right if (logo_right and os.path.isfile(logo_right)) else None,
        logo2_path=logo_left if (logo_left and os.path.isfile(logo_left)) else None,
        report=report,
    )
    if blob is None:
        raise ModuleNotFoundError("python-pptx 가 설치되어 있지 않습니다")

    blob = _mark_pending(blob)
    with open(out_path, "wb") as f:
        f.write(blob)
    return out_path


def _asset(name: str) -> Optional[str]:
    """assets/ 에 있으면 경로, 없으면 None (로고는 선택 사항)."""
    p = os.path.join(str(builder.ASSETS_DIR), name)
    return p if os.path.isfile(p) else None
