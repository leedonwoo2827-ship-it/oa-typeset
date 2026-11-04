"""Markdown 출력 — 원문 보존(미리보기/백업용). 항상 동시 생성된다."""
from __future__ import annotations

import json
from typing import Any


def _md_table(headers, rows) -> list:
    """헤더+행 → Markdown 표. 헤더가 없으면 첫 행을 머리로 쓴다."""
    out = []
    headers = [str(h) for h in (headers or [])]
    rows = [[("" if c is None else str(c)) for c in (r or [])] for r in (rows or [])]
    if not headers and rows:
        headers, rows = rows[0], rows[1:]
    if not headers:
        return out
    out.append("| " + " | ".join(headers) + " |")
    out.append("| " + " | ".join("---" for _ in headers) + " |")
    for r in rows:
        cells = (r + [""] * len(headers))[:len(headers)]
        out.append("| " + " | ".join(cells) + " |")
    return out


def _slides_to_markdown(payload: dict) -> str:
    """덱 spec → Markdown.

    ★ 이 결과가 '자동 검수' 단계의 입력이다. 차트 수치나 표를 빠뜨리면 검수가
      근거와 대조할 것을 못 보고 그냥 통과시킨다 — 제목만 뽑던 예전 판이
      실제로 그랬다. 슬라이드에 인쇄되는 값은 전부 여기 나와야 한다.
    """
    lines = [f"# {payload.get('title', '')}".rstrip()]
    if payload.get("subtitle"):
        lines.append(f"*{payload['subtitle']}*")

    for i, s in enumerate(payload.get("slides", []) or [], 1):
        if not isinstance(s, dict):
            continue
        typ = str(s.get("type") or "")
        head = f"\n## {i}. {s.get('title', '')}".rstrip()
        if typ:
            head += f"  `{typ}`"
        lines.append(head)
        if s.get("kicker"):
            lines.append(f"*{s['kicker']}*")
        if s.get("chip"):
            lines.append(f"> {s['chip']}")

        for b in s.get("bullets", []) or []:
            lines.append(f"- {b}")

        for it in s.get("items", []) or []:
            if isinstance(it, dict):
                label = str(it.get("label") or "")
                desc = str(it.get("desc") or "")
                lines.append(f"- **{label}**" + (f" — {desc}" if desc else ""))
                for ln in (it.get("lines") or []):
                    lines.append(f"  - {ln}")
            else:
                lines.append(f"- {it}")

        if s.get("headers") or s.get("rows"):
            lines += _md_table(s.get("headers"), s.get("rows"))

        # 차트 — 카테고리를 행, 계열을 열로 펴서 값이 그대로 보이게 한다.
        for c in (s.get("charts") or []):
            if not isinstance(c, dict):
                continue
            if c.get("caption"):
                lines.append(f"\n**{c['caption']}**")
            cats = [str(x) for x in (c.get("categories") or [])]
            series = c.get("series") or []
            names, cols = [], []
            for se in series:
                if isinstance(se, dict):
                    names.append(str(se.get("name") or ""))
                    cols.append(list(se.get("values") or []))
            if cats and cols:
                rows = [[cats[r]] + [("" if r >= len(v) or v[r] is None else str(v[r]))
                                     for v in cols] for r in range(len(cats))]
                lines += _md_table(["항목"] + names, rows)

        if s.get("question"):
            lines.append(f"- 문항: {s['question']}")
            for ch in (s.get("choices") or []):
                lines.append(f"  - {ch}")
            if s.get("answer"):
                lines.append(f"  - 정답: {s['answer']}")

        if s.get("key"):
            lines.append(f"\n**핵심** {s['key']}")

    return "\n".join(lines)


def to_markdown(payload: Any) -> str:
    """어떤 payload 든 사람이 읽을 Markdown 으로 변환(미리보기 공용)."""
    if isinstance(payload, str):
        return payload
    if isinstance(payload, dict) and "slides" in payload:
        return _slides_to_markdown(payload)
    if isinstance(payload, dict) and "rows" in payload:
        cols = payload.get("columns", [])
        lines = []
        if payload.get("title"):
            lines.append(f"# {payload['title']}")
        if cols:
            lines.append("| " + " | ".join(map(str, cols)) + " |")
            lines.append("| " + " | ".join("---" for _ in cols) + " |")
        for row in payload.get("rows", []):
            lines.append("| " + " | ".join("" if c is None else str(c) for c in row) + " |")
        return "\n".join(lines)
    return "```json\n" + json.dumps(payload, ensure_ascii=False, indent=2) + "\n```"


def render(payload: Any, out_path: str) -> str:
    text = to_markdown(payload)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    return out_path
