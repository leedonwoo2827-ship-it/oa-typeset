"""아주 가벼운 Markdown → 블록 파서.

생성기(docx/hwpx)가 '제목/본문/목록/강조/표' 스타일로 매핑하기 위한 최소 구조만 추출한다.
지원: # ~ ###### 헤딩, 불릿(-,*,+ / 1.), 인용(>), 코드펜스(```), 표(| a | b |), 일반 단락, **굵게**.
HTML <br> 은 셀/문장 안에서 줄 구분으로 쓰이므로 ' / ' 로 정규화한다.
"""
from __future__ import annotations

import re
from typing import Dict, List, Tuple

_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
_BULLET = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$")
_QUOTE = re.compile(r"^>\s?(.*)$")
_BOLD = re.compile(r"\*\*(.+?)\*\*|__(.+?)__")
_TABLE_ROW = re.compile(r"^\s*\|.*\|\s*$")
_BR = re.compile(r"<\s*br\s*/?\s*>", re.I)
_IMAGE = re.compile(r"^!\[([^\]]*)\]\((.+)\)\s*$")


def parse(md: str) -> List[Dict]:
    """블록 리스트 반환. 각 블록: {type, ...}.

    type: heading(level,text) | para(text) | bullet(level,text) | quote(text)
          | code(text) | table(header[], rows[[]]) | image(path, text=alt)
    """
    blocks: List[Dict] = []
    lines = (md or "").replace("\r\n", "\n").split("\n")
    i = 0
    para: List[str] = []

    def flush_para():
        if para:
            text = " ".join(s.strip() for s in para).strip()
            if text:
                blocks.append(_mk("para", text))
            para.clear()

    while i < len(lines):
        line = lines[i]
        # 코드펜스
        if line.strip().startswith("```"):
            flush_para()
            i += 1
            buf = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            blocks.append({"type": "code", "text": "\n".join(buf), "runs": [("\n".join(buf), False)]})
            continue
        # 표: | ... | 로 시작/끝나는 연속 줄
        if _TABLE_ROW.match(line):
            flush_para()
            tbl_lines = []
            while i < len(lines) and _TABLE_ROW.match(lines[i]):
                tbl_lines.append(lines[i])
                i += 1
            blocks.append(_table(tbl_lines))
            continue
        if not line.strip():
            flush_para()
            i += 1
            continue
        m = _IMAGE.match(line.strip())
        if m:
            flush_para()
            blocks.append({"type": "image", "level": 0, "text": m.group(1), "path": m.group(2).strip(),
                           "runs": [(m.group(1), False)]})
            i += 1
            continue
        m = _HEADING.match(line)
        if m:
            flush_para()
            blocks.append(_mk("heading", m.group(2).strip(), level=len(m.group(1))))
            i += 1
            continue
        m = _BULLET.match(line)
        if m:
            flush_para()
            level = 1 + len(m.group(1)) // 2
            blocks.append(_mk("bullet", m.group(3).strip(), level=level))
            i += 1
            continue
        m = _QUOTE.match(line)
        if m:
            flush_para()
            blocks.append(_mk("quote", m.group(1).strip()))
            i += 1
            continue
        para.append(line)
        i += 1
    flush_para()
    return blocks


def _split_row(line: str) -> List[str]:
    return [_cell(c) for c in line.strip().strip("|").split("|")]


def _is_separator(cells: List[str]) -> bool:
    joined = "".join(cells).replace(" ", "")
    return joined != "" and set(joined) <= set("-:")


def _table(tbl_lines: List[str]) -> Dict:
    rows = [_split_row(l) for l in tbl_lines]
    rows = [r for r in rows if r]  # 빈 줄 제거
    header: List[str] = []
    body: List[List[str]] = []
    if rows:
        # 두 번째 줄이 구분선이면 첫 줄을 헤더로
        if len(rows) >= 2 and _is_separator(rows[1]):
            header = rows[0]
            body = [r for r in rows[2:] if not _is_separator(r)]
        else:
            body = [r for r in rows if not _is_separator(r)]
    ncol = max([len(header)] + [len(r) for r in body] or [0]) if (header or body) else 0
    header = (header + [""] * ncol)[:ncol] if header else []
    body = [(r + [""] * ncol)[:ncol] for r in body]
    return {"type": "table", "header": header, "rows": body, "ncol": ncol}


def _mk(btype: str, text: str, level: int = 0) -> Dict:
    return {"type": btype, "level": level, "text": _strip_inline(text), "runs": _runs(text)}


def _runs(text: str) -> List[Tuple[str, bool]]:
    """**굵게** 를 (조각, bold) 런 리스트로 분해."""
    text = _BR.sub(" / ", text)
    runs: List[Tuple[str, bool]] = []
    pos = 0
    for m in _BOLD.finditer(text):
        if m.start() > pos:
            runs.append((_strip_inline(text[pos:m.start()]), False))
        runs.append((m.group(1) or m.group(2) or "", True))
        pos = m.end()
    if pos < len(text):
        runs.append((_strip_inline(text[pos:]), False))
    return [r for r in runs if r[0]] or [(_strip_inline(text), False)]


def _cell(text: str) -> str:
    """표 셀: <br> → ' / ', 인라인 마크업 제거."""
    return _strip_inline(_BR.sub(" / ", text))


def _strip_inline(text: str) -> str:
    text = _BR.sub(" / ", text)
    text = _BOLD.sub(lambda m: m.group(1) or m.group(2) or "", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)              # 인라인 코드
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)  # 링크 → 텍스트만
    return text.strip()
