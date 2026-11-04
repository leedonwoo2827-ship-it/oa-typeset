"""HTML → Markdown(블록) — docx·pptx 생성기에 넣기 위한 변환.

hwpx 는 html2hwpx 가 HTML 을 직접 읽는다. docx·pptx 생성기(aim-local 출신)는
Markdown 블록(mdblocks)을 받으므로 여기서 HTML 을 그 모양으로 옮긴다.

옮기는 것: h1~h6, p/div, ul/ol/li(중첩), blockquote, pre, table(th·td, colspan 은 빈칸으로 채움),
b/strong(굵게), br, img(→ ![alt](절대경로) — docx 에 그림으로 들어간다).
버리는 것: script/style/head/svg/noscript, 인라인 꾸밈(span 색 등).
표준 라이브러리만 쓴다.
"""
from __future__ import annotations

import html as _html
import os
import re
from html.parser import HTMLParser
from typing import List, Optional

SKIP = {"script", "style", "head", "title", "svg", "noscript", "template"}
BLOCKS = {"p", "div", "section", "article", "header", "footer", "main", "aside", "figure", "figcaption",
          "h1", "h2", "h3", "h4", "h5", "h6", "li", "ul", "ol", "blockquote", "pre", "table", "tr", "dl", "dt", "dd",
          "hr", "nav"}


class _Conv(HTMLParser):
    def __init__(self, base_dir: str):
        super().__init__(convert_charrefs=True)
        self.base = base_dir
        self.out: List[str] = []          # 완성된 Markdown 줄
        self.buf: List[str] = []          # 지금 쌓는 인라인 글
        self.skip = 0
        self.lists: List[str] = []        # "ul" | "ol"
        self.olcount: List[int] = []
        self.heading: Optional[int] = None
        self.quote = 0
        self.pre = False
        self.table: Optional[List[List[str]]] = None
        self.row: Optional[List[str]] = None
        self.cell: Optional[List[str]] = None
        self.cell_span = 1
        self.header_row = False
        self.title = ""
        self._in_title = False

    # ── 인라인 ─────────────────────────────────────────────
    def _w(self, s: str) -> None:
        (self.cell if self.cell is not None else self.buf).append(s)

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "title":
            self._in_title = True
        if tag in SKIP:
            self.skip += 1
            return
        if self.skip:
            return
        if tag in ("b", "strong"):
            self._w("**")
        elif tag == "br":
            self._w(" <br> " if self.cell is not None else "\n")
        elif tag == "img":
            src = a.get("src") or ""
            if src and not src.startswith(("http:", "https:", "data:")):
                path = os.path.normpath(os.path.join(self.base, _html.unescape(src)))
                self._flush()
                self.out.append(f"![{a.get('alt') or ''}]({path})")
                self.out.append("")
        elif tag in ("th", "td"):
            self.cell = []
            try:
                self.cell_span = max(1, int(a.get("colspan") or 1))
            except ValueError:
                self.cell_span = 1
            if tag == "th":
                self.header_row = True
        elif tag == "tr":
            self.row = []
            self.header_row = False
        elif tag == "table":
            self._flush()
            self.table = []
        elif tag in ("ul", "ol"):
            self._flush()
            self.lists.append(tag)
            self.olcount.append(0)
        elif tag == "li":
            self._flush()
            if self.olcount:
                self.olcount[-1] += 1
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._flush()
            self.heading = int(tag[1])
        elif tag == "blockquote":
            self._flush()
            self.quote += 1
        elif tag == "pre":
            self._flush()
            self.pre = True
        elif tag in BLOCKS:
            self._flush()

    def handle_endtag(self, tag):
        if tag == "title":
            self._in_title = False
        if tag in SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag in ("b", "strong"):
            self._w("**")
        elif tag in ("th", "td") and self.cell is not None and self.row is not None:
            txt = _clean("".join(self.cell)).replace("|", "／")
            self.row.append(txt)
            self.row.extend([""] * (self.cell_span - 1))
            self.cell = None
        elif tag == "tr" and self.row is not None and self.table is not None:
            self.table.append(self.row + (["__TH__"] if self.header_row else []))
            self.row = None
        elif tag == "table" and self.table is not None:
            self._emit_table(self.table)
            self.table = None
        elif tag in ("ul", "ol"):
            self._flush()
            if self.lists:
                self.lists.pop()
                self.olcount.pop()
            if not self.lists:
                self.out.append("")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._flush()
            self.heading = None
        elif tag == "blockquote":
            self._flush()
            self.quote = max(0, self.quote - 1)
        elif tag == "pre":
            txt = "".join(self.buf)
            self.buf = []
            self.out += ["```", txt.strip("\n"), "```", ""]
            self.pre = False
        elif tag in BLOCKS:
            self._flush()

    def handle_data(self, data):
        if self._in_title:
            self.title += data
        if self.skip:
            return
        if self.pre:
            self.buf.append(data)
            return
        self._w(re.sub(r"\s+", " ", data))

    # ── 블록 내보내기 ──────────────────────────────────────
    def _flush(self) -> None:
        if self.cell is not None or self.pre:
            return
        txt = _clean("".join(self.buf))
        self.buf = []
        if not txt:
            return
        for line in [x.strip() for x in txt.split("\n") if x.strip()]:
            if self.heading:
                self.out += ["#" * self.heading + " " + line, ""]
            elif self.lists:
                depth = len(self.lists) - 1
                mark = f"{self.olcount[-1]}." if self.lists[-1] == "ol" else "-"
                self.out.append("  " * depth + f"{mark} {line}")
            elif self.quote:
                self.out += ["> " + line, ""]
            else:
                self.out += [line, ""]

    def _emit_table(self, rows: List[List[str]]) -> None:
        if not rows:
            return
        is_head = [r and r[-1] == "__TH__" for r in rows]
        rows = [r[:-1] if h else r for r, h in zip(rows, is_head)]
        n = max(len(r) for r in rows) or 1
        rows = [(r + [""] * n)[:n] for r in rows]
        head = rows[0] if is_head[0] else [""] * n
        body = rows[1:] if is_head[0] else rows
        self.out.append("| " + " | ".join(head) + " |")
        self.out.append("| " + " | ".join("---" for _ in range(n)) + " |")
        for r in body:
            self.out.append("| " + " | ".join(r) + " |")
        self.out.append("")


def _clean(s: str) -> str:
    s = s.replace("****", "")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\*\*\s+\*\*", " ", s)
    s = re.sub(r"\*\*\s*\*\*", "", s)
    return s.strip()


def html_to_markdown(src: str, base_dir: str = ".") -> str:
    """HTML 문자열 → Markdown. base_dir 은 img src 상대경로의 기준."""
    c = _Conv(base_dir)
    c.feed(src)
    c.close()
    c._flush()
    md = "\n".join(c.out)
    md = re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"
    if c.title.strip() and not md.lstrip().startswith("# "):
        md = f"# {c.title.strip()}\n\n" + md
    return md


def file_to_markdown(path: str) -> str:
    with open(path, encoding="utf-8") as f:
        return html_to_markdown(f.read(), os.path.dirname(os.path.abspath(path)))
