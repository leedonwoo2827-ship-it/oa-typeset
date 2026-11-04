"""양식(샘플) 폴더 — "샘플을 넣으면 그대로 조판".

양식 하나 = 폴더 하나:

    forms/<이름>/
        form.hwpx         한글 양식 (글꼴·문단 스타일·쪽 설정·바탕쪽을 가져온다)
        form.docx         Word 양식 (명명 스타일 '제목 1~3'·'본문'·'목록')      — 선택
        form.pptx         PowerPoint 시안 (배색 15색·글꼴을 가져온다)          — 선택
        form.json         이름·판형·스타일 짝·집필 지침                         — 선택
        logo_left.png / logo_right.png   슬라이드 로고                         — 선택

찾는 곳(앞이 우선): ① 함수 인자 forms_dir ② 환경변수 OA_FORMS_DIR(; 로 여러 개)
③ 현재 폴더의 forms/ ④ 패키지에 들어 있는 기본 양식.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

PKG_FORMS = Path(__file__).resolve().parent / "forms"


@dataclass
class Form:
    name: str
    dir: Path
    meta: Dict = field(default_factory=dict)

    def file(self, ext: str) -> Optional[str]:
        p = self.dir / f"form.{ext}"
        return str(p) if p.is_file() else None

    @property
    def title(self) -> str:
        return self.meta.get("title") or self.name

    @property
    def page(self):
        """html2hwpx 의 page 인자. "keep"(양식 그대로) 또는 (가로mm, 세로mm)."""
        v = str(self.meta.get("page") or "keep")
        if v == "keep":
            return "keep"
        w, h = v.lower().split("x")
        return (float(w), float(h))

    @property
    def styles(self) -> Dict[str, str]:
        """HTML → 한글 양식 스타일 짝. 예: {"body": "본문", "h1": "제목 1"}"""
        return dict(self.meta.get("styles") or {})

    @property
    def guide(self) -> str:
        """LLM 집필 지침(양식마다 다르게)."""
        g = self.meta.get("guide") or ""
        return "\n".join(g) if isinstance(g, list) else str(g)

    def logo(self, side: str) -> Optional[str]:
        p = self.dir / f"logo_{side}.png"
        return str(p) if p.is_file() else None

    def info(self) -> Dict:
        return {"name": self.name, "title": self.title, "dir": str(self.dir),
                "hwpx": bool(self.file("hwpx")), "docx": bool(self.file("docx")), "pptx": bool(self.file("pptx")),
                "page": self.meta.get("page") or "keep", "description": self.meta.get("description", "")}


def search_dirs(forms_dir: Optional[str] = None) -> List[Path]:
    dirs: List[Path] = []
    if forms_dir:
        dirs.append(Path(forms_dir))
    for d in (os.environ.get("OA_FORMS_DIR") or "").split(";"):
        if d.strip():
            dirs.append(Path(d.strip()))
    dirs.append(Path.cwd() / "forms")
    dirs.append(PKG_FORMS)
    seen, out = set(), []
    for d in dirs:
        r = d.resolve()
        if r not in seen and r.is_dir():
            seen.add(r)
            out.append(r)
    return out


def _load(d: Path) -> Form:
    meta = {}
    j = d / "form.json"
    if j.is_file():
        meta = json.loads(j.read_text(encoding="utf-8"))
    return Form(d.name, d, meta)


def list_forms(forms_dir: Optional[str] = None) -> List[Form]:
    out: Dict[str, Form] = {}
    for base in search_dirs(forms_dir):
        for d in sorted(base.iterdir()):
            if d.is_dir() and any((d / f"form.{e}").is_file() for e in ("hwpx", "docx", "pptx")):
                out.setdefault(d.name, _load(d))
    return list(out.values())


def get_form(name_or_path: Optional[str], forms_dir: Optional[str] = None) -> Form:
    """이름(폴더명) 또는 경로(폴더, 혹은 .hwpx 파일 하나)로 양식을 찾는다."""
    if not name_or_path:
        forms = list_forms(forms_dir)
        if not forms:
            raise FileNotFoundError("양식이 없습니다. forms/<이름>/form.hwpx 를 넣어 주세요.")
        return forms[0]
    p = Path(name_or_path)
    if p.is_dir():
        return _load(p)
    if p.is_file() and p.suffix.lower() in (".hwpx", ".docx", ".pptx"):
        # 파일 하나만 준 경우: 그 파일이 있는 폴더를 양식으로 보되, 이름이 form.* 가 아니어도 받아 준다.
        f = Form(p.stem, p.parent, {"page": "keep"})
        f.file = lambda ext, _p=p: str(_p) if _p.suffix.lower() == "." + ext else None  # type: ignore[method-assign]
        return f
    for f in list_forms(forms_dir):
        if f.name == name_or_path:
            return f
    names = ", ".join(f.name for f in list_forms(forms_dir)) or "(없음)"
    raise FileNotFoundError(f"양식 '{name_or_path}' 을 찾지 못했습니다. 있는 양식: {names}")


# ── 양식 살펴보기 / 스타일 짝 제안 ─────────────────────────────────────────
def hwpx_styles(path: str) -> List[str]:
    """양식 .hwpx 의 문단 스타일 이름 목록."""
    with zipfile.ZipFile(path) as z:
        h = z.read("Contents/header.xml").decode("utf-8", "replace")
    return re.findall(r'<hh:style [^>]*?type="PARA"[^>]*?name="([^"]*)"', h) or \
        re.findall(r'<hh:style [^>]*?name="([^"]*)"', h)


_GUESS = {
    "body": ["*본문", "본문", "바탕글"],
    "h1": ["*절제목", "제목 1", "제목1", "개요 1"],
    "h2": ["*1.제목", "제목 2", "제목2", "개요 2"],
    "h3": ["*(1)제목", "제목 3", "제목3", "개요 3"],
    "h4": ["*1)제목", "제목 4", "개요 4"],
    "cell": ["*표내용", "표내용", "표 내용"],
}


def suggest_styles(path: str) -> Dict[str, str]:
    """양식의 스타일 이름을 보고 HTML 태그 짝을 제안한다(form.json 의 styles 초안)."""
    have = [s.strip() for s in hwpx_styles(path)]
    out = {}
    for key, cands in _GUESS.items():
        for c in cands:
            if c in have:
                out[key] = c
                break
    return out
