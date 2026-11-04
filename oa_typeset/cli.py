"""명령줄 — `oa-typeset <명령>` 또는 `python -m oa_typeset <명령>`.

  build   원고.html [--form 이름] [-f hwpx,docx,pptx] [-o 출력폴더] [--finish] [--deck spec.json] [--llm codex]
  write   자료... --brief "요청" [--form 이름] [--pages 70] [-o work/폴더] [--llm codex|claude|auto] [--build hwpx,docx]
  forms   양식 목록
  styles  양식.hwpx               (스타일 이름과 form.json 'styles' 초안)
  finish  파일.pptx|파일.hwpx      (Office 마무리만: 끝줄 정리·넘침 / 한글 다시 앉히기·PDF)
  llm     LLM 로그인 상태
  studio  [--port 7090]           (로컬 웹 화면)

--json 을 주면 결과를 JSON 한 줄로 stdout 에 낸다(다른 프로그램이 부를 때).
종료 코드: 0 성공 · 1 실패 · 2 사용법 오류
"""
from __future__ import annotations

import argparse
import json
import sys


def _print(obj, as_json: bool) -> None:
    if as_json:
        print(json.dumps(obj, ensure_ascii=False, default=str))
    else:
        print(json.dumps(obj, ensure_ascii=False, indent=1, default=str))


def main(argv=None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    ap = argparse.ArgumentParser(prog="oa-typeset", description="HTML 원고 + 양식 → hwpx · docx · pptx")
    ap.add_argument("--forms-dir", help="양식 폴더(기본: ./forms + 패키지 기본 양식)")
    ap.add_argument("--json", action="store_true", help="결과를 JSON 한 줄로")
    sub = ap.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="조판")
    b.add_argument("src", help="원고 .html (또는 .md)")
    b.add_argument("--form", help="양식 이름·폴더·.hwpx 경로 (기본: 첫 양식)")
    b.add_argument("-f", "--formats", default="hwpx", help="hwpx,docx,pptx,md 중 쉼표로")
    b.add_argument("-o", "--out", help="출력 폴더 (기본: 원고 옆)")
    b.add_argument("--finish", action="store_true", help="Office 마무리(끝줄 정리·한글 PDF) — Windows+Office")
    b.add_argument("--deck", help="pptx 덱 spec JSON 파일")
    b.add_argument("--llm", help="pptx 덱을 LLM 이 구성 (codex|claude|auto)")
    b.add_argument("--head", help="hwpx 바탕쪽 머리말 글자 (기본: <title>)")

    w = sub.add_parser("write", help="LLM 집필 → book.html")
    w.add_argument("sources", nargs="*", help="자료 파일(pdf·hwp·hwpx·docx·xlsx·pptx·txt·md)")
    w.add_argument("--brief", required=True, help="무엇을 누구에게 쓰는지")
    w.add_argument("--form")
    w.add_argument("--pages", type=int)
    w.add_argument("-o", "--out", default="work")
    w.add_argument("--llm", default=None, help="codex(기본) | claude | auto")
    w.add_argument("--only", help="이 장만 (예: 3 또는 2,5)")
    w.add_argument("--redo", action="store_true", help="이미 쓴 장도 다시")
    w.add_argument("--build", help="다 쓰면 바로 조판 (예: hwpx,docx,pptx)")

    sub.add_parser("forms", help="양식 목록")
    s = sub.add_parser("styles", help="양식 .hwpx 의 스타일과 짝 제안")
    s.add_argument("hwpx")
    fz = sub.add_parser("finish", help="Office 마무리만")
    fz.add_argument("path")
    fz.add_argument("--png", help="pptx 슬라이드 PNG 폴더")
    sub.add_parser("llm", help="LLM 로그인 상태")
    st = sub.add_parser("studio", help="로컬 웹 화면")
    st.add_argument("--host", default="127.0.0.1")
    st.add_argument("--port", type=int, default=7090)

    a = ap.parse_args(argv)
    try:
        if a.cmd == "build":
            from .typeset import typeset
            r = typeset(a.src, a.form, a.formats.split(","), a.out, forms_dir=a.forms_dir, finish=a.finish,
                        deck_spec=a.deck, provider=a.llm, head=a.head)
            if a.json:
                _print({"ok": True, **r}, True)
            else:
                for k in ("hwpx", "docx", "pptx", "md"):
                    if k in r:
                        extra = " · ".join(str(n) for n in r[k].get("notes", []))
                        print(f"→ {r[k]['path']}" + (f"   ({extra})" if extra else ""))
        elif a.cmd == "write":
            from .writer import write_book
            only = [int(x) for x in a.only.split(",")] if a.only else None
            log = (lambda m: print(m, file=sys.stderr, flush=True)) if a.json else print
            path = write_book(a.sources, a.brief, a.form, a.out, a.pages, a.llm, forms_dir=a.forms_dir,
                              only=only, redo=a.redo, log=log)
            res = {"ok": True, "html": path}
            if a.build:
                from .typeset import typeset
                res["build"] = typeset(path, a.form, a.build.split(","), a.out, forms_dir=a.forms_dir)
            _print(res, a.json) if a.json else print(f"→ {path}")
        elif a.cmd == "forms":
            from .forms import list_forms
            fs = [f.info() for f in list_forms(a.forms_dir)]
            if a.json:
                _print(fs, True)
            else:
                for f in fs:
                    kinds = "·".join(k for k in ("hwpx", "docx", "pptx") if f[k])
                    print(f"{f['name']:<20} {kinds:<16} {f['description']}")
        elif a.cmd == "styles":
            from .forms import hwpx_styles, suggest_styles
            _print({"styles": hwpx_styles(a.hwpx), "suggest": suggest_styles(a.hwpx)}, a.json)
        elif a.cmd == "finish":
            if a.path.lower().endswith(".pptx"):
                from .office import ppt
                r = ppt.finish(a.path, png_dir=a.png)
                _print(r, True) if a.json else print(ppt.summary(r))
            else:
                from .office import hwp
                _print(hwp.finish(a.path), a.json)
        elif a.cmd == "llm":
            from . import llm
            _print(llm.status(), a.json)
        elif a.cmd == "studio":
            from .studio.app import run
            run(a.host, a.port)
        return 0
    except Exception as e:  # noqa: BLE001
        if a.json:
            _print({"ok": False, "error": str(e), "type": type(e).__name__}, True)
        else:
            print(f"실패: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
