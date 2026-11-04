"""로컬 웹 스튜디오 — 양식 고르기 → (자료 넣고 LLM 집필) → HTML 검토 → 조판·내려받기.

    oa-typeset studio            → http://127.0.0.1:7090
    pip install "oa-typeset[studio]" 필요 (fastapi · uvicorn · python-multipart)

다른 프로그램(젠오피스 등)이 HTTP 로 부를 수 있는 곳:
    POST /api/typeset   multipart: file(.html|.md|.zip), form, formats="hwpx,docx,pptx", finish=0|1
                        → {"ok":true, "job":…, "files":{"hwpx":"/files/<job>/x.hwpx", …}, "result":{…}}
    GET  /api/forms     양식 목록
    GET  /files/<job>/<이름>  결과 내려받기
작업 파일은 OA_DATA_DIR(기본: 현재 폴더의 data/) 아래 jobs/<작업>/ 에 쌓인다. 이 PC 밖으로 나가지 않는다.
"""

import json
import os
import re
import shutil
import threading
import time
import traceback
import uuid
import zipfile
from pathlib import Path
from typing import Dict, List, Optional

DATA = Path(os.environ.get("OA_DATA_DIR") or (Path.cwd() / "data")).resolve()
JOBS = DATA / "jobs"
USER_FORMS = Path(os.environ.get("OA_FORMS_DIR") or (Path.cwd() / "forms")).resolve()
STATIC = Path(__file__).resolve().parent / "static"
_STATE: Dict[str, Dict] = {}
_LOCK = threading.Lock()


def _job_dir(job: str) -> Path:
    if not re.fullmatch(r"[0-9a-f\-]{8,40}", job or ""):
        raise ValueError("잘못된 작업 번호")
    return JOBS / job


def _new_job(kind: str) -> str:
    job = time.strftime("%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
    job = re.sub(r"[^0-9a-f\-]", "", job.replace("-", "-"))
    (JOBS / job).mkdir(parents=True, exist_ok=True)
    with _LOCK:
        _STATE[job] = {"job": job, "kind": kind, "status": "running", "log": [], "files": {}, "started": time.time()}
    return job


def _log(job: str, msg: str) -> None:
    with _LOCK:
        _STATE.setdefault(job, {"log": []})["log"].append(time.strftime("[%H:%M:%S] ") + str(msg))


def _safe_name(name: str) -> str:
    name = os.path.basename(name or "file")
    return re.sub(r'[\\/:*?"<>|]', "_", name) or "file"


def _files_of(job: str) -> Dict[str, str]:
    d = _job_dir(job)
    out = {}
    for p in d.rglob("*"):
        if p.is_file() and p.suffix.lower() in (".hwpx", ".docx", ".pptx", ".pdf", ".html", ".md"):
            out[str(p.relative_to(d)).replace("\\", "/")] = f"/files/{job}/{str(p.relative_to(d)).replace(os.sep, '/')}"
    return out


def _unpack(upload_path: Path, into: Path) -> Path:
    """.zip 이면 풀어서 그 안의 첫 .html(없으면 .md)을, 아니면 파일 그대로를 원고로."""
    if upload_path.suffix.lower() != ".zip":
        return upload_path
    root = into / "src"
    with zipfile.ZipFile(upload_path) as z:
        for m in z.infolist():
            target = (root / m.filename).resolve()
            if not str(target).startswith(str(root.resolve())):
                continue  # zip slip 방지
            if m.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(z.read(m))
    for pat in ("*.html", "*.htm", "*.md"):
        found = sorted(root.rglob(pat))
        if found:
            return found[0]
    raise ValueError("zip 안에 .html 이나 .md 원고가 없습니다")


def create_app():
    from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
    from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

    from .. import llm
    from ..forms import get_form, hwpx_styles, list_forms, suggest_styles
    from ..typeset import typeset

    app = FastAPI(title="oa-typeset studio")
    JOBS.mkdir(parents=True, exist_ok=True)

    def forms_dir() -> Optional[str]:
        return str(USER_FORMS) if USER_FORMS.is_dir() else None

    @app.middleware("http")
    async def local_only(request: Request, call_next):
        # 로그인·집필은 이 PC 에서만(LAN 에 열어도 남이 LLM 계정을 쓰지 못하게). 조판·내려받기는 허용.
        host = request.client.host if request.client else ""
        if request.url.path.startswith(("/api/write", "/api/forms/upload")) and host not in ("127.0.0.1", "::1", "localhost"):
            return JSONResponse({"ok": False, "error": "이 기능은 이 PC 에서만 쓸 수 있습니다"}, status_code=403)
        return await call_next(request)

    @app.get("/", response_class=HTMLResponse)
    def index():
        return (STATIC / "index.html").read_text(encoding="utf-8")

    @app.get("/api/forms")
    def api_forms():
        return [f.info() for f in list_forms(forms_dir())]

    @app.get("/api/forms/{name}/styles")
    def api_form_styles(name: str):
        f = get_form(name, forms_dir())
        p = f.file("hwpx")
        return {"styles": hwpx_styles(p) if p else [], "suggest": suggest_styles(p) if p else {}, "current": f.styles}

    @app.post("/api/forms/upload")
    async def api_form_upload(name: str = Form(...), files: List[UploadFile] = File(...)):
        name = _safe_name(name).strip(". ")
        d = USER_FORMS / name
        d.mkdir(parents=True, exist_ok=True)
        got = []
        for uf in files:
            ext = Path(uf.filename or "").suffix.lower()
            if ext in (".hwpx", ".docx", ".pptx"):
                (d / f"form{ext}").write_bytes(await uf.read())
                got.append(ext[1:])
            elif ext == ".png" and "logo" in (uf.filename or "").lower():
                side = "left" if "left" in uf.filename.lower() else "right"
                (d / f"logo_{side}.png").write_bytes(await uf.read())
        meta_p = d / "form.json"
        meta = json.loads(meta_p.read_text(encoding="utf-8")) if meta_p.is_file() else {"title": name, "page": "keep"}
        if (d / "form.hwpx").is_file() and not meta.get("styles"):
            meta["styles"] = suggest_styles(str(d / "form.hwpx"))
        meta_p.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"ok": True, "form": name, "saved": got, "styles": meta.get("styles")}

    @app.get("/api/llm/status")
    def api_llm_status():
        return llm.status()

    @app.post("/api/typeset")
    async def api_typeset(file: UploadFile = File(...), form: str = Form(""), formats: str = Form("hwpx"),
                          finish: str = Form("0"), provider: str = Form("")):
        job = _new_job("typeset")
        d = _job_dir(job)
        up = d / _safe_name(file.filename or "원고.html")
        up.write_bytes(await file.read())
        try:
            src = _unpack(up, d)
            r = typeset(str(src), form or None, formats.split(","), str(d / "out"), forms_dir=forms_dir(),
                        finish=finish in ("1", "true", "on"), provider=provider or None)
            files = {k: f"/files/{job}/out/{Path(v['path']).name}" for k, v in r.items()
                     if isinstance(v, dict) and v.get("path")}
            for k, v in r.items():
                if isinstance(v, dict) and v.get("pdf"):
                    files["pdf"] = f"/files/{job}/out/{Path(v['pdf']).name}"
            with _LOCK:
                _STATE[job].update(status="done", files=files, result=r)
            return {"ok": True, "job": job, "files": files, "result": r}
        except Exception as e:  # noqa: BLE001
            with _LOCK:
                _STATE[job].update(status="error", error=str(e))
            return JSONResponse({"ok": False, "job": job, "error": str(e)}, status_code=400)

    @app.post("/api/write")
    async def api_write(brief: str = Form(...), form: str = Form(""), pages: int = Form(0),
                        provider: str = Form("codex"), formats: str = Form("hwpx,docx"),
                        files: List[UploadFile] = File(default=[])):
        job = _new_job("write")
        d = _job_dir(job)
        srcs = []
        for uf in files:
            if not uf.filename:
                continue
            p = d / "sources" / _safe_name(uf.filename)
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_bytes(await uf.read())
            srcs.append(str(p))

        def run():
            try:
                from ..writer import write_book
                html = write_book(srcs, brief, form or None, str(d / "work"), pages or None, provider,
                                  forms_dir=forms_dir(), log=lambda m: _log(job, m))
                _log(job, "원고 완성 → 조판")
                typeset(html, form or None, formats.split(","), str(d / "out"), forms_dir=forms_dir())
                with _LOCK:
                    _STATE[job].update(status="done", files=_files_of(job))
                _log(job, "끝")
            except Exception as e:  # noqa: BLE001
                _log(job, "실패: " + str(e))
                _log(job, traceback.format_exc()[-800:])
                with _LOCK:
                    _STATE[job].update(status="error", error=str(e), files=_files_of(job))

        threading.Thread(target=run, daemon=True).start()
        return {"ok": True, "job": job}

    @app.get("/api/jobs/{job}")
    def api_job(job: str):
        with _LOCK:
            st = dict(_STATE.get(job) or {})
        if not st:
            d = _job_dir(job)
            if not d.is_dir():
                raise HTTPException(404, "작업 없음")
            st = {"job": job, "status": "unknown"}
        st["files"] = _files_of(job)
        return st

    @app.get("/api/jobs")
    def api_jobs():
        out = []
        for d in sorted(JOBS.glob("*"), reverse=True)[:30]:
            if d.is_dir():
                with _LOCK:
                    s = _STATE.get(d.name, {})
                out.append({"job": d.name, "status": s.get("status", "done"), "kind": s.get("kind", ""),
                            "files": _files_of(d.name)})
        return out

    @app.get("/files/{job}/{path:path}")
    def files(job: str, path: str):
        base = _job_dir(job).resolve()
        p = (base / path).resolve()
        if not str(p).startswith(str(base)) or not p.is_file():
            raise HTTPException(404, "파일 없음")
        inline = p.suffix.lower() in (".html", ".pdf")
        return FileResponse(str(p), filename=None if inline else p.name)

    return app


def run(host: str = "127.0.0.1", port: int = 7090) -> None:
    try:
        import uvicorn
    except ImportError:
        raise SystemExit('스튜디오에는 추가 설치가 필요합니다:  pip install "oa-typeset[studio]"')
    print(f"oa-typeset studio → http://{host if host != '0.0.0.0' else '127.0.0.1'}:{port}   (데이터: {DATA})")
    uvicorn.run(create_app(), host=host, port=port, log_level="warning")
