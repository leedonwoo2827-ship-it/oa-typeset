"""LLM 호출 — API 키 없이 **로그인한 CLI** 에 위임한다.

    codex  (기본) : OpenAI Codex CLI `codex exec` — ChatGPT 계정 OAuth (`codex login`)
    claude (보조) : Claude Code CLI `claude -p`   — Claude 구독 OAuth (`claude auth login`)

`ask(prompt, provider="codex")` 하나만 쓰면 된다. provider="auto" 면 codex 를 먼저 쓰고,
설치·로그인·한도 문제로 실패하면 claude 로 넘어간다.
"""
from __future__ import annotations

import os
from typing import Dict, Optional

PROVIDERS = ("codex", "claude")
DEFAULT = os.environ.get("OA_LLM", "codex")


def _run(cmd, timeout=20):
    import subprocess
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return p.returncode, (p.stdout or "") + (p.stderr or "")
    except Exception as e:  # noqa: BLE001
        return -1, str(e)


def status() -> Dict[str, Dict]:
    """각 공급자의 설치·로그인 상태(화면 상단 칩·설정 창용). 실제 CLI 에 물어본다."""
    out: Dict[str, Dict] = {}
    # ── codex ──
    try:
        from . import codex_auth
        path = codex_auth.codex_path()
        d = {"label": "Codex (ChatGPT 로그인)", "installed": bool(path), "authenticated": False,
             "email": None, "login": "codex login", "message": "", "config_error": False}
        if path:
            rc, msg = _run([path, "login", "status"])
            d["authenticated"] = rc == 0
            d["message"] = msg.strip()[:400]
            d["config_error"] = "error loading configuration" in msg.lower()
            if d["authenticated"]:
                d["email"] = codex_auth.get_account_email()
        out["codex"] = d
    except Exception as e:  # noqa: BLE001
        out["codex"] = {"label": "Codex (ChatGPT 로그인)", "installed": False, "authenticated": False, "message": str(e)}
    # ── claude ──
    try:
        from . import claude
        d = {"label": "Claude (구독 로그인)", "installed": claude.available(), "authenticated": False,
             "email": None, "login": "claude auth login", "message": ""}
        if d["installed"]:
            import json as _json
            import shutil
            rc, msg = _run([shutil.which("claude"), "auth", "status"])
            try:
                j = _json.loads(msg[msg.find("{"): msg.rfind("}") + 1])
                d["authenticated"] = bool(j.get("loggedIn"))
                d["email"] = j.get("email") if d["authenticated"] else None
            except Exception:  # noqa: BLE001
                d["authenticated"] = rc == 0 and _claude_logged_in()
            d["message"] = "" if d["authenticated"] else msg.strip()[:400]
        out["claude"] = d
    except Exception as e:  # noqa: BLE001
        out["claude"] = {"label": "Claude (구독 로그인)", "installed": False, "authenticated": False, "message": str(e)}
    return out


def _cli(provider: str):
    import shutil
    if provider == "codex":
        from .codex_auth import codex_path
        return codex_path()
    if provider == "claude":
        return shutil.which("claude")
    raise ValueError(f"모르는 공급자: {provider}")


def open_login(provider: str) -> str:
    """로그인 창 열기 — 새 콘솔 창에서 로그인 명령을 띄운다(브라우저 인증이 이어서 열림).
    창을 띄우는 것까지만 하고 기다리지 않는다. 끝나면 화면에서 [새로고침]."""
    import subprocess
    exe = _cli(provider)
    if not exe:
        raise FileNotFoundError(f"{provider} CLI 가 설치되어 있지 않습니다")
    args = [exe, "login"] if provider == "codex" else [exe, "auth", "login"]
    if os.name == "nt":
        # 새 콘솔 창 — 로그인 안내(코드·주소)를 사람이 볼 수 있게. 끝나도 창이 남도록 cmd /k.
        subprocess.Popen(["cmd", "/k"] + args, creationflags=subprocess.CREATE_NEW_CONSOLE)
    else:
        subprocess.Popen(args)
    return " ".join([os.path.basename(exe)] + args[1:])


def logout(provider: str) -> str:
    exe = _cli(provider)
    if not exe:
        raise FileNotFoundError(f"{provider} CLI 가 설치되어 있지 않습니다")
    rc, msg = _run([exe, "logout"] if provider == "codex" else [exe, "auth", "logout"], timeout=30)
    return msg.strip() or ("로그아웃했습니다" if rc == 0 else f"종료 코드 {rc}")


def fix_codex_config() -> Dict:
    r"""~/.codex/config.toml 의 흔한 오류를 고친다: 큰따옴표 안 윈도우 경로("C:\Users…" 의 \U 를 유니코드로 읽음).
    윈도우 경로(C:\…)가 든 큰따옴표 문자열만 작은따옴표(리터럴 문자열)로 바꾼다. 고치기 전에 백업을 남긴다."""
    import re
    import shutil
    import time
    from pathlib import Path
    cfg = Path(os.environ.get("CODEX_HOME") or (Path.home() / ".codex")) / "config.toml"
    if not cfg.is_file():
        raise FileNotFoundError(str(cfg))
    text = cfg.read_text(encoding="utf-8")
    changed = []

    def rep(m):
        body = m.group(1)
        # 윈도우 경로(드라이브 문자 + 역슬래시)이고, 작은따옴표·이중 역슬래시가 없을 때만 — 일부러 쓴 \n 같은 이스케이프는 건드리지 않는다
        if re.search(r"[A-Za-z]:\\", body) and "'" not in body and "\\\\" not in body:
            changed.append(body[:60])
            return "'" + body + "'"
        return m.group(0)

    new = re.sub(r'"([^"\n]*)"', rep, text)
    if not changed:
        return {"changed": 0, "backup": None, "file": str(cfg)}
    backup = cfg.with_name(f"config.toml.bak-{time.strftime('%Y%m%d-%H%M%S')}")
    shutil.copyfile(cfg, backup)
    cfg.write_text(new, encoding="utf-8")
    return {"changed": len(changed), "backup": str(backup), "file": str(cfg), "items": changed}


def ask(prompt: str, provider: Optional[str] = None, *, model: Optional[str] = None,
        system: Optional[str] = None, timeout: int = 900) -> str:
    """프롬프트 한 번 → 응답 텍스트."""
    provider = (provider or DEFAULT).lower()
    if provider == "auto":
        try:
            return ask(prompt, "codex", model=model, system=system, timeout=timeout)
        except Exception as e:  # noqa: BLE001
            print(f"[llm] codex 실패 → claude 로 넘김: {e}", flush=True, file=__import__("sys").stderr)
            return ask(prompt, "claude", model=None, system=system, timeout=timeout)
    if provider == "codex":
        from .codex_runner import client
        client.timeout = max(client.timeout, timeout)
        msgs = ([{"role": "system", "content": system}] if system else []) + [{"role": "user", "content": prompt}]
        return client.chat(model, msgs).text
    if provider == "claude":
        from . import claude
        full = (system + "\n\n" + prompt) if system else prompt
        return claude.ask(full, model=model or "opus", timeout=timeout)
    raise ValueError(f"모르는 LLM 공급자: {provider} (codex | claude | auto)")
