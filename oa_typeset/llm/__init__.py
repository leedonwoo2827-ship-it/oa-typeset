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


def status() -> Dict[str, Dict]:
    """각 공급자의 설치·로그인 상태(화면 상단 칩용)."""
    out: Dict[str, Dict] = {}
    try:
        from . import codex_auth
        out["codex"] = {"label": "Codex (ChatGPT 로그인)", "installed": codex_auth.is_installed(),
                        "authenticated": codex_auth.is_authenticated(),
                        "email": codex_auth.get_account_email(), "login": "codex login"}
    except Exception as e:  # noqa: BLE001
        out["codex"] = {"label": "Codex", "installed": False, "authenticated": False, "error": str(e)}
    try:
        from . import claude
        ok = claude.available()
        out["claude"] = {"label": "Claude (구독 로그인)", "installed": ok,
                         "authenticated": ok and _claude_logged_in(), "login": "claude auth login"}
    except Exception as e:  # noqa: BLE001
        out["claude"] = {"label": "Claude", "installed": False, "authenticated": False, "error": str(e)}
    return out


def _claude_logged_in() -> bool:
    """~/.claude/.credentials.json 이 있으면 로그인한 것으로 본다(내용은 읽지 않는다)."""
    from pathlib import Path
    return (Path.home() / ".claude" / ".credentials.json").is_file() or bool(os.environ.get("CLAUDE_CODE_OAUTH_TOKEN"))


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
