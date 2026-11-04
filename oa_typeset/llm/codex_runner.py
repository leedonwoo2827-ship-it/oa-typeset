"""CodexClient — OpenAI Codex CLI(`codex`)를 비대화식으로 호출하는 LLM 래퍼.

설계 의도:
- API 키를 쓰지 않는다. 인증/할당량은 `codex` 가 담당(사용자가 `codex login` 1회, ChatGPT OAuth).
- 회사 LiteLLM 프록시와 동일한 `.chat(model, messages, max_tokens) -> resp(.text)` 시그니처(드롭인).
- `codex exec` 는 최종 답을 stdout 으로 내보내므로 PTY 불필요(일반 subprocess 캡처).
- 코딩 에이전트이므로 `-s read-only -a never` + 가드 프롬프트로 파일변경/도구사용을 막는다.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from .errors import (
    LLMError, LLMNotInstalled, LLMNotAuthenticated, LLMQuotaExceeded,
)
from .codex_auth import CODEX_BIN, codex_path

CODEX_TIMEOUT = int(os.environ.get("CODEX_EXEC_TIMEOUT", "300"))
CODEX_MODEL_ENV = os.environ.get("CODEX_MODEL", "").strip()
_MODEL_FILE = Path(os.environ.get("OA_DATA_DIR", Path.home() / ".oa-typeset")) / "codex_model.json"

_GUARD_SYSTEM = (
    "You are a text and document generation assistant. "
    "Do NOT run shell commands, do NOT read or write files, do NOT use any tools, "
    "and do NOT perform any coding or repository actions. "
    "Respond with ONLY the requested content as your message — nothing else."
)


class CodexError(LLMError):
    pass


class CodexNotInstalled(LLMNotInstalled):
    pass


class CodexNotAuthenticated(LLMNotAuthenticated):
    pass


class CodexQuotaExceeded(LLMQuotaExceeded):
    pass


class CodexResult:
    __slots__ = ("text", "raw")

    def __init__(self, text: str, raw: Any = None):
        self.text = text
        self.raw = raw


# ── 모델 선택 저장 ───────────────────────────────────────────────────────────
def get_model() -> str:
    try:
        if _MODEL_FILE.is_file():
            m = json.loads(_MODEL_FILE.read_text(encoding="utf-8")).get("model", "")
            if isinstance(m, str) and m.strip():
                return m.strip()
    except Exception:
        pass
    return CODEX_MODEL_ENV


def set_model(name: str) -> None:
    try:
        _MODEL_FILE.parent.mkdir(parents=True, exist_ok=True)
        _MODEL_FILE.write_text(json.dumps({"model": (name or "").strip()}, ensure_ascii=False),
                               encoding="utf-8")
    except Exception:
        pass


_MODELS_CACHE: Optional[List[str]] = None
_MODEL_ID_RE = re.compile(r"^(gpt|o\d|chatgpt|codex)[A-Za-z0-9.\-_]*$", re.IGNORECASE)


def _parse_models(out: str) -> List[str]:
    """`codex debug models` 출력 파싱. 최신 codex 는 JSON({"models":[{slug,..}]})을 낸다.
    실패하면 줄 단위 휴리스틱으로 폴백. visibility='hide' 인 내부 모델은 제외."""
    out = (out or "").strip()
    # 1) JSON 우선 (출력 앞뒤 잡음 대비해 첫 '{' ~ 마지막 '}' 만 잘라 시도)
    if "{" in out and '"models"' in out:
        try:
            blob = out[out.index("{"): out.rindex("}") + 1]
            data = json.loads(blob)
            ids: List[str] = []
            for m in data.get("models", []):
                if not isinstance(m, dict):
                    continue
                if str(m.get("visibility", "")).lower() == "hide":
                    continue
                slug = (m.get("slug") or m.get("id") or m.get("name") or "").strip()
                if slug and slug not in ids:
                    ids.append(slug)
            if ids:
                return ids
        except Exception:
            pass
    # 2) 폴백: 줄 단위에서 모델 ID 형태 토큰 추출
    models: List[str] = []
    for line in out.splitlines():
        s = line.strip().strip("-*•> \t")
        tok = s.split()[0] if s.split() else ""
        if _MODEL_ID_RE.match(tok) and tok not in models:
            models.append(tok)
    return models


def list_models(force: bool = False) -> List[str]:
    """`codex debug models` 출력에서 모델 ID 추출(캐시). 실패 시 빈 목록."""
    global _MODELS_CACHE
    if _MODELS_CACHE is not None and not force:
        return _MODELS_CACHE
    path = codex_path()
    if not path:
        return []
    models: List[str] = []
    try:
        p = subprocess.run([path, "debug", "models"], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=45)
        models = _parse_models((p.stdout or "") + "\n" + (p.stderr or ""))
    except Exception:
        pass
    if models:
        _MODELS_CACHE = models
    return models


def _compose_prompt(messages: List[Dict[str, str]]) -> str:
    parts: List[str] = [_GUARD_SYSTEM]
    for m in messages:
        role = (m.get("role") or "user").lower()
        content = m.get("content") or ""
        if not content:
            continue
        if role == "system":
            parts.append(f"[지침]\n{content}")
        elif role == "assistant":
            parts.append(f"[이전 답변]\n{content}")
        else:
            parts.append(f"[요청]\n{content}")
    return "\n\n".join(parts)


def _classify_error(stdout: str, stderr: str, rc: int) -> LLMError:
    blob = f"{stdout}\n{stderr}".lower()
    if any(k in blob for k in ("not logged in", "please run", "codex login", "unauthorized",
                               "401", "authenticate", "sign in")):
        return CodexNotAuthenticated(
            "codex 에 ChatGPT 로그인이 필요합니다. 터미널에서 `codex login` 을 실행해 로그인하세요."
        )
    # "out of credits" 는 워크스페이스 크레딧 소진 — 한도 초과와 안내가 달라 먼저 본다.
    if "out of credits" in blob or "refill" in blob:
        return CodexQuotaExceeded(
            "ChatGPT 워크스페이스 크레딧이 소진되었습니다. 워크스페이스 소유자에게 충전을 요청하거나, "
            "[연결 상태]에서 다른 계정으로 로그인하세요(codex logout → codex login)."
        )
    if any(k in blob for k in ("quota", "rate limit", "429", "too many requests",
                               "usage limit", "exceeded")):
        return CodexQuotaExceeded(
            "ChatGPT 계정 사용 한도를 초과했습니다. 잠시 후 다시 시도하거나 상위 구독(Plus/Pro)을 사용하세요."
        )
    snippet = (stderr or stdout or "").strip()[:300]
    return CodexError(f"codex 호출 실패(exit={rc}): {snippet}")


class CodexClient:
    def __init__(self, bin_path: Optional[str] = None, default_model: Optional[str] = None,
                 timeout: Optional[int] = None):
        self.bin = bin_path or CODEX_BIN
        self.default_model = default_model
        self.timeout = timeout or CODEX_TIMEOUT

    def chat(self, model: Optional[str], messages: List[Dict[str, str]],
             max_tokens: int = 4000) -> CodexResult:
        """LiteLLM 프록시와 시그니처를 맞춘 드롭인.

        **max_tokens 는 받기만 하고 쓰지 않는다** — `codex exec` 에 출력 길이를 지정하는
        옵션이 없다. 응답이 짧게 끊기더라도 호출부에서 이 값을 올려 봐야 아무 일도 일어나지
        않는다(그렇게 시간을 버린 적이 있다). 분량은 프롬프트로 조절해야 한다.
        """
        prompt = _compose_prompt(messages)
        text = self._run(prompt, model)
        return CodexResult(text)

    def quick(self, prompt: str, model: Optional[str] = None) -> str:
        return self._run(_compose_prompt([{"role": "user", "content": prompt}]), model)

    def _run(self, prompt: str, model: Optional[str]) -> str:
        path = codex_path() if self.bin == CODEX_BIN else (
            self.bin if os.path.isfile(self.bin) else shutil.which(self.bin))
        if not path:
            raise CodexNotInstalled(
                "OpenAI Codex CLI(`codex`)가 설치되어 있지 않습니다. "
                "docs/openai-codex/install.md 를 참고해 설치 후 `codex login` 으로 로그인하세요."
            )
        sel = (model or "").strip() or get_model()
        model_opt = (["-m", sel] if sel else [])
        # 최종 메시지를 파일로 받는다(-o, 가장 깔끔). stdout 은 폴백.
        out_path = None
        try:
            fd, out_path = tempfile.mkstemp(prefix="codex_out_", suffix=".txt")
            os.close(fd)
        except Exception:
            out_path = None
        # 프롬프트는 **stdin** 으로 전달한다. 위치 인자로 주면 Windows 명령행 길이 제한
        # (~32KB)에 걸려 긴 프롬프트(첨부 다수)가 잘려 codex 가 빈 입력으로 되묻는다.
        # 위치 인자를 생략하면 codex 가 stdin 에서 지시를 읽는다.
        # (-a/approval 플래그는 이 버전에 없음. read-only 샌드박스로 파일변경 차단.)
        cmd = [path, "exec", "--skip-git-repo-check", "-s", "read-only", *model_opt]
        if out_path:
            cmd += ["-o", out_path]
        try:
            proc = subprocess.run(cmd, input=prompt, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace",
                                  timeout=self.timeout + 30)
        except FileNotFoundError as e:
            raise CodexNotInstalled(f"codex 실행 실패: {e}") from e
        except subprocess.TimeoutExpired as e:
            raise CodexError(f"codex 응답 시간 초과({self.timeout}s).") from e

        # 최종 메시지: -o 파일 우선(진행상황/토큰 트레일러 없이 답만), 없으면 stdout.
        text = ""
        if out_path:
            try:
                with open(out_path, "r", encoding="utf-8", errors="replace") as f:
                    text = f.read().strip()
            except Exception:
                text = ""
            try:
                os.remove(out_path)
            except Exception:
                pass
        if not text:
            text = (proc.stdout or "").strip()

        if not text:
            raise _classify_error(proc.stdout or "", proc.stderr or "", proc.returncode)
        return text


client = CodexClient()
