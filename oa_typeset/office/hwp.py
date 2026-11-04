"""한글 마무리 — **실제 한컴 한글을 띄워** 문서를 다시 앉히고 PDF·쪽수를 얻는다(Windows + 한글 + pywin32).

html2hwpx 는 줄·쪽 배치 정보(linesegarray)를 쓰지 않는다 — 한글이 열 때 다시 계산한다.
그래서 쪽수·쪽 나뉨을 확인하려면 한글로 한 번 열어 봐야 한다.

하는 일
  1) 열기 → 다시 저장(HWPX) : 배치 정보가 채워진 '완성본'이 된다(다른 뷰어에서도 같게 보인다).
  2) PDF 내보내기          : 눈으로 보는 검증용 / 배포용.
  3) 쪽수                  : 판형·분량(예: 70쪽 목표) 확인.

한글 보안 모듈(FilePathCheckDLL)이 등록되어 있지 않으면 파일 접근 승인 창이 뜰 수 있다.
"""
from __future__ import annotations

import os
from typing import Dict, Optional


def available() -> bool:
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "HWPFrame.HwpObject").Close()
        return True
    except Exception:
        return False


def _save_as(hwp, path: str, fmt: str) -> bool:
    pset = hwp.HParameterSet.HFileOpenSave
    hwp.HAction.GetDefault("FileSaveAs_S", pset.HSet)
    pset.filename = path
    pset.Format = fmt
    return bool(hwp.HAction.Execute("FileSaveAs_S", pset.HSet))


def finish(hwpx: str, out: Optional[str] = None, *, pdf: bool = True, resave: bool = True) -> Dict:
    """hwpx 를 한글로 열어 다시 저장(+PDF). 결과: {"file","pdf","pages"}."""
    import pythoncom
    import win32com.client as win32

    src = os.path.abspath(hwpx)
    dst = os.path.abspath(out or hwpx)
    rep: Dict = {"file": dst, "pdf": None, "pages": None}
    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.DispatchEx("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        if not hwp.Open(src, "HWPX", "forceopen:true"):
            raise RuntimeError(f"한글이 파일을 열지 못했습니다: {src}")
        try:
            rep["pages"] = int(hwp.PageCount)
        except Exception:
            pass
        if resave:
            _save_as(hwp, dst, "HWPX")
        if pdf:
            p = os.path.splitext(dst)[0] + ".pdf"
            _save_as(hwp, p, "PDF")
            rep["pdf"] = p if os.path.exists(p) else None
    finally:
        if hwp is not None:          # 실패해도 숨은 한글 창이 남아 파일을 잠그지 않게
            try:
                hwp.Clear(1)
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return rep
