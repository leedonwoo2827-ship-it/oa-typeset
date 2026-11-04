"""Office 마무리(선택) — 실제 PowerPoint·한글을 자동으로 띄워 조판을 다듬고 확인한다.

조판 자체(hwpx/docx/pptx 파일 만들기)는 Office 없이 된다. 이 층은 Windows 에
PowerPoint·한컴 한글·pywin32 가 있을 때만 쓰는 '마무리 손질'이다.

    ppt.finish(path)  — 끝줄 정리(자간), 넘친 상자 보고, 슬라이드 PNG
    hwp.finish(path)  — 한글로 다시 앉히기(배치 확정), PDF, 쪽수
"""
from . import hwp, ppt  # noqa: F401
