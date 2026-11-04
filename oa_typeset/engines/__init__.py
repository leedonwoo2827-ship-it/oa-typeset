"""조판 엔진.

html2hwpx  HTML → HWPX (표준 라이브러리만). 261103-new-etest-ebook/tools/html2hwpx 에서 가져와 몇 군데 고쳤다
           (구역 수 secCnt 맞춤 · 외톨이줄 보호 · 목록 글머리 · SVG 굽기 재시도) — 각 자리에 'oa_typeset 수정' 주석.
hwpx_md    Markdown → HWPX (양식 ZIP 복제)      ┐
docx_gen   Markdown → DOCX (python-docx)       │ aim-local(services/studio/generators) 에서 가져옴
pptx_gen   Markdown → PPTX 단순형               │
deck13     덱 spec → PPTX 13종 레이아웃         │
md_gen · mdblocks  Markdown 보조                ┘
"""
