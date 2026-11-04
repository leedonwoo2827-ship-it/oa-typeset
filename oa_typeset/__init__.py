"""oa-typeset — HTML 원고 + 양식(샘플) → 한글(hwpx) · Word(docx) · PowerPoint(pptx) 조판.

    from oa_typeset import typeset, list_forms
    typeset("book.html", form="교재형-46배판", formats=["hwpx", "docx", "pptx"], out_dir="out")

    from oa_typeset.writer import write_book          # LLM 집필(Codex/Claude 로그인)
    from oa_typeset.office import ppt, hwp            # 실제 PowerPoint·한글로 마무리(끝줄 정리·PDF)
"""
from .forms import Form, get_form, list_forms  # noqa: F401
from .typeset import md_to_deck, typeset  # noqa: F401

__version__ = "0.1.0"
