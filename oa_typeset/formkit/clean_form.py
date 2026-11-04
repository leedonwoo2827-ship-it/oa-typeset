"""양식 .hwpx 를 '빈 양식'으로 정리한다 — 공개 레포에 올리거나 남에게 줄 때.

    python -m oa_typeset.formkit.clean_form 원본.hwpx forms/<이름>/form.hwpx [--section 1] [--head "머리말"]

남기는 것: 글꼴·문단/글자 모양·스타일(header.xml), 쪽 설정, 바탕쪽(머리말·꼬리말 틀과 그림)
지우는 것: 본문 글(전부), 나머지 구역, 본문에서만 쓰던 그림, 미리보기 글·그림
  --section N : 쪽 설정을 가져올 구역 번호(표지 구역 말고 본문 구역을 고를 때. 기본 0)
  --head      : 바탕쪽의 글자를 이것으로 바꾼다(기본 '머리말'). 원래 책 제목이 남지 않게.
"""
from __future__ import annotations

import argparse
import re
import zipfile

SEC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {ns}>'
       '<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
       '<hp:run charPrIDRef="0">{secpr}</hp:run><hp:run charPrIDRef="0"><hp:t/></hp:run></hp:p></hs:sec>')


def clean(src: str, dst: str, section: int = 0, head: str = "머리말") -> dict:
    z = zipfile.ZipFile(src)
    files = {i.filename: z.read(i.filename) for i in z.infolist()}
    sec = files[f"Contents/section{section}.xml"].decode("utf-8")
    ns = re.search(r"<hs:sec ([^>]*)>", sec).group(1)
    m = re.search(r"<hp:secPr.*?</hp:secPr>", sec, re.S)
    if not m:
        raise SystemExit(f"section{section} 에 쪽 설정(secPr)이 없습니다 — --section 을 바꿔 보세요")
    # 단 설정(colPr)은 양식마다 secPr 앞에도 뒤에도 온다 → 찾아서 secPr 뒤에 표준 순서로 붙인다
    col = re.search(r"<hp:ctrl>\s*<hp:colPr [^>]*/>\s*</hp:ctrl>", sec)
    secpr = m.group(0) + (col.group(0) if col else
                          '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>')
    # 쪽 설정 안의 머리말·꼬리말 컨트롤 말고 다른 개체(표·그림)는 넣지 않는다
    new_sec = SEC.format(ns=ns, secpr=secpr)

    # 바탕쪽 글자 → head
    for name in [n for n in files if re.match(r"Contents/masterpage\d+\.xml$", n)]:
        x = files[name].decode("utf-8")
        x = re.sub(r"<hp:t>[^<]*</hp:t>", f"<hp:t>{head}</hp:t>", x, count=1)
        files[name] = x.encode("utf-8")

    # 쓰는 그림만: header(테두리·배경), 바탕쪽, 새 구역
    used = set()
    for name, data in files.items():
        if name == "Contents/header.xml" or re.match(r"Contents/masterpage\d+\.xml$", name):
            used |= set(re.findall(r'binaryItemIDRef="([^"]+)"', data.decode("utf-8", "replace")))
    used |= set(re.findall(r'binaryItemIDRef="([^"]+)"', new_sec))

    hpf = files["Contents/content.hpf"].decode("utf-8")
    hpf = re.sub(r'<opf:item id="([^"]+)" href="BinData/[^"]*"[^>]*/>',
                 lambda mm: mm.group(0) if mm.group(1) in used else "", hpf)
    hpf = re.sub(r'<opf:item id="section[1-9]\d*"[^>]*/>', "", hpf)
    hpf = re.sub(r'<opf:itemref idref="section[1-9]\d*"[^>]*/>', "", hpf)
    hpf = re.sub(r"<opf:title>.*?</opf:title>", "<opf:title>양식</opf:title>", hpf, flags=re.S)
    hpf = re.sub(r'(<opf:meta name="(?:creator|lastsaveby|subject|description|keyword)"[^>]*>)[^<]*(</opf:meta>)',
                 r"\1\2", hpf)
    rdf = files.get("META-INF/container.rdf", b"").decode("utf-8")
    rdf = re.sub(r'<rdf:Description rdf:about="">\s*<ns0:hasPart [^>]*rdf:resource="Contents/section[1-9]\d*.xml"/>\s*</rdf:Description>'
                 r'<rdf:Description rdf:about="Contents/section[1-9]\d*.xml">.*?</rdf:Description>', "", rdf, flags=re.S)

    kept, dropped = 0, 0
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        out.writestr(zipfile.ZipInfo("mimetype"), b"application/hwp+zip", compress_type=zipfile.ZIP_STORED)
        for name, data in files.items():
            if name == "mimetype" or re.match(r"Contents/section\d+\.xml$", name):
                continue
            if name.startswith("BinData/"):
                if re.sub(r"\.\w+$", "", name.split("/", 1)[1]) not in used:
                    dropped += 1
                    continue
                kept += 1
            if name == "Preview/PrvText.txt":
                data = "양식".encode("utf-8")
            elif name == "Preview/PrvImage.png":
                continue
            elif name == "Contents/header.xml":
                data = re.sub(rb'secCnt="\d+"', b'secCnt="1"', data, count=1)
            elif name == "Contents/content.hpf":
                data = hpf.encode("utf-8")
            elif name == "META-INF/container.rdf":
                data = rdf.encode("utf-8")
            out.writestr(name, data)
        out.writestr("Contents/section0.xml", new_sec.encode("utf-8"))
    return {"out": dst, "images_kept": kept, "images_dropped": dropped}


def main(argv=None):
    ap = argparse.ArgumentParser(description="양식 .hwpx → 본문 없는 빈 양식")
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--section", type=int, default=0)
    ap.add_argument("--head", default="머리말")
    a = ap.parse_args(argv)
    r = clean(a.src, a.dst, a.section, a.head)
    print(f"→ {r['out']}  (그림 {r['images_kept']}개 남김 · {r['images_dropped']}개 뺌)")


if __name__ == "__main__":
    main()
