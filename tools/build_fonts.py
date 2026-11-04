# -*- coding: utf-8 -*-
r"""덱에 심을 폰트 만들기 — 가변폰트(VF) → 무게별 정적 TTF + 부분집합.

    python tools/build_fonts.py [원본.ttf] [--family "Noto Sans KR"]

왜 필요한가
-----------
① **파워포인트는 가변폰트를 임베드하지 못한다.** 임베드(`fntdata`)는 정적 폰트를
   전제한다. 윈도우에 깔린 `NotoSansKR-VF.ttf` 를 그대로 심으면 받는 쪽에서
   글꼴이 무시되거나 깨진다. 그래서 wght 400/700 정적 인스턴스로 굽는다.
② **원본이 10MB 라 그대로 심으면 파일이 무겁다.** 한글 음절·라틴·기호만 남기면
   무게당 2.4MB 로 줄어든다.

폰트를 심지 않으면 그 글꼴이 없는 PC 에서 제목·본문·표가 모두 대체 글꼴로 열리고
줄바꿈까지 달라진다 — 납품물에서는 그게 더 큰 문제다.

라이선스: Noto Sans KR 은 OFL 이고 fsType=0(임베드 자유)이다. 다른 글꼴을 쓸 때는
**fsType 을 반드시 확인**할 것(2=임베드 금지). 이 스크립트가 확인하고 거부한다.

만들어진 TTF 는 `oa_typeset/assets/fonts/` 에 놓이고 저장소에는 커밋하지 않는다(.gitignore).
빌드에는 fontTools 가 필요하다: `pip install fonttools`
"""
from __future__ import annotations

import os
import struct
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "oa_typeset", "assets", "fonts")   # deck13 이 여기서 찾는다
DEFAULT_SRC = r"C:\Windows\Fonts\NotoSansKR-VF.ttf"

# 덱에 나올 수 있는 글자만 남긴다 — 한자는 뺀다(들어가면 다시 10MB 가 된다).
def _unicodes():
    u = list(range(0x0020, 0x007F))            # ASCII
    u += list(range(0x00A0, 0x0100))           # 라틴-1 보충
    u += [0x2018, 0x2019, 0x201C, 0x201D, 0x2013, 0x2014, 0x2022, 0x00B7,
          0x2026, 0x00D7, 0x2212, 0x25B8, 0x25CF, 0x2192, 0x2190, 0x2191,
          0x2193, 0x203B, 0x00B0, 0x2030, 0x20A9, 0x20AC, 0x00A3, 0x00A5]
    u += list(range(0x2460, 0x2474))           # ①~⑳
    u += list(range(0x3131, 0x3164))           # 한글 낱자
    u += list(range(0xAC00, 0xD7A4))           # 현대 한글 음절 11,172자
    return u


def _fstype(path: str):
    """임베드 허용 여부(OS/2 fsType). 2 = 임베드 금지."""
    with open(path, "rb") as f:
        d = f.read()
    n = struct.unpack(">H", d[4:6])[0]
    for i in range(n):
        r = 12 + i * 16
        if d[r:r + 4] == b"OS/2":
            off = struct.unpack(">I", d[r + 8:r + 12])[0]
            return struct.unpack(">H", d[off + 8:off + 10])[0]
    return None


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    src = args[0] if args else DEFAULT_SRC
    if not os.path.isfile(src):
        print(f"원본 폰트를 찾지 못했습니다: {src}")
        return 2

    ft = _fstype(src)
    if ft is not None and (ft & 0x0F) == 2:
        print(f"이 폰트는 임베드가 금지되어 있습니다(fsType={ft}). 중단합니다.")
        return 1

    try:
        from fontTools.ttLib import TTFont
        from fontTools.varLib import instancer
        from fontTools import subset
    except ImportError:
        print("fontTools 가 필요합니다:  pip install fonttools")
        return 2

    os.makedirs(OUT_DIR, exist_ok=True)
    uni = _unicodes()
    made = []
    for wght, name in ((400, "NotoSansKR-Regular.ttf"), (700, "NotoSansKR-Bold.ttf")):
        f = TTFont(src)
        if "fvar" in f:
            f = instancer.instantiateVariableFont(
                f, {"wght": wght}, inplace=True, updateFontNames=True)
        opt = subset.Options()
        opt.layout_features = ["*"]
        opt.name_IDs = ["*"]
        opt.notdef_outline = True
        opt.drop_tables += ["DSIG"]
        s = subset.Subsetter(options=opt)
        s.populate(unicodes=uni)
        s.subset(f)
        p = os.path.join(OUT_DIR, name)
        f.save(p)
        made.append((name, os.path.getsize(p), f["name"].getDebugName(1)))

    for name, size, fam in made:
        print(f"  {name:<28} {size / 1048576:5.2f} MB   family={fam}")
    print(f"→ {OUT_DIR}")
    print("이제 덱을 만들면 이 글꼴이 파일 안에 심깁니다(어느 PC 에서나 같은 모양).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
