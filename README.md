# oa-typeset — HTML 원고를 한글·Word·PowerPoint 양식대로 조판

- **무엇**: HTML 원고 하나와 **양식(샘플) 파일**을 넣으면, 그 양식의 글꼴·문단 스타일·쪽 설정·바탕쪽을 그대로 입은 **hwpx · docx · pptx** 가 나옵니다.
- **왜**: 글을 바로 한글 문서로 쓰게 하면 서식이 '바탕글' 하나로 뭉개집니다. **글은 HTML 로 쓰고, 조판은 양식에 맡기면** 출판물 수준으로 나옵니다. 컴활 2급 교재 5회분(문단 1,901 · 표 12 · 그림 42)을 이 방식으로 만들었습니다.
- **어떻게 붙이나**: `pip install` 한 줄, 함수 한 줄. 명령줄·HTTP 로도 부를 수 있습니다.

```bash
pip install "git+https://github.com/leedonwoo2827-ship-it/oa-typeset"
```
```python
from oa_typeset import typeset
typeset("book.html", form="교재형-46배판", formats=["hwpx", "docx", "pptx"], out_dir="out")
```

---

## 목차

1. [젠오피스 결과와 무엇이 다른가](#1-젠오피스-결과와-무엇이-다른가)
2. [5분 안에 돌려 보기](#2-5분-안에-돌려-보기)
3. [젠오피스에 붙이는 세 가지 방법](#3-젠오피스에-붙이는-세-가지-방법)
4. [양식(샘플) — 넣으면 그대로 조판](#4-양식샘플--넣으면-그대로-조판)
5. [HTML 작성 규칙](#5-html-작성-규칙)
6. [Office 마무리 — 실제 PowerPoint·한글로 다듬기](#6-office-마무리--실제-powerpoint한글로-다듬기)
7. [LLM 집필 (선택)](#7-llm-집필-선택)
8. [알려진 한계](#8-알려진-한계)
9. [폴더 구조와 출처](#9-폴더-구조와-출처)

---

## 1. 젠오피스 결과와 무엇이 다른가

같은 종류의 문서(수험 요약집)를 두 방식으로 만든 결과를 hwpx 안을 열어 비교했습니다.

| 항목 | 젠오피스로 바로 생성 (변액보험 요약집) | 이 방식 (컴활 2급 교재) |
|---|---|---|
| 문단 스타일 | **1개** ('바탕글' 하나에 크기만 바꿈) | **49개** (양식의 `*본문` `*1.제목` `*(1)제목` `*표내용` …) |
| 바탕쪽(머리말·쪽 테두리·쪽 번호) | 없음 | 양식 것 그대로 (홀·짝 2종) |
| 판형 | A4 기본 | 양식 판형(46배판) 그대로 |
| 표 | 12 | 12 (머리 행 음영·테두리 양식 규칙) |
| 그림 | 0 | 42 (SVG 는 PNG 로 구워 삽입) |
| 나중에 손보기 | 문단마다 직접 서식 | **스타일만 바꾸면 문서 전체가 바뀜** |

핵심 차이는 **스타일 매핑**입니다. HTML 의 `<h2>` 는 양식의 `*1.제목` 스타일로, `<p>` 는 `*본문` 으로 갑니다. 그래서 편집자가 한글에서 스타일 하나를 고치면 책 전체가 따라옵니다.

## 2. 5분 안에 돌려 보기

Windows 기준입니다. Python 3.10 이상이 필요합니다.

```bat
git clone https://github.com/leedonwoo2827-ship-it/oa-typeset
cd oa-typeset
setup.bat                       :: .venv 를 만들고 설치 (처음 한 번)
```

견본 원고를 두 양식으로 조판해 봅니다.

```bat
.venv\Scripts\activate
oa-typeset build samples\요약집-예시.html --form 교재형-46배판 -f hwpx,docx,pptx -o out
oa-typeset build samples\요약집-예시.html --form 보고서형-A4   -f hwpx,docx,pptx -o out
```

미리 만들어 둔 결과가 [`samples/out/`](samples/out/) 에 있습니다(hwpx · docx · pptx · 한글이 내보낸 pdf).

다른 실행 방법도 있습니다.

- **끌어다 놓기**: `.html` 파일을 `typeset.bat` 위에 놓으면 옆에 세 파일이 생깁니다.
- **웹 화면**: `run-studio.bat` 을 실행하고 http://127.0.0.1:7090 에 접속합니다.

## 3. 젠오피스에 붙이는 세 가지 방법

젠오피스가 어떤 언어로 되어 있든 셋 중 하나로 붙습니다. 모두 **이 PC 안에서만** 동작하고 파일이 밖으로 나가지 않습니다.

### ① Python 에서 import (젠오피스 백엔드가 Python 일 때)

```python
from oa_typeset import typeset, list_forms

# 양식 목록 — forms/ 폴더(기본 2종 + 사용자 폴더)
for f in list_forms():
    print(f.name, f.title)

r = typeset(
    "C:/zen/jobs/123/book.html",     # 원고 (그림은 HTML 기준 상대경로로)
    form="교재형-46배판",              # 양식 이름 · 양식 폴더 경로 · 또는 .hwpx 파일 경로
    formats=["hwpx", "docx", "pptx"],
    out_dir="C:/zen/jobs/123/out",
    forms_dir="C:/zen/forms",        # 회사 양식 폴더 (선택)
    finish=False,                    # True 면 PowerPoint 끝줄 정리 · 한글 PDF (Office 필요)
)
r["hwpx"]["path"]     # → ...\out\book.hwpx
r["hwpx"]["tables"]   # → 표 개수, 그림 개수, 걸린 시간, notes(경고) 등
```

### ② 명령줄 (어떤 언어에서든 서브프로세스로)

`--json` 을 주면 결과가 **stdout 에 JSON 한 줄**로 나옵니다. 종료 코드는 0 성공, 1 실패, 2 사용법 오류입니다.

```bat
oa-typeset --json build book.html --form 교재형-46배판 -f hwpx,docx -o out
```
```json
{"ok": true, "form": "교재형-46배판", "src": "…\\book.html",
 "hwpx": {"path": "…\\out\\book.hwpx", "engine": "html2hwpx", "paragraphs": 30, "tables": 3, "images": 1, "notes": [], "seconds": 1.4},
 "docx": {"path": "…\\out\\book.docx", "notes": [], "seconds": 0.3}}
```
```json
{"ok": false, "error": "양식 '없는양식' 을 찾지 못했습니다. 있는 양식: 교재형-46배판, 보고서형-A4", "type": "FileNotFoundError"}
```

C# 예:

```csharp
var p = Process.Start(new ProcessStartInfo("oa-typeset",
    "--json build \"book.html\" --form 교재형-46배판 -f hwpx -o out")
    { RedirectStandardOutput = true, StandardOutputEncoding = Encoding.UTF8 });
var json = p.StandardOutput.ReadToEnd(); p.WaitForExit();
```

### ③ 로컬 HTTP (서버로 띄워 두고 부를 때)

`oa-typeset studio --port 7090` 으로 띄운 뒤 호출합니다(`pip install "oa-typeset[studio]"` 필요).

```bash
curl -F "file=@book.zip" -F "form=교재형-46배판" -F "formats=hwpx,docx,pptx" http://127.0.0.1:7090/api/typeset
```
```json
{"ok": true, "job": "1104-091530-a1b2c3",
 "files": {"hwpx": "/files/1104-091530-a1b2c3/out/book.hwpx", "docx": "…", "pptx": "…"},
 "result": { …①과 같은 결과… }}
```

- **그림이 있는 원고는 `.zip`** 으로 보냅니다(HTML + 그림 폴더). zip 안의 첫 `.html` 을 원고로 씁니다.
- `GET /api/forms` 는 양식 목록, `GET /files/<job>/<경로>` 는 결과 내려받기입니다.
- 집필(`/api/write`)과 양식 올리기는 이 PC(127.0.0.1)에서만 받습니다.

## 4. 양식(샘플) — 넣으면 그대로 조판

양식 하나는 폴더 하나입니다. **회사 양식을 넣으면 그 모양 그대로 조판됩니다.**

```
forms/
  교재형-46배판/            ← 기본 제공. 컴활 2급 교재를 만든 양식
    form.hwpx               한글 양식: 글꼴·문단 스타일·쪽 설정·바탕쪽을 가져온다
    form.docx               Word 양식: 'Heading 1~3'·'Normal' 등 명명 스타일
    form.pptx               PowerPoint 시안: 배색 15색·글꼴을 가져온다
    form.json               이름·판형·HTML↔스타일 짝·집필 지침
  보고서형-A4/              ← 기본 제공. 계획서·보고서용 A4
```

**새 양식 만들기** — 자세한 내용은 [docs/양식_만들기.md](docs/양식_만들기.md) 에 있습니다.

1. 한글에서 회사 양식을 **다른 이름으로 저장 → HWPX** 로 저장합니다.
2. 본문을 비운 빈 양식으로 정리합니다(남에게 줄 때만 필요):
   `python -m oa_typeset.formkit.clean_form 회사양식.hwpx forms/우리회사/form.hwpx`
3. 스타일 짝을 확인합니다: `oa-typeset styles forms/우리회사/form.hwpx` → 결과의 `suggest` 를 `form.json` 의 `"styles"` 에 넣습니다.

```json
{
  "title": "우리회사 보고서",
  "page": "keep",
  "styles": {"body": "본문", "h1": "제목 1", "h2": "제목 2", "h3": "제목 3", "cell": "표내용"},
  "guide": ["LLM 에게 주는 집필 지침 — 문체, 번호 체계, 표를 쓰는 기준 등"]
}
```

양식을 찾는 순서는 ① `forms_dir` 인자 ② 환경변수 `OA_FORMS_DIR`(`;` 로 여러 개) ③ 현재 폴더의 `forms/` ④ 패키지에 든 기본 양식입니다.

## 5. HTML 작성 규칙

조판기가 옮기는 태그는 정해져 있습니다. 이 규칙대로 쓰면(사람이든 LLM 이든) 양식이 그대로 먹습니다. 전체 규칙은 [docs/HTML_작성규칙.md](docs/HTML_작성규칙.md) 에 있습니다.

| HTML | 한글(hwpx) | Word(docx) | PowerPoint(pptx) |
|---|---|---|---|
| `<h1>` | 양식 `h1` 스타일 (예: `*절제목`) | Heading 1 | 간지(section) |
| `<h2>` `<h3>` `<h4>` | `h2`~`h4` 스타일 | Heading 2~3 | 슬라이드 제목 |
| `<p>` | `body` 스타일 | Normal | 불릿(문장 단위) |
| `<ul>/<ol><li>` | 들여쓴 문단 + `· ` / `1. ` | List Bullet | 불릿 |
| `<table>` `<th>` `<td colspan>` | **진짜 한글 표** (머리 행 음영) | 표 | 표 슬라이드 (8행씩) |
| `<img src="x.png">` / `.svg` | 그림 (SVG 는 PNG 로 구움) | 그림 | — |
| `<b>` · `<small>` | 굵게 · 작은 회색 | 굵게 | — |
| `style="break-before:page"` | 쪽 나눔 | — | — |

## 6. Office 마무리 — 실제 PowerPoint·한글로 다듬기

파일 생성 자체는 Office 없이 됩니다. 하지만 **줄이 어디서 접히는지는 PowerPoint·한글이 글꼴을 실제로 놓아 봐야** 압니다. 그래서 마무리는 Office 를 자동으로 띄워서 합니다(Windows · `pip install "oa-typeset[office]"`).

```bat
oa-typeset build book.html -f hwpx,pptx --finish
oa-typeset finish 발표.pptx --png 미리보기폴더
oa-typeset finish 교재.hwpx
```

| 대상 | 하는 일 |
|---|---|
| PowerPoint | **끝줄 정리**: 문단의 마지막 줄이 3자 이하이거나 윗줄의 20% 이하이면, 자간을 0.25pt 씩 좁혀(최대 −1pt) 윗줄로 끌어올립니다. 안 되면 원래대로 두고 보고합니다. **넘침 확인**: 글상자보다 글이 길면 몇 장인지 알려 줍니다(자르지 않음). 슬라이드 PNG 도 내보냅니다. |
| 한글 | 한글로 열어 **다시 저장**합니다(줄·쪽 배치가 확정된 완성본이 됨). 그리고 **PDF**·**쪽수**(70쪽 목표 확인용)를 냅니다. |
| 조판 단계 (Office 없이) | 모든 문단에 **외톨이줄 보호**, 제목에 **다음 문단과 함께**를 켭니다. |

PowerPoint 는 **고친 곳이 있을 때만** 저장합니다. PowerPoint 가 저장하면 임베드 글꼴을 통째로 다시 넣어 파일이 3~4배 커지기 때문입니다(실측 1.9MB → 7.6MB).

> **한글 보안 승인 창** — 자동화가 파일을 열 때 "접근 허용" 창이 뜨면 **[모두 허용]** 을 누릅니다.
> 아예 안 뜨게 하려면 한컴의 `FilePathCheckerModule.dll` 을 레지스트리
> `HKCU\Software\HNC\HwpAutomation\Modules` 에 `FilePathCheckerModule` 이라는 이름(문자열 값 = dll 경로)으로 등록합니다.

## 7. LLM 집필 (선택)

원고가 없을 때 자료(PDF·hwp·hwpx·docx·xlsx·pptx)를 주면 LLM 이 **목차 → 장별 HTML** 을 씁니다. **API 키는 쓰지 않습니다.** 이 PC 에 로그인한 CLI 에 맡깁니다.

| 공급자 | 로그인 | 비고 |
|---|---|---|
| `codex` (기본) | `codex login` — ChatGPT 계정 | `codex exec` 비대화 모드 |
| `claude` (보조) | `claude auth login` — Claude 구독 | 사용 한도에 걸리면 재설정 시각까지 기다렸다가 이어 씀 |
| `auto` | — | codex 가 실패하면 claude 로 |

```bat
oa-typeset llm                                   :: 로그인 상태
oa-typeset write 자료.pdf 기출.hwpx --brief "변액보험 판매관리사 요약집, 수험생 대상" ^
           --form 교재형-46배판 --pages 70 --llm codex -o work\변액 --build hwpx,docx
```

- 장마다 `work\변액\chapters\chNN.html` 로 저장합니다. 중간에 멈춰도 다시 실행하면 끝난 장은 건너뜁니다.
- 목차는 `outline.json` 에 남습니다. 고쳐 두고 다시 실행하면 고친 목차로 씁니다.
- 장 하나만 다시 쓰려면 `--only 3 --redo` 를 줍니다. 장 파일을 손으로 고쳤으면 `python -c "from oa_typeset.writer import assemble; assemble('work/변액')"` 로 `book.html` 만 다시 묶습니다.
- 발표자료는 `build --llm codex` 를 주면 LLM 이 원고를 13종 레이아웃 덱으로 재구성합니다. 주지 않으면 제목 구조로 자동 구성합니다.

## 8. 알려진 한계

- **줄·쪽 배치는 한글이 열 때 계산합니다.** 생성 직후 파일에는 배치 정보가 없습니다. 다른 뷰어에서 똑같이 보여야 하면 `--finish`(한글로 다시 저장)를 씁니다.
- 표의 **세로 병합(rowspan)은 지원하지 않습니다.** 가로 병합(colspan)만 됩니다.
- **SVG 그림**은 Edge/Chrome 을 헤드리스로 띄워 PNG 로 굽습니다. 브라우저가 없으면 SVG 옆에 같은 이름의 `.png` 를 두어야 합니다.
- **docx** 는 양식의 명명 스타일만 상속합니다(바탕쪽에 해당하는 머리글·바닥글은 form.docx 에 있으면 그대로 남음).
- **pptx** 는 `form.pptx` 에서 **색과 글꼴만** 읽습니다. 좌표·레이아웃은 엔진(deck13)에 고정되어 있습니다.
- Office 마무리는 Windows + 해당 Office + pywin32 에서만 됩니다.

## 9. 폴더 구조와 출처

```
oa_typeset/
  typeset.py        조판 진입점 — typeset(), md_to_deck()
  forms.py          양식 찾기·form.json·스타일 짝 제안
  htmlblocks.py     HTML → Markdown 블록 (docx·pptx 용)
  writer.py         LLM 집필 (목차 → 장별 HTML → book.html)
  cli.py            명령줄
  engines/
    html2hwpx.py    HTML → HWPX (표준 라이브러리만)
    hwpx_md.py      Markdown → HWPX (양식 ZIP 복제)
    docx_gen.py · pptx_gen.py · md_gen.py · mdblocks.py
    deck13/         13종 레이아웃 디자인 덱 (배색·글꼴은 form.pptx 에서)
  office/           ppt.py (끝줄 정리·넘침·PNG) · hwp.py (다시 앉히기·PDF·쪽수)
  llm/              codex(ChatGPT 로그인) · claude(구독 로그인) 호출
  extractors.py     자료 읽기 — pdf · hwp · hwpx · docx · xlsx · pptx
  formkit/          clean_form (회사 양식 → 빈 양식) · make_docx_form
  studio/           로컬 웹 화면 (FastAPI)
  forms/            기본 양식 2종
samples/            견본 원고와 결과
docs/               HTML 작성 규칙 · 양식 만들기 · 젠오피스 연동 메모
tools/              시안(form.pptx) 다시 그리기 등 개발용
```

**출처**

- `engines/html2hwpx.py` 는 사내 해설북 스튜디오(`261103-new-etest-ebook/tools/html2hwpx`)에서 가져왔습니다. 고친 곳마다 `oa_typeset 수정` 주석을 달았습니다.
  - 양식의 구역이 여럿일 때 `secCnt` 를 맞춤 (안 맞추면 한글이 파일을 못 엶)
  - 외톨이줄 보호를 켬
  - 목록에 글머리 기호·번호를 붙임
  - SVG 굽기를 재시도함
- `engines/docx_gen · pptx_gen · hwpx_md · deck13`, `llm/codex_*`, `extractors.py` 는 사내 aim-local 의 문서 생산 스튜디오에서 조판·LLM 호출 부분만 떼어 왔습니다. 사업 관리·대시보드·사업별 레시피는 가져오지 않았습니다.
- `forms/*/form.pptx` 에는 **Noto Sans KR**(SIL Open Font License 1.1) 글꼴이 임베드되어 있습니다.
