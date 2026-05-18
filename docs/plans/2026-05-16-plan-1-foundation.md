# 계획 1 — 기반: 리포 스캐폴드 · 마스터 템플릿 · 슬라이드 카탈로그

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** DEEPNOID PPT Generator 플러그인의 저장소 골격을 세우고, 원본 57장에서 10장을 추출한 마스터 템플릿과 그 슬라이드 카탈로그를 만든다.

**Architecture:** 원본 `.pptx`에서 카탈로그 대상 10장만 남긴 `deepnoid-template.pptx`를 생성한다. 각 슬라이드의 XML을 직접 분석해 편집 가능한 슬롯(텍스트·차트·표)의 위치와 용량을 `slide-catalog.json`에 기록한다. 카탈로그는 이후 계획에서 레이아웃 설계자와 슬라이드 조립가가 사용한다.

**Tech Stack:** Python 3.12 (python-pptx), Node.js (pdf-to-img 렌더링), LibreOffice (`soffice`), git.

**참고 — 설계 문서:** `docs/specs/2026-05-16-deepnoid-ppt-generator-design.md`

**경로 기준:** 모든 경로는 저장소 루트 `deepnoid-ppt-generator/` 기준 상대경로다. 저장소 절대경로: `C:\Users\user\Desktop\Design Skills\deepnoid-ppt-generator`. 원본 자료는 상위 폴더 `C:\Users\user\Desktop\Design Skills\`에 있다.

**Python 실행:** 이 환경의 `python`/`python3` 은 Windows Store 스텁이라 동작하지 않는다. 모든 Python 명령은 전체 경로 `C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe` 로 실행한다. 아래 스텝의 `python ...` 표기는 이 전체 경로로 바꿔 실행할 것.

---

## 파일 구조 (이 계획에서 생성)

| 경로 | 책임 |
|---|---|
| `.claude-plugin/plugin.json` | 플러그인 매니페스트 |
| `LICENSE` | 라이선스 (회사 내부용 — proprietary 표기) |
| `README.md` | 임시 README (정식 한글본은 계획 3) |
| `requirements.txt` | Python 런타임 의존성 |
| `tools/extract_template.py` | 1회성 빌드 스크립트 — 원본 57장 → 10장 추출 |
| `tools/build_slide_map.py` | 추출본의 위치(1~10)↔실제 슬라이드 파트 경로 매핑 생성 |
| `tools/validate_catalog.py` | `slide-catalog.json` 스키마·정합성 검증 스크립트 |
| `tools/render_pptx.mjs` | pptx → 슬라이드 PNG 렌더 (검증·미리보기용) |
| `skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx` | 10장 마스터 템플릿 |
| `skills/generate-deepnoid-ppt/assets/slide-map.json` | 위치↔슬라이드 파트 매핑 |
| `skills/generate-deepnoid-ppt/assets/slide-catalog.json` | 슬라이드 카탈로그 |
| `skills/generate-deepnoid-ppt/assets/catalog-previews/*.png` | 유형별 미리보기 이미지 (10장) |
| `docs/reference/original-template.pptx` | 원본 57장 (레퍼런스 보관) |
| `docs/reference/slides-by-type/` | 원본 57장 유형별 분류 이미지 |
| `docs/reference/xml-notes.md` | 10장 슬라이드 XML 분석 노트 |

---

## Task 1: 저장소 골격 생성

**Files:**
- Create: `.claude-plugin/plugin.json`
- Create: `LICENSE`
- Create: `README.md`

- [ ] **Step 1: 디렉터리 골격 생성**

Run:
```bash
cd "C:/Users/user/Desktop/Design Skills/deepnoid-ppt-generator"
mkdir -p .claude-plugin commands agents tools \
  skills/generate-deepnoid-ppt/assets/catalog-previews \
  skills/generate-deepnoid-ppt/scripts \
  docs/reference/slides-by-type
```

- [ ] **Step 2: 플러그인 매니페스트 작성**

Create `.claude-plugin/plugin.json`:
```json
{
  "name": "deepnoid-ppt-generator",
  "version": "0.1.0",
  "description": "Excel/Word 문서를 DEEPNOID 공식 PPT 템플릿 디자인 그대로 PPTX로 자동 생성하는 플러그인",
  "author": { "name": "via-soi", "email": "mkt.songhee@gmail.com" }
}
```

- [ ] **Step 3: LICENSE 작성**

Create `LICENSE`:
```
Copyright (c) 2026 DEEPNOID Inc. All rights reserved.

이 저장소는 DEEPNOID Inc. 사내용 자료입니다.
DEEPNOID PPT 템플릿과 브랜드 자산이 포함되어 있으며,
DEEPNOID Inc.의 사전 서면 동의 없이 외부 배포·복제·공개를 금합니다.
```

- [ ] **Step 4: 임시 README 작성**

Create `README.md`:
```markdown
# DEEPNOID PPT Generator

Excel/Word 문서를 DEEPNOID 공식 PPT 템플릿 디자인 그대로 PPTX로 자동 생성하는
Claude Code 플러그인입니다.

> 정식 사용 안내(한글)는 개발 완료 시 이 문서에 작성됩니다.
> 설계 문서: `docs/specs/2026-05-16-deepnoid-ppt-generator-design.md`

## 개발 상태

- [ ] 계획 1 — 기반: 리포 골격, 마스터 템플릿, 슬라이드 카탈로그
- [ ] 계획 2 — 파이프라인 스크립트
- [ ] 계획 3 — 스킬·에이전트·명령
```

(계획 1이 끝나는 Task 10에서 `계획 1` 항목을 `[x]` 로 바꾼다. Task 1 시점에는
계획 1이 아직 진행 중이므로 `[ ]` 로 둔다.)

- [ ] **Step 5: plugin.json 유효성 확인**

Run: `node -e "JSON.parse(require('fs').readFileSync('.claude-plugin/plugin.json','utf8')); console.log('plugin.json OK')"`
Expected: `plugin.json OK`

- [ ] **Step 6: 커밋**

```bash
git add .claude-plugin/plugin.json LICENSE README.md
git commit -m "저장소 골격: 플러그인 매니페스트, LICENSE, 임시 README"
```

---

## Task 2: Python 환경 구성

**Files:**
- Create: `requirements.txt`

- [ ] **Step 1: Python 설치 여부 확인**

Run: `python --version 2>&1 || py --version 2>&1 || echo "PYTHON_MISSING"`
Expected: `Python 3.x` 출력. `PYTHON_MISSING` 이면 다음 스텝 진행.

- [ ] **Step 2: Python 미설치 시 설치**

Python이 없으면 winget으로 설치:
```bash
winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
```
winget이 없으면 https://www.python.org/downloads/ 에서 3.12 설치 후 PATH 등록.
설치 후 새 셸에서 `python --version` 으로 재확인.

- [ ] **Step 3: requirements.txt 작성**

Create `requirements.txt`:
```
python-pptx==1.0.2
openpyxl==3.1.5
python-docx==1.1.2
markitdown[pptx]==0.1.1
defusedxml==0.7.1
Pillow==11.0.0
```

- [ ] **Step 4: 의존성 설치**

Run: `python -m pip install -r requirements.txt`
Expected: 모든 패키지 `Successfully installed`.

- [ ] **Step 5: import 검증**

Run:
```bash
python -c "import pptx, openpyxl, docx, defusedxml, PIL; print('deps OK')"
```
Expected: `deps OK`

- [ ] **Step 6: 커밋**

```bash
git add requirements.txt
git commit -m "Python 런타임 의존성 정의 (requirements.txt)"
```

---

## Task 3: 레퍼런스 자료를 저장소에 보관

**Files:**
- Create: `docs/reference/original-template.pptx`
- Create: `docs/reference/slides-by-type/` (분류 이미지 57장)

- [ ] **Step 1: 원본 템플릿 복사**

Run:
```bash
cp "../DEEPNOID 2026 PPT 작성 템플릿.pptx" docs/reference/original-template.pptx
```

- [ ] **Step 2: 분류 이미지 복사**

Run:
```bash
cp -r "../slides_by_type/." docs/reference/slides-by-type/
```

- [ ] **Step 3: 보관 확인**

Run: `ls docs/reference/ && find docs/reference/slides-by-type -name "*.png" | wc -l`
Expected: `original-template.pptx` 와 `slides-by-type/` 표시, PNG 개수 `57`.

- [ ] **Step 4: 커밋**

```bash
git add docs/reference
git commit -m "레퍼런스 보관: 원본 57장 템플릿과 유형별 분류 이미지"
```

---

## Task 4: 마스터 템플릿 추출 (10장)

원본 57장에서 카탈로그 대상 10장만 남긴 `deepnoid-template.pptx`를 만든다.
대상 슬라이드(1-based): **20, 21, 22, 27, 28, 29, 30, 38, 46, 57**.
python-pptx 0-based 인덱스: **19, 20, 21, 26, 27, 28, 29, 37, 45, 56**.

**Files:**
- Create: `tools/extract_template.py`
- Create: `skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx`

- [ ] **Step 1: 추출 스크립트 작성**

Create `tools/extract_template.py`:
```python
"""원본 57장 PPTX에서 카탈로그 대상 10장만 남긴 마스터 템플릿을 생성한다.

추출 후 슬라이드 순서(1-based)와 원본 슬라이드 번호:
  1=원본20 cover     2=원본21 index     3=원본22 divider
  4=원본27 content-1col  5=원본28 content-2col
  6=원본29 content-3col  7=원본30 content-4col
  8=원본38 chart     9=원본46 table     10=원본57 closing
"""
import sys
from pptx import Presentation

SRC = "docs/reference/original-template.pptx"
DST = "skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx"
KEEP_0BASED = [19, 20, 21, 26, 27, 28, 29, 37, 45, 56]


def main():
    prs = Presentation(SRC)
    slides = list(prs.slides._sldIdLst)
    if len(slides) != 57:
        sys.exit(f"원본 슬라이드 수가 57이 아님: {len(slides)}")
    keep = set(KEEP_0BASED)
    for i, sld in enumerate(slides):
        if i not in keep:
            prs.slides._sldIdLst.remove(sld)
    prs.save(DST)
    out = Presentation(DST)
    n = len(out.slides._sldIdLst)
    if n != 10:
        sys.exit(f"추출 결과가 10장이 아님: {n}")
    print(f"추출 완료: {DST} ({n}장)")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 추출 실행**

Run: `python tools/extract_template.py`
Expected: `추출 완료: skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx (10장)`

- [ ] **Step 3: 추출본이 열리는지 확인**

Run:
```bash
python -c "from pptx import Presentation; p=Presentation('skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx'); print('슬라이드 수:', len(p.slides._sldIdLst))"
```
Expected: `슬라이드 수: 10`

- [ ] **Step 4: 커밋**

```bash
git add tools/extract_template.py skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx
git commit -m "마스터 템플릿 추출: 원본 57장에서 카탈로그 10장 추출"
```

---

## Task 5: 추출 템플릿 충실도 검증

추출본 10장이 원본의 해당 슬라이드와 픽셀 단위로 동일하게 렌더되는지 확인한다.

**Files:**
- Create: `tools/render_pptx.mjs`

- [ ] **Step 1: 렌더 스크립트 작성**

Create `tools/render_pptx.mjs`:
```javascript
// pptx → 슬라이드 PNG. 사용법: node tools/render_pptx.mjs <pptx> <출력디렉터리>
import { pdf } from "pdf-to-img";
import { execFileSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

const SOFFICE = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

async function main() {
  const [src, outDir] = process.argv.slice(2);
  if (!src || !outDir) { console.error("사용법: node render_pptx.mjs <pptx> <outDir>"); process.exit(1); }
  await fs.mkdir(outDir, { recursive: true });
  const pdfDir = path.resolve(outDir);
  execFileSync(SOFFICE, ["--headless", "--convert-to", "pdf", "--outdir", pdfDir, path.resolve(src)]);
  const pdfPath = path.join(pdfDir, path.basename(src).replace(/\.pptx$/i, ".pdf"));
  const doc = await pdf(pdfPath, { scale: 2 });
  const pad = String(doc.length).length;
  let i = 1;
  for await (const img of doc) {
    await fs.writeFile(path.join(outDir, `slide-${String(i).padStart(pad, "0")}.png`), img);
    i++;
  }
  await fs.unlink(pdfPath);
  console.log(`렌더 완료: ${i - 1}장 → ${outDir}`);
}
main();
```

- [ ] **Step 2: 렌더 의존성 설치**

Run: `npm init -y >/dev/null 2>&1; npm install pdf-to-img`
Expected: `added N packages`. (`package.json`/`node_modules`는 `.gitignore`에 이미 제외됨 — node_modules만 제외, package.json은 커밋.)

- [ ] **Step 3: 추출 템플릿 렌더**

Run: `node tools/render_pptx.mjs skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx _work/extracted`
Expected: `렌더 완료: 10장 → _work/extracted`

- [ ] **Step 4: 추출본과 원본 분류 이미지 육안 대조**

`_work/extracted/slide-01.png` ~ `slide-10.png` 를 Read 도구로 열어, `docs/reference/slides-by-type/` 의 원본 이미지와 1:1 대조한다:
- slide-01 ↔ 원본 slide-20 (01_Cover)
- slide-02 ↔ 원본 slide-21 (02_Index)
- slide-03 ↔ 원본 slide-22 (03_Divider)
- slide-04 ↔ 원본 slide-27 (08_Table)
- slide-05 ↔ 원본 slide-28 (06_Content)
- slide-06 ↔ 원본 slide-29 (07_Chart)
- slide-07 ↔ 원본 slide-30 (06_Content)
- slide-08 ↔ 원본 slide-38 (07_Chart)
- slide-09 ↔ 원본 slide-46 (08_Table)
- slide-10 ↔ 원본 slide-57 (09_Closing)

Expected: 각 쌍이 디자인·텍스트·색상 동일. 차이가 있으면 Task 4의 `KEEP_0BASED` 인덱스를 점검하고 재실행.

- [ ] **Step 5: package.json 커밋**

```bash
git add tools/render_pptx.mjs package.json
git commit -m "pptx 렌더 스크립트 추가 + 추출 템플릿 충실도 검증 통과"
```

---

## Task 6: 슬라이드 맵 생성 + 10장 XML 분석 (디스커버리)

**중요 — 슬라이드 파트 이름 문제:** Task 4의 추출 스크립트(`python-pptx`의 `_sldIdLst.remove`)는
발표에서 빠진 47장의 `<p:sldId>` 참조만 제거하고, 슬라이드 XML 파트 파일 자체는
패키지에 남긴다. 또한 살아남은 10장의 파트 파일은 **원본 이름(`slide20.xml` 등)을
그대로 유지**한다. 즉 추출본을 풀면 `ppt/slides/` 에 `slide1.xml`~`slide10.xml` 이
아니라 원본 번호의 파일 약 57개가 보인다. 따라서 "발표 순서상 N번째 슬라이드"와
"실제 파트 파일"을 잇는 매핑 파일 `slide-map.json` 을 먼저 만든다. 이후 모든 Task는
이 매핑을 통해 슬라이드를 찾는다. (47개 고아 파트의 제거·정리는 계획 2의 정리
스크립트에서 다룬다 — 계획 1에서는 매핑만으로 충분하다.)

**Files:**
- Create: `tools/build_slide_map.py`
- Create: `skills/generate-deepnoid-ppt/assets/slide-map.json`
- Create: `docs/reference/xml-notes.md`

- [ ] **Step 1: 슬라이드 맵 생성 스크립트 작성**

Create `tools/build_slide_map.py` with EXACTLY this content:
```python
"""추출 마스터 템플릿의 슬라이드를 발표 순서대로 훑어
position(1~10) → 유형 → 실제 슬라이드 파트 경로 매핑을 slide-map.json 으로 출력한다.

python-pptx 의 슬라이드 제거 방식은 원본 파트 이름(slideXX.xml)을 유지하므로,
추출본의 슬라이드 파트는 slide1~slide10.xml 이 아니다. 이 매핑이 위치↔파트의
유일한 신뢰 가능한 연결고리다."""
import json
import sys
from pptx import Presentation

ASSETS = "skills/generate-deepnoid-ppt/assets"
TEMPLATE = f"{ASSETS}/deepnoid-template.pptx"
OUT = f"{ASSETS}/slide-map.json"
TYPES = ["cover", "index", "divider", "content-1col", "content-2col",
         "content-3col", "content-4col", "chart", "table", "closing"]


def main():
    prs = Presentation(TEMPLATE)
    slides = list(prs.slides)
    if len(slides) != 10:
        sys.exit(f"슬라이드 수가 10이 아님: {len(slides)}")
    mapping = []
    for pos, (slide, typ) in enumerate(zip(slides, TYPES), start=1):
        part = str(slide.part.partname).lstrip("/")
        mapping.append({"position": pos, "type": typ, "part": part})
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"slide-map 작성 완료: {OUT}")
    for m in mapping:
        print(f"  {m['position']:2d}  {m['type']:14s}  {m['part']}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 슬라이드 맵 생성 실행**

Run (full python path — bare `python` 은 동작하지 않음): `<PYTHON> tools/build_slide_map.py`
Expected: `slide-map 작성 완료: ...` 와 함께 10줄의 `position 유형 파트경로` 표가 출력된다.
출력된 10개 파트 경로(예: `ppt/slides/slide20.xml`)를 기록해 둔다 — 다음 스텝에서 쓴다.

- [ ] **Step 3: 추출 템플릿 압축 해제**

Run:
```bash
mkdir -p _work/unpacked
cd _work/unpacked && unzip -o "../../skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx" >/dev/null && cd ../..
```
`_work/unpacked/ppt/slides/` 에는 고아 파트를 포함해 약 57개 파일이 있다 — 정상이다.
분석 대상은 오직 `slide-map.json` 에 적힌 10개 파트뿐이다.

- [ ] **Step 4: 각 슬라이드 XML 구조 분석**

`slide-map.json` 의 10개 파트 파일(`_work/unpacked/ppt/slides/<파트>`)을 Read 도구로
읽으며, 각 슬라이드마다 다음을 파악한다:
- 편집 대상 텍스트 도형: `<p:sp>` 의 `<p:nvSpPr><p:cNvPr name="...">` 이름, 안에 든 `<a:t>` 텍스트(플레이스홀더 문구), 글꼴 크기(`sz`).
- 차트 보유 슬라이드(position 6 = content-3col, position 8 = chart): `<p:graphicFrame>` 의 차트 rel(`r:id`) → `ppt/slides/_rels/<파트>.rels` → `ppt/charts/chartN.xml` → 임베디드 데이터(`<c:numCache>`, `<c:strCache>`, 임베디드 xlsx). 차트 종류(`<c:barChart>` 등)·계열 수·카테고리 수와 차트 파트 경로를 기록.
- 표 보유 슬라이드(position 4 = content-1col, position 9 = table): `<a:tbl>` 의 `<a:gridCol>` 수(열), `<a:tr>` 수(행), 각 셀 텍스트.
- 반복 도형 그룹: 콜아웃·카드처럼 같은 구조가 N회 반복되는지(position 5, 7).

- [ ] **Step 5: 분석 노트 작성**

Create `docs/reference/xml-notes.md` — 슬라이드별로 아래 형식으로 기록:
```markdown
# 추출 템플릿 XML 분석 노트

## position 1 — cover (원본 20, 파트 ppt/slides/slideXX.xml)
- 편집 텍스트 도형:
  - cNvPr name="..." — 역할: 메인 타이틀 — 현재 문구: "..." — sz=...
  - cNvPr name="..." — 역할: ... — 현재 문구: "..." — sz=...
- 차트: 없음
- 표: 없음
- 반복 그룹: 없음

## position 2 — index (원본 21, 파트 ...)
... (이하 position 3~10 동일 형식)
```
실제 파트 경로는 `slide-map.json` 값으로 채운다. 모든 도형의 `name` 과 현재 텍스트를
빠짐없이 기록한다 — 다음 Task의 `anchor` 값이 된다.

- [ ] **Step 6: 커밋**

```bash
git add tools/build_slide_map.py skills/generate-deepnoid-ppt/assets/slide-map.json docs/reference/xml-notes.md
git commit -m "슬라이드 맵 생성 + 10장 XML 구조 분석 노트 작성"
```

---

## Task 7: slide-catalog.json 작성

Task 6의 분석 노트를 바탕으로 카탈로그 10개 항목을 작성한다.

**Files:**
- Create: `skills/generate-deepnoid-ppt/assets/slide-catalog.json`

- [ ] **Step 1: 카탈로그 골격 작성**

Create `skills/generate-deepnoid-ppt/assets/slide-catalog.json` — 최상위는 다음 구조:
```json
{
  "templateFile": "deepnoid-template.pptx",
  "sourceDeck": "docs/reference/original-template.pptx",
  "entries": []
}
```

- [ ] **Step 2: 10개 항목 채우기**

`entries` 배열에 아래 10개 항목을 작성한다. 각 항목의 `slots`·`chartSpec`·`tableSpec` 값은 `docs/reference/xml-notes.md` 에서 가져온다.

**`sourceSlide` 의 의미:** 발표 순서상의 위치(1~10)이며 Task 6의 `slide-map.json` 의 `position` 과 같다. 실제 슬라이드 파트 경로는 카탈로그에 넣지 않는다 — 위치↔파트 변환은 `slide-map.json` 이 담당하고, 검증 스크립트(Task 8)가 이를 사용한다.

항목별 고정 메타데이터:
| type | sourceSlide | description |
|---|---|---|
| `cover` | 1 | 발표 표지 — 메인 타이틀 |
| `index` | 2 | 목차 (Contents) |
| `divider` | 3 | 섹션 간지 — 큰 번호 + 섹션 제목 |
| `content-1col` | 4 | 1단 그리드 본문 |
| `content-2col` | 5 | 2단 그리드 — 수치 콜아웃형 |
| `content-3col` | 6 | 3단 그리드 본문 |
| `content-4col` | 7 | 4단 그리드 — 카드형 |
| `chart` | 8 | 내장 막대 차트 슬라이드 |
| `table` | 9 | 표 슬라이드 |
| `closing` | 10 | 마무리 — 문의처 |

각 항목의 형식 (값은 분석 노트 기준 실제 값으로 채움):
```json
{
  "type": "content-2col",
  "sourceSlide": 5,
  "description": "2단 그리드 — 수치 콜아웃형",
  "preview": "catalog-previews/content-2col.png",
  "slots": [
    {
      "slotId": "title",
      "role": "title",
      "anchor": { "shapeName": "<cNvPr name 실제값>" },
      "currentText": "<현재 플레이스홀더 문구>",
      "capacity": { "maxChars": 40, "maxLines": 1 }
    }
  ],
  "chartSpec": null,
  "tableSpec": null
}
```
규칙:
- `anchor.shapeName` — 분석 노트의 `cNvPr name` 실제값.
- `capacity` — 현재 플레이스홀더 문구의 글자 수와 글꼴 크기를 기준으로 산정(현재 문구 길이를 상한으로 잡고 약간의 여유). `maxLines` 는 도형 높이로 판단.
- 반복 도형(콜아웃·카드)은 슬롯에 `"repeatable": true, "min": N, "max": N` 추가하고 `slotId` 를 `stat-1`, `stat-2` 식으로.
- 차트 항목은 `chartSpec`: `{ "chartType": "barChart", "seriesCount": N, "categoryCount": N, "chartPart": "ppt/charts/chartN.xml", "embeddedData": "ppt/embeddings/..." }`.
- 표 항목은 `tableSpec`: `{ "rows": N, "cols": N, "headerRow": true }`.

- [ ] **Step 3: JSON 문법 확인**

Run: `python -c "import json; d=json.load(open('skills/generate-deepnoid-ppt/assets/slide-catalog.json',encoding='utf-8')); print('항목 수:', len(d['entries']))"`
Expected: `항목 수: 10`

- [ ] **Step 4: 커밋**

```bash
git add skills/generate-deepnoid-ppt/assets/slide-catalog.json
git commit -m "슬라이드 카탈로그 작성: 10개 템플릿 유형 정의"
```

---

## Task 8: 카탈로그 검증 스크립트

`slide-catalog.json` 이 스키마를 만족하고 실제 템플릿과 정합한지 자동 검증한다.

**Files:**
- Create: `tools/validate_catalog.py`

- [ ] **Step 1: 검증 스크립트 작성**

Create `tools/validate_catalog.py`:
```python
"""slide-catalog.json 의 구조와 템플릿 정합성을 검증한다.

위치(sourceSlide)→실제 슬라이드 파트 변환은 slide-map.json 으로 한다."""
import json
import sys
import zipfile

ASSETS = "skills/generate-deepnoid-ppt/assets"
CATALOG = f"{ASSETS}/slide-catalog.json"
SLIDE_MAP = f"{ASSETS}/slide-map.json"
TEMPLATE = f"{ASSETS}/deepnoid-template.pptx"
EXPECTED_TYPES = ["cover", "index", "divider", "content-1col", "content-2col",
                  "content-3col", "content-4col", "chart", "table", "closing"]


def fail(msg):
    print(f"검증 실패: {msg}")
    sys.exit(1)


def main():
    catalog = json.load(open(CATALOG, encoding="utf-8"))
    entries = catalog.get("entries", [])
    if len(entries) != 10:
        fail(f"항목 수가 10이 아님: {len(entries)}")

    types = [e["type"] for e in entries]
    if sorted(types) != sorted(EXPECTED_TYPES):
        fail(f"유형 집합 불일치: {types}")

    positions = sorted(e["sourceSlide"] for e in entries)
    if positions != list(range(1, 11)):
        fail(f"sourceSlide 가 1~10 이 아님: {positions}")

    slide_map = json.load(open(SLIDE_MAP, encoding="utf-8"))
    if len(slide_map) != 10:
        fail(f"slide-map 항목 수가 10이 아님: {len(slide_map)}")
    pos_to_part = {m["position"]: m["part"] for m in slide_map}
    pos_to_type = {m["position"]: m["type"] for m in slide_map}

    zf = zipfile.ZipFile(TEMPLATE)
    names = set(zf.namelist())

    for e in entries:
        pos = e["sourceSlide"]
        if pos not in pos_to_part:
            fail(f"{e['type']}: sourceSlide {pos} 가 slide-map 에 없음")
        if pos_to_type[pos] != e["type"]:
            fail(f"sourceSlide {pos}: 카탈로그 유형 '{e['type']}' 와 "
                 f"slide-map 유형 '{pos_to_type[pos]}' 불일치")
        part = pos_to_part[pos]
        if part not in names:
            fail(f"{e['type']}: 슬라이드 파트 '{part}' 가 템플릿에 없음")
        xml = zf.read(part).decode("utf-8")
        if not e.get("slots"):
            fail(f"{e['type']}: slots 가 비어 있음")
        for slot in e["slots"]:
            for key in ("slotId", "role", "anchor", "capacity"):
                if key not in slot:
                    fail(f"{e['type']}/{slot.get('slotId','?')}: '{key}' 누락")
            name = slot["anchor"].get("shapeName")
            if not name:
                fail(f"{e['type']}/{slot['slotId']}: anchor.shapeName 누락")
            if f'name="{name}"' not in xml:
                fail(f"{e['type']}/{slot['slotId']}: "
                     f"도형 '{name}' 이 {part} 에 없음")
        if e["type"] == "chart" and not e.get("chartSpec"):
            fail("chart: chartSpec 누락")
        if e["type"] == "table" and not e.get("tableSpec"):
            fail("table: tableSpec 누락")

    print("검증 통과: 10개 항목, 모든 anchor 가 템플릿과 정합")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 검증 실행**

Run: `python tools/validate_catalog.py`
Expected: `검증 통과: 10개 항목, 모든 anchor 가 템플릿과 정합`
실패 시 메시지가 가리키는 항목을 `slide-catalog.json` 또는 `xml-notes.md` 에서 수정하고 재실행.

- [ ] **Step 3: 커밋**

```bash
git add tools/validate_catalog.py
git commit -m "카탈로그 검증 스크립트 추가 + 검증 통과"
```

---

## Task 9: 카탈로그 미리보기 이미지 생성

레이아웃 설계자가 유형을 고를 때 참고할 유형별 미리보기 PNG 10장을 만든다.

**Files:**
- Create: `skills/generate-deepnoid-ppt/assets/catalog-previews/*.png` (10장)

- [ ] **Step 1: 추출 템플릿을 슬라이드별로 렌더**

Run: `node tools/render_pptx.mjs skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx _work/previews`
Expected: `렌더 완료: 10장 → _work/previews`

- [ ] **Step 2: 유형 이름으로 복사**

Run:
```bash
cd skills/generate-deepnoid-ppt/assets
P=../../../_work/previews
cp $P/slide-01.png catalog-previews/cover.png
cp $P/slide-02.png catalog-previews/index.png
cp $P/slide-03.png catalog-previews/divider.png
cp $P/slide-04.png catalog-previews/content-1col.png
cp $P/slide-05.png catalog-previews/content-2col.png
cp $P/slide-06.png catalog-previews/content-3col.png
cp $P/slide-07.png catalog-previews/content-4col.png
cp $P/slide-08.png catalog-previews/chart.png
cp $P/slide-09.png catalog-previews/table.png
cp $P/slide-10.png catalog-previews/closing.png
cd ../../..
```

- [ ] **Step 3: 미리보기 개수 확인**

Run: `ls skills/generate-deepnoid-ppt/assets/catalog-previews/ | wc -l`
Expected: `10`

- [ ] **Step 4: 커밋**

```bash
git add skills/generate-deepnoid-ppt/assets/catalog-previews
git commit -m "카탈로그 유형별 미리보기 이미지 10장 추가"
```

---

## Task 10: 계획 1 마무리

- [ ] **Step 1: 작업 산출물 정리**

Run: `rm -rf _work && echo "임시 작업물 정리 완료"`
Expected: `임시 작업물 정리 완료` (`_work/` 는 `.gitignore` 에 있어 커밋엔 영향 없음)

- [ ] **Step 2: 전체 구조 확인**

Run: `git ls-files | sort`
Expected: 아래 파일들이 모두 포함되어야 함:
- `.claude-plugin/plugin.json`, `LICENSE`, `README.md`, `requirements.txt`, `package.json`, `.gitignore`
- `tools/extract_template.py`, `tools/build_slide_map.py`, `tools/validate_catalog.py`, `tools/render_pptx.mjs`
- `skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx`
- `skills/generate-deepnoid-ppt/assets/slide-map.json`
- `skills/generate-deepnoid-ppt/assets/slide-catalog.json`
- `skills/generate-deepnoid-ppt/assets/catalog-previews/` PNG 10장
- `docs/specs/2026-05-16-deepnoid-ppt-generator-design.md`
- `docs/plans/2026-05-16-plan-1-foundation.md`
- `docs/reference/original-template.pptx`, `docs/reference/xml-notes.md`, `docs/reference/slides-by-type/` PNG 57장

- [ ] **Step 3: 최종 검증 일괄 실행**

Run (`python` 은 헤더의 전체 경로로 실행):
```bash
python tools/validate_catalog.py && \
python -c "from pptx import Presentation; assert len(Presentation('skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx').slides._sldIdLst)==10; print('템플릿 10장 OK')"
```
Expected: `검증 통과: ...` 와 `템플릿 10장 OK` 둘 다 출력.

- [ ] **Step 4: README 개발 상태 갱신 후 커밋**

`README.md` 의 `- [ ] 계획 1` 을 `- [x] 계획 1` 로 수정한 뒤:
```bash
git add README.md
git commit -m "계획 1 완료: 기반 골격·마스터 템플릿·슬라이드 카탈로그"
```

---

## 완료 기준 (Definition of Done)

- `skills/generate-deepnoid-ppt/assets/deepnoid-template.pptx` 가 10장이며 원본 해당 슬라이드와 동일하게 렌더된다.
- `slide-catalog.json` 이 `tools/validate_catalog.py` 검증을 통과한다.
- 카탈로그 미리보기 PNG 10장이 존재한다.
- 모든 커밋이 완료되고 `git status` 가 깨끗하다.

## 다음 계획

계획 2 — 파이프라인 스크립트: `parse_input.py`(Excel/Word → content.json), pptx 압축·해제·슬라이드 클론·정리, 차트 재바인딩, QA 렌더링 스크립트를 단위 테스트와 함께 작성한다.

**계획 2로 넘기는 기술 부채:** Task 4의 추출본 `deepnoid-template.pptx` 에는 발표에서
빠진 47장의 고아 슬라이드 파트(및 일부 미참조 미디어)가 그대로 남아 파일이 비대하다.
계획 2의 정리(clean) 스크립트로 참조 그래프를 따라 미참조 파트를 제거해, 출하용
마스터 템플릿을 슬림화한다. 계획 1에서는 `slide-map.json` 으로 위치↔파트를 잇는
것으로 충분하므로 정리는 하지 않는다.
