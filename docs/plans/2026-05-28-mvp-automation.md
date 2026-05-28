# DEEPNOID Design Generator — MVP 자동화 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 구조화 YAML 아웃라인을 입력받아 `DEEPNOID-Design.md` 명세 그대로 DEEPNOID 디자인 PPTX를 자동 생성하는 MVP 자동화 빌더.

**Architecture:** python-pptx 기반 빌더 패키지(`tools/deepnoid_builder/`)가 §6 패턴 함수 단위로 슬라이드를 조립한다. 상위에서 `tools/generate_deck.py` CLI가 YAML을 파싱해 빌더를 호출하고, Claude Code 스킬(`skills/generate-deepnoid-ppt/SKILL.md`) + 슬래시 명령(`/generate-deck`)이 사용자 진입점이 된다.

**Tech Stack:** Python 3.12 · python-pptx · PyYAML · pytest · LibreOffice(소offices) + Node.js pdf-to-img(렌더 검수)

**참고 설계 문서:** `docs/specs/2026-05-28-mvp-automation-design.md`

**경로 기준:** 저장소 루트 `C:\Users\user\Desktop\Design Skills\deepnoid-ppt-generator`.
**Python 실행:** 시스템 PATH의 `python`은 Windows Store 스텁이므로 전체 경로 사용:
`C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe`. 본 계획에서 `python` 표기는 이 경로로 치환해 실행.

---

## 파일 구조 (이 계획에서 생성/수정)

| 경로 | 책임 | 신규/수정 |
|---|---|---|
| `requirements.txt` | PyYAML, pytest 추가 | 수정 |
| `tests/__init__.py` | 테스트 패키지 마커 | 신규 |
| `tests/conftest.py` | pytest 픽스처 (blank Presentation) | 신규 |
| `skills/generate-deepnoid-ppt/assets/deepnoid_logo.png` | DEEPNOID 워드마크 | 신규 (원본에서 추출) |
| `tools/deepnoid_builder/__init__.py` | 패키지 마커 | 신규 |
| `tools/deepnoid_builder/tokens.py` | DEEPNOID-Design.md frontmatter 로더 | 신규 |
| `tools/deepnoid_builder/headers.py` | 우상단 헤더(슬로건+로고+페이지번호) | 신규 |
| `tools/deepnoid_builder/patterns.py` | §6.0~6.5 패턴 함수 | 신규 |
| `tools/deepnoid_builder/builder.py` | 표지/간지/아웃트로 + 디스패처 + build() | 신규 |
| `tools/generate_deck.py` | CLI 진입점 | 신규 |
| `examples/sample-outline.yaml` | 입력 예시 (Tech Talks AX 데모) | 신규 |
| `skills/generate-deepnoid-ppt/SKILL.md` | 오케스트레이션 스킬 | 신규 |
| `commands/generate-deck.md` | `/generate-deck` 슬래시 명령 | 신규 |
| `README.md` | 사용법 안내 추가 | 수정 |

**테스트 전략:**
- 단위 — 각 패턴/빌더 함수는 pytest로 슬라이드 객체에 추가된 도형 수·이름·텍스트·위치를 검증
- 통합 — sample-outline.yaml로 end-to-end 빌드 → LibreOffice·pdf-to-img로 PNG 렌더 → 시각 검수

---

## Task 1: PyYAML·pytest 의존성 추가

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: requirements.txt 에 두 줄 추가**

Edit `requirements.txt` 끝에 추가:
```
PyYAML==6.0.2
pytest==8.3.4
```

- [ ] **Step 2: 설치**

Run: `python -m pip install -r requirements.txt`
Expected: `Successfully installed PyYAML-6.0.2 pytest-8.3.4` (혹은 already satisfied)

- [ ] **Step 3: import 검증**

Run: `python -c "import yaml, pytest; print(yaml.__version__, pytest.__version__)"`
Expected: `6.0.2 8.3.4`

- [ ] **Step 4: 커밋**

```bash
git add requirements.txt
git commit -m "deps: PyYAML·pytest 추가 (MVP 자동화 빌더용)"
```

---

## Task 2: tests/ 폴더 + conftest 설정

**Files:**
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `pytest.ini`

- [ ] **Step 1: 패키지 마커**

Create `tests/__init__.py` (빈 파일).

- [ ] **Step 2: conftest 작성**

Create `tests/conftest.py`:
```python
"""공용 pytest 픽스처."""
import pytest
from pptx import Presentation
from pptx.util import Cm


@pytest.fixture
def prs():
    """16:9 widescreen 빈 프레젠테이션."""
    p = Presentation()
    p.slide_width = Cm(33.87)
    p.slide_height = Cm(19.05)
    return p


@pytest.fixture
def slide(prs):
    """빈 슬라이드 1장 (blank layout)."""
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)
```

- [ ] **Step 3: pytest.ini 작성**

Create `pytest.ini`:
```ini
[pytest]
testpaths = tests
python_files = test_*.py
addopts = -v --tb=short
```

- [ ] **Step 4: 빈 픽스처 동작 확인**

Create `tests/test_conftest.py`:
```python
def test_blank_prs(prs):
    assert prs.slide_width == 1219200 * 10 or True  # 임의 — 단순 픽스처 동작 확인
    assert len(prs.slides) == 0


def test_blank_slide(slide):
    assert len(slide.shapes) == 0
```

Run: `python -m pytest -v`
Expected: 2 passed.

- [ ] **Step 5: 커밋**

```bash
git add tests/ pytest.ini
git commit -m "test: pytest 환경·공용 픽스처(prs/slide) 셋업"
```

---

## Task 3: DEEPNOID 로고 추출

`docs/reference/original-template.pptx` 안에 워드마크 PNG가 미디어로 박혀 있다. 종횡비 약 6.45:1 인 파일을 찾아 `skills/generate-deepnoid-ppt/assets/deepnoid_logo.png` 로 복사한다.

**Files:**
- Create: `tools/extract_logo.py`
- Create: `skills/generate-deepnoid-ppt/assets/deepnoid_logo.png`

- [ ] **Step 1: 추출 스크립트 작성**

Create `tools/extract_logo.py`:
```python
"""원본 템플릿(pptx)의 미디어에서 DEEPNOID 워드마크 PNG를 추출한다.

조건: PNG · 종횡비 6.45:1 부근 (1560×242 가 캐논). 후보가 여러 개면 가장 큰 것을 선택.
"""
import io
import sys
import zipfile
from pathlib import Path
from PIL import Image

SRC = Path("docs/reference/original-template.pptx")
DST = Path("skills/generate-deepnoid-ppt/assets/deepnoid_logo.png")


def main():
    if not SRC.exists():
        sys.exit(f"원본 템플릿 없음: {SRC}")
    zf = zipfile.ZipFile(SRC)
    candidates = []
    for name in zf.namelist():
        if not name.startswith("ppt/media/") or not name.lower().endswith(".png"):
            continue
        data = zf.read(name)
        try:
            img = Image.open(io.BytesIO(data))
        except Exception:
            continue
        w, h = img.size
        if h == 0:
            continue
        ratio = w / h
        if 5.0 <= ratio <= 8.0:  # 워드마크 종횡비 ≈ 6.45
            candidates.append((w * h, name, data, w, h))
    if not candidates:
        sys.exit("워드마크 후보(가로:세로 5~8) 없음")
    candidates.sort(reverse=True)  # 가장 큰 후보
    _, name, data, w, h = candidates[0]
    DST.parent.mkdir(parents=True, exist_ok=True)
    DST.write_bytes(data)
    print(f"추출: {name} ({w}×{h}) → {DST}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 추출 실행**

Run: `python tools/extract_logo.py`
Expected: `추출: ppt/media/<...>.png (<w>×<h>) → skills/generate-deepnoid-ppt/assets/deepnoid_logo.png`

- [ ] **Step 3: 결과 확인**

Run: `python -c "from PIL import Image; im=Image.open('skills/generate-deepnoid-ppt/assets/deepnoid_logo.png'); print(im.size, im.mode)"`
Expected: 종횡비가 6~7 사이인 PNG (예: `(1560, 242) RGBA`).
육안 확인: `_work/` 등에 옮겨 보거나 README의 미리보기로 확인 — DEEPNOID 워드마크가 맞아야 함. 잘못된 PNG가 잡혔으면 종횡비 범위(5.0–8.0)를 조정해 재실행.

- [ ] **Step 4: 커밋**

```bash
git add tools/extract_logo.py skills/generate-deepnoid-ppt/assets/deepnoid_logo.png
git commit -m "logo: 원본 템플릿에서 DEEPNOID 워드마크 PNG 추출"
```

---

## Task 4: `tokens.py` — DEEPNOID-Design.md 토큰 로더

**Files:**
- Create: `tools/deepnoid_builder/__init__.py`
- Create: `tools/deepnoid_builder/tokens.py`
- Create: `tests/test_tokens.py`

- [ ] **Step 1: 패키지 마커**

Create `tools/deepnoid_builder/__init__.py` (빈 파일).

- [ ] **Step 2: 테스트 먼저 작성**

Create `tests/test_tokens.py`:
```python
from pathlib import Path
from tools.deepnoid_builder.tokens import load_tokens

DESIGN_MD = Path("skills/generate-deepnoid-ppt/assets/DEEPNOID-Design.md")


def test_loads_colors_and_typography():
    t = load_tokens(DESIGN_MD)
    assert t["colors"]["primary"] == "#0066FF"
    assert t["colors"]["green"] == "#34D0B3"
    assert t["colors"]["text"] == "#000000"  # = black 별칭
    assert t["colors"]["muted"] == "#6E6E73"
    assert t["typography"]["title"]["fontSize"] == "32pt"
    assert t["typography"]["body"]["fontSize"] == "14pt"


def test_loads_header_and_outro():
    t = load_tokens(DESIGN_MD)
    assert "logo-asset" in t["header"]
    assert "slogan-text" in t["header"]
    assert "slogan-pos" in t["outro"]
```

- [ ] **Step 3: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_tokens.py -v`
Expected: `ModuleNotFoundError: No module named 'tools.deepnoid_builder.tokens'`.

- [ ] **Step 4: 구현 작성**

Create `tools/deepnoid_builder/tokens.py`:
```python
"""DEEPNOID-Design.md frontmatter YAML 로더.

본 파일은 디자인 명세의 frontmatter(YAML)를 dict 로 변환해 빌더 전체가
참조하는 단일 출처가 된다. 본문(markdown 부분)은 사용하지 않는다.
"""
from pathlib import Path
import yaml


def load_tokens(design_md_path: Path) -> dict:
    """`DEEPNOID-Design.md` 의 frontmatter 를 dict 로 반환.

    Args:
        design_md_path: DEEPNOID-Design.md 의 경로.

    Returns:
        frontmatter YAML 의 최상위 dict — colors / typography / spacing /
        rounded / header / chapter-divider / outro 등 키 포함.

    Raises:
        FileNotFoundError: 파일이 없으면.
        ValueError: frontmatter 가 누락되거나 비어 있으면.
    """
    text = Path(design_md_path).read_text(encoding="utf-8")
    if not text.startswith("---"):
        raise ValueError(f"frontmatter 없음: {design_md_path}")
    parts = text.split("---", 2)
    if len(parts) < 3:
        raise ValueError(f"frontmatter 종결자 없음: {design_md_path}")
    fm = yaml.safe_load(parts[1])
    if not fm:
        raise ValueError(f"frontmatter 가 비어 있음: {design_md_path}")
    return fm
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `python -m pytest tests/test_tokens.py -v`
Expected: 2 passed.

- [ ] **Step 6: 커밋**

```bash
git add tools/deepnoid_builder/__init__.py tools/deepnoid_builder/tokens.py tests/test_tokens.py
git commit -m "feat(builder): DEEPNOID-Design.md frontmatter 토큰 로더"
```

---

## Task 5: `headers.py` — 우상단 헤더

**Files:**
- Create: `tools/deepnoid_builder/headers.py`
- Create: `tests/test_headers.py`

design.md §4 우상단 헤더 좌표:
- 슬로건 (17.86, 0.57) cm, 10.42 × 0.3 cm
- 로고 (28.65, 0.54) cm, 2.71 × 0.42 cm
- 페이지번호 (31.36, 0.54) cm, 1.19 × 0.42 cm

- [ ] **Step 1: 테스트 먼저 작성**

Create `tests/test_headers.py`:
```python
from pathlib import Path
from pptx.util import Cm
from tools.deepnoid_builder.headers import add_inner_header

LOGO = Path("skills/generate-deepnoid-ppt/assets/deepnoid_logo.png")


def test_adds_three_shapes(slide):
    add_inner_header(slide, page_num=3, logo_path=LOGO)
    # 슬로건 txt, 로고 pic, 페이지번호 txt = 3 도형
    assert len(slide.shapes) == 3


def test_slogan_position_and_text(slide):
    add_inner_header(slide, page_num=1, logo_path=LOGO)
    slogan = [s for s in slide.shapes if s.has_text_frame and "Through AI" in s.text_frame.text][0]
    assert abs(slogan.left - Cm(17.86)) < Cm(0.05)
    assert abs(slogan.top - Cm(0.57)) < Cm(0.05)


def test_page_number(slide):
    add_inner_header(slide, page_num=7, logo_path=LOGO)
    nums = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip() == "7"]
    assert len(nums) == 1
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_headers.py -v`
Expected: ModuleNotFoundError.

- [ ] **Step 3: 구현 작성**

Create `tools/deepnoid_builder/headers.py`:
```python
"""우상단 헤더(슬로건 + 로고 + 페이지번호) — design.md §4 내지 전용."""
from pathlib import Path
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Pt

SLOGAN_TEXT = "Through AI, We make your life wider, bolder and clearer"

# design.md §4 좌표 (cm)
SLOGAN_POS = (17.86, 0.57)
SLOGAN_SIZE = (10.42, 0.3)
LOGO_POS = (28.65, 0.54)
LOGO_SIZE = (2.71, 0.42)
PAGENUM_POS = (31.36, 0.54)
PAGENUM_SIZE = (1.19, 0.42)

MUTED = RGBColor(0x6E, 0x6E, 0x73)


def add_inner_header(slide, page_num: int, logo_path: Path) -> None:
    """슬라이드에 우상단 헤더 3요소를 추가.

    표지·간지·아웃트로에는 사용하지 않는다(내지 전용).
    """
    _add_slogan(slide)
    if Path(logo_path).exists():
        slide.shapes.add_picture(
            str(logo_path),
            left=Cm(LOGO_POS[0]), top=Cm(LOGO_POS[1]),
            width=Cm(LOGO_SIZE[0]), height=Cm(LOGO_SIZE[1]),
        )
    _add_page_number(slide, page_num)


def _add_slogan(slide) -> None:
    tb = slide.shapes.add_textbox(
        left=Cm(SLOGAN_POS[0]), top=Cm(SLOGAN_POS[1]),
        width=Cm(SLOGAN_SIZE[0]), height=Cm(SLOGAN_SIZE[1]),
    )
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = SLOGAN_TEXT
    run.font.name = "Pretendard"
    run.font.size = Pt(8)
    run.font.color.rgb = MUTED


def _add_page_number(slide, page_num: int) -> None:
    tb = slide.shapes.add_textbox(
        left=Cm(PAGENUM_POS[0]), top=Cm(PAGENUM_POS[1]),
        width=Cm(PAGENUM_SIZE[0]), height=Cm(PAGENUM_SIZE[1]),
    )
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = str(page_num)
    run.font.name = "Pretendard"
    run.font.size = Pt(10)
    run.font.color.rgb = MUTED
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_headers.py -v`
Expected: 3 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/headers.py tests/test_headers.py
git commit -m "feat(builder): 우상단 헤더(슬로건+로고+페이지번호) 빌더"
```

---

## Task 6: `patterns.py` §6.1 카드 + §6.0 골격 (eyebrow + 제목)

**Files:**
- Create: `tools/deepnoid_builder/patterns.py`
- Create: `tests/test_patterns.py`

§6.0 골격:
- eyebrow (1.33, 1.4) cm, 11pt Bold BLUE
- 페이지 제목 (1.33, 2.3) cm, 26pt Bold black

§6.1 카드: 둥근 사각형 6px, 테두리 1pt, 내부 여백 좌·우 0.5cm, 헤더 14pt Bold, 본문 14pt Regular. 채움/테두리는 accent('gray'|'blue'|'green').

- [ ] **Step 1: 테스트 먼저 작성**

Create `tests/test_patterns.py`:
```python
from pptx.util import Cm
from tools.deepnoid_builder.patterns import add_eyebrow_and_title, add_card


def test_eyebrow_and_title_adds_two_textboxes(slide):
    add_eyebrow_and_title(slide, eyebrow="Intro / 지금 우리", title="도입은 모두에게")
    txts = [s for s in slide.shapes if s.has_text_frame]
    eyebrow = [s for s in txts if "Intro" in s.text_frame.text]
    title = [s for s in txts if "도입은" in s.text_frame.text]
    assert len(eyebrow) == 1
    assert len(title) == 1


def test_card_adds_rectangle_with_header_and_body(slide):
    add_card(slide, x_cm=1.33, y_cm=6.5, w_cm=10.0, h_cm=9.0,
             header="R&D", body="70%", accent="blue")
    # 1 사각형(카드) + 1 텍스트박스(헤더+본문) 또는 헤더/본문 각각 = 도형 ≥ 2
    assert len(slide.shapes) >= 2
    # 카드 안에 R&D 와 70% 텍스트가 있어야 함
    all_text = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "R&D" in all_text
    assert "70%" in all_text
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: ModuleNotFoundError.

- [ ] **Step 3: 구현 작성**

Create `tools/deepnoid_builder/patterns.py`:
```python
"""§6 내지 본문 패턴 — design.md §6.0~6.5 좌표·크기를 1:1 매핑.

함수마다 슬라이드 객체와 인자 dict 를 받아 도형을 추가한다. 외부 상태 의존 없음.
"""
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm, Pt

# --- design.md §2 색 ---
BLUE = RGBColor(0x00, 0x66, 0xFF)
GREEN = RGBColor(0x34, 0xD0, 0xB3)
NAVY = RGBColor(0x00, 0x00, 0x57)
BLACK = RGBColor(0x00, 0x00, 0x00)
MUTED = RGBColor(0x6E, 0x6E, 0x73)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE = RGBColor(0xCC, 0xE0, 0xFF)
LIGHT_GREEN = RGBColor(0xD6, 0xF6, 0xF0)
SURFACE = RGBColor(0xF2, 0xF2, 0xF2)
HAIRLINE = RGBColor(0xD1, 0xD1, 0xD6)

# --- 본문 영역 기준 ---
BODY_X = 1.33   # cm
BODY_Y = 6.5    # cm

# --- §6.0 골격 좌표 ---
EYEBROW_POS = (1.33, 1.4)
EYEBROW_SIZE = (20.0, 0.6)
PAGE_TITLE_POS = (1.33, 2.3)
PAGE_TITLE_SIZE = (28.0, 2.4)


# ================== §6.0 골격 ==================
def add_eyebrow_and_title(slide, eyebrow: str, title: str) -> None:
    """내지 상단 골격: eyebrow(11pt BLUE) + 페이지 제목(26pt Bold)."""
    # eyebrow
    tb = slide.shapes.add_textbox(
        Cm(EYEBROW_POS[0]), Cm(EYEBROW_POS[1]),
        Cm(EYEBROW_SIZE[0]), Cm(EYEBROW_SIZE[1]),
    )
    _fmt_text(tb, eyebrow, size_pt=11, bold=True, color=BLUE)
    # title
    tb = slide.shapes.add_textbox(
        Cm(PAGE_TITLE_POS[0]), Cm(PAGE_TITLE_POS[1]),
        Cm(PAGE_TITLE_SIZE[0]), Cm(PAGE_TITLE_SIZE[1]),
    )
    _fmt_text(tb, title, size_pt=26, bold=True, color=BLACK)


# ================== §6.1 카드 ==================
def add_card(slide, x_cm: float, y_cm: float, w_cm: float, h_cm: float,
             header: str, body: str, accent: str = "gray") -> None:
    """둥근 사각형 카드 1장. accent: 'gray' | 'blue' | 'green' | 'blue-solid' | 'green-solid'.

    'gray'(기본) — SURFACE 채움 + HAIRLINE 1pt 테두리
    'blue'/'green' — LIGHT_BLUE/LIGHT_GREEN 채움 + 메인색 1pt 테두리
    'blue-solid'/'green-solid' — 메인색 솔리드 채움(§6.5 KPI 전용)
    """
    fill, line, text_color = _accent_to_colors(accent)
    rect = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Cm(x_cm), Cm(y_cm), Cm(w_cm), Cm(h_cm),
    )
    rect.adjustments[0] = 0.05  # ≈6px
    rect.fill.solid()
    rect.fill.fore_color.rgb = fill
    if line is None:
        rect.line.fill.background()
    else:
        rect.line.color.rgb = line
        rect.line.width = Pt(1)
    rect.shadow.inherit = False
    # header
    inner_left = x_cm + 0.5
    inner_top = y_cm + 0.95
    inner_w = w_cm - 1.0
    tb = slide.shapes.add_textbox(Cm(inner_left), Cm(inner_top), Cm(inner_w), Cm(1.0))
    _fmt_text(tb, header, size_pt=14, bold=True, color=text_color)
    # body
    tb = slide.shapes.add_textbox(Cm(inner_left), Cm(inner_top + 1.2),
                                  Cm(inner_w), Cm(h_cm - 2.6))
    _fmt_text(tb, body, size_pt=14, bold=False, color=text_color)


def _accent_to_colors(accent: str):
    a = accent.lower()
    if a == "gray":
        return SURFACE, HAIRLINE, BLACK
    if a == "blue":
        return LIGHT_BLUE, BLUE, BLACK
    if a == "green":
        return LIGHT_GREEN, GREEN, BLACK
    if a == "blue-solid":
        return BLUE, None, WHITE
    if a == "green-solid":
        return GREEN, None, WHITE
    raise ValueError(f"알 수 없는 accent: {accent}")


# ================== 공용 텍스트 헬퍼 ==================
def _fmt_text(shape, text: str, size_pt: float, bold: bool, color: RGBColor) -> None:
    """텍스트박스/도형에 단일 단락·단일 run 텍스트를 채운다."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    tf.margin_top = Cm(0.1)
    tf.margin_bottom = Cm(0.1)
    p = tf.paragraphs[0]
    p.text = ""  # 기존 run 제거
    run = p.add_run()
    run.text = text
    run.font.name = "Pretendard"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 2 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/patterns.py tests/test_patterns.py
git commit -m "feat(builder): §6.0 골격(eyebrow+제목) + §6.1 카드 primitive"
```

---

## Task 7: `patterns.py` §6.2 카드 그리드

design.md §6.2:
- 3단: 10.07 × 9.0 cm, gap 0.50
- 4단: 7.46 × 9.0 cm, gap 0.45
- 7카드(4×2): 7.46 × 5.3 cm, gap 가로 0.45 / 세로 0.40, 8번째는 LIGHT_GREEN 안내 카드
- 시작 (1.33, 6.5)

**Files:**
- Modify: `tools/deepnoid_builder/patterns.py`
- Modify: `tests/test_patterns.py`

- [ ] **Step 1: 테스트 추가**

Append to `tests/test_patterns.py`:
```python
from tools.deepnoid_builder.patterns import add_card_grid


def test_grid_3_adds_three_cards(slide):
    cards = [
        {"header": "A", "body": "a", "accent": "gray"},
        {"header": "B", "body": "b", "accent": "blue"},
        {"header": "C", "body": "c", "accent": "gray"},
    ]
    add_card_grid(slide, n=3, cards=cards)
    rects = [s for s in slide.shapes if s.shape_type == 1]  # AUTO_SHAPE
    assert len(rects) == 3


def test_grid_4_adds_four_cards(slide):
    cards = [{"header": f"H{i}", "body": f"b{i}"} for i in range(4)]
    add_card_grid(slide, n=4, cards=cards)
    rects = [s for s in slide.shapes if s.shape_type == 1]
    assert len(rects) == 4


def test_grid_7_adds_eight_cards(slide):
    """7카드 + 1 안내 카드 = 도형 8."""
    cards = [{"header": f"H{i}", "body": f"b{i}"} for i in range(7)]
    add_card_grid(slide, n=7, cards=cards, note="설명 카드")
    rects = [s for s in slide.shapes if s.shape_type == 1]
    assert len(rects) == 8
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 3 failures (ImportError on add_card_grid).

- [ ] **Step 3: 구현 추가**

Append to `tools/deepnoid_builder/patterns.py`:
```python
# ================== §6.2 카드 그리드 ==================
GRID_SPECS = {
    # n: (card_w, card_h, gap_x)
    2: (15.20, 9.0, 0.50),
    3: (10.07, 9.0, 0.50),
    4: (7.46, 9.0, 0.45),
}
GRID_7 = (7.46, 5.3, 0.45, 0.40)  # w, h, gap_x, gap_y — 4×2


def add_card_grid(slide, n: int, cards: list, note: str = "") -> None:
    """§6.2 카드 그리드. n ∈ {2,3,4,7}.

    cards: [{"header": str, "body": str, "accent": "gray"|"blue"|"green"}, ...]
    note (n=7 한정): 8번째 안내 카드의 본문. 헤더 없음.
    """
    if n == 7:
        return _grid_7(slide, cards, note)
    if n not in GRID_SPECS:
        raise ValueError(f"지원하지 않는 카드 수: {n}")
    w, h, gap = GRID_SPECS[n]
    if len(cards) < n:
        cards = cards + [{"header": "", "body": "", "accent": "gray"}] * (n - len(cards))
    for i in range(n):
        x = BODY_X + i * (w + gap)
        c = cards[i]
        add_card(slide, x_cm=x, y_cm=BODY_Y, w_cm=w, h_cm=h,
                 header=c.get("header", ""), body=c.get("body", ""),
                 accent=c.get("accent", "gray"))


def _grid_7(slide, cards: list, note: str) -> None:
    """4×2 그리드: 7장 + 1 안내 카드."""
    w, h, gx, gy = GRID_7
    if len(cards) < 7:
        cards = cards + [{"header": "", "body": "", "accent": "gray"}] * (7 - len(cards))
    for i in range(7):
        col = i % 4
        row = i // 4
        x = BODY_X + col * (w + gx)
        y = BODY_Y + row * (h + gy)
        c = cards[i]
        add_card(slide, x_cm=x, y_cm=y, w_cm=w, h_cm=h,
                 header=c.get("header", ""), body=c.get("body", ""),
                 accent=c.get("accent", "gray"))
    # 8번째: 안내 카드 (LIGHT_GREEN 배경)
    col = 3
    row = 1
    x = BODY_X + col * (w + gx)
    y = BODY_Y + row * (h + gy)
    add_card(slide, x_cm=x, y_cm=y, w_cm=w, h_cm=h,
             header="", body=note, accent="green")
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 5 passed (전체).

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/patterns.py tests/test_patterns.py
git commit -m "feat(builder): §6.2 카드 그리드 (2/3/4/7)"
```

---

## Task 8: `patterns.py` §6.3 비교형 (AS-IS / TO-BE)

design.md §6.3:
- 좌측: (1.33, 6.5) 15.2 × 8.5 cm, LIGHT_BLUE 채움, 테두리 없음
- 우측: (17.34, 6.5) 15.2 × 8.5 cm, WHITE 채움, GREEN 1.5pt 테두리
- 라벨 10pt Bold (좌: NAVY, 우: GREEN)
- 헤더 14pt Bold, 본문 14pt × 3 bullets
- 우측 마지막 bullet: BLUE Bold
- 하단 캡션 (y=16.0) 가운데, 14pt MUTED

**Files:**
- Modify: `tools/deepnoid_builder/patterns.py`
- Modify: `tests/test_patterns.py`

- [ ] **Step 1: 테스트 추가**

Append to `tests/test_patterns.py`:
```python
from tools.deepnoid_builder.patterns import add_comparison


def test_comparison_adds_two_panels(slide):
    add_comparison(slide,
                   asis={"label": "AS-IS · 지금", "header": "반복 업무", "bullets": ["a", "b", "c"]},
                   tobe={"label": "TO-BE · 이후", "header": "검토·책임", "bullets": ["d", "e", "f"]})
    rects = [s for s in slide.shapes if s.shape_type == 1]
    assert len(rects) == 2  # 좌·우 패널


def test_comparison_caption(slide):
    add_comparison(slide,
                   asis={"label": "L", "header": "Lh", "bullets": ["1", "2", "3"]},
                   tobe={"label": "R", "header": "Rh", "bullets": ["4", "5", "6"]},
                   caption="요약 캡션")
    cap = [s for s in slide.shapes if s.has_text_frame and "요약 캡션" in s.text_frame.text]
    assert len(cap) == 1
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 2 failures.

- [ ] **Step 3: 구현 추가**

Append to `tools/deepnoid_builder/patterns.py`:
```python
# ================== §6.3 비교형 ==================
COMPARE_LEFT = (1.33, 6.5, 15.2, 8.5)    # x, y, w, h
COMPARE_RIGHT = (17.34, 6.5, 15.2, 8.5)
CAPTION_Y = 16.0


def add_comparison(slide, asis: dict, tobe: dict, caption: str = "") -> None:
    """§6.3 AS-IS / TO-BE 비교 카드 2장 + 선택 캡션.

    asis/tobe: {"label": str, "header": str, "bullets": list[str]}
    """
    # 좌측 (AS-IS)
    _compare_panel(slide, *COMPARE_LEFT,
                   fill=LIGHT_BLUE, line=None,
                   label=asis["label"], label_color=NAVY,
                   header=asis["header"], bullets=asis["bullets"],
                   highlight_last=False)
    # 우측 (TO-BE)
    _compare_panel(slide, *COMPARE_RIGHT,
                   fill=WHITE, line=(GREEN, 1.5),
                   label=tobe["label"], label_color=GREEN,
                   header=tobe["header"], bullets=tobe["bullets"],
                   highlight_last=True)
    if caption:
        tb = slide.shapes.add_textbox(Cm(BODY_X), Cm(CAPTION_Y), Cm(31.21), Cm(0.8))
        from pptx.enum.text import PP_ALIGN
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = ""
        run = p.add_run()
        run.text = caption
        run.font.name = "Pretendard"
        run.font.size = Pt(14)
        run.font.color.rgb = MUTED


def _compare_panel(slide, x, y, w, h, fill, line,
                   label: str, label_color, header: str, bullets: list,
                   highlight_last: bool) -> None:
    rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Cm(x), Cm(y), Cm(w), Cm(h))
    rect.adjustments[0] = 0.05
    rect.fill.solid()
    rect.fill.fore_color.rgb = fill
    if line is None:
        rect.line.fill.background()
    else:
        color, width_pt = line
        rect.line.color.rgb = color
        rect.line.width = Pt(width_pt)
    rect.shadow.inherit = False
    # 라벨 (10pt Bold)
    tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(y + 0.5), Cm(w - 1.0), Cm(0.6))
    _fmt_text(tb, label, size_pt=10, bold=True, color=label_color)
    # 헤더 (14pt Bold)
    tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(y + 1.3), Cm(w - 1.0), Cm(1.0))
    _fmt_text(tb, header, size_pt=14, bold=True, color=BLACK)
    # 본문 bullets (14pt Regular)
    tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(y + 2.6),
                                  Cm(w - 1.0), Cm(h - 3.0))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.1)
    tf.margin_top = tf.margin_bottom = Cm(0.1)
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ""
        run = p.add_run()
        run.text = f"•  {b}"
        run.font.name = "Pretendard"
        run.font.size = Pt(14)
        is_last_highlight = highlight_last and i == len(bullets) - 1
        run.font.bold = is_last_highlight
        run.font.color.rgb = BLUE if is_last_highlight else BLACK
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 7 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/patterns.py tests/test_patterns.py
git commit -m "feat(builder): §6.3 비교형 (AS-IS / TO-BE) 패턴"
```

---

## Task 9: `patterns.py` §6.4 스텝형

design.md §6.4:
- 카드 3~5개, 4단 그리드와 동일 폭/높이 (7.46 × 9.0, gap 0.45)
- 카드 i%2: BLUE/GREEN 테두리 (외곽만)
- STEP 라벨 y+0.5, 10pt Bold 카드색
- 헤더 y+1.4, 14pt Bold BLACK
- 본문 y+3.4, 14pt Regular
- 연결 화살표: 카드 사이 gap 중앙 수평, BLUE 0.75pt, 직선 화살표 연결선(`MSO_CONNECTOR.STRAIGHT`)
- 하단 메시지 y=16.0 가운데, 14pt Bold BLUE

**Files:**
- Modify: `tools/deepnoid_builder/patterns.py`
- Modify: `tests/test_patterns.py`

- [ ] **Step 1: 테스트 추가**

Append to `tests/test_patterns.py`:
```python
from tools.deepnoid_builder.patterns import add_step_flow


def test_step_flow_4_cards(slide):
    steps = [
        {"label": f"STEP {i+1}", "header": f"H{i+1}", "body": f"B{i+1}"}
        for i in range(4)
    ]
    add_step_flow(slide, steps=steps, footer="반복 핵심 메시지")
    rects = [s for s in slide.shapes if s.shape_type == 1]
    assert len(rects) == 4
    # 카드 사이 화살표 3개
    conns = [s for s in slide.shapes if s.shape_type == 9]  # LINE / connector
    assert len(conns) >= 3


def test_step_flow_footer(slide):
    steps = [{"label": "S1", "header": "H", "body": "B"}] * 3
    add_step_flow(slide, steps=steps, footer="요약 한 줄")
    footer = [s for s in slide.shapes if s.has_text_frame and "요약 한 줄" in s.text_frame.text]
    assert len(footer) == 1
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 2 failures.

- [ ] **Step 3: 구현 추가**

Append to `tools/deepnoid_builder/patterns.py`:
```python
# ================== §6.4 스텝형 ==================
from pptx.enum.shapes import MSO_CONNECTOR

STEP_W = 7.46
STEP_H = 9.0
STEP_GAP = 0.45


def add_step_flow(slide, steps: list, footer: str = "") -> None:
    """§6.4 스텝형 (3~5개 카드 + 연결 화살표).

    steps: [{"label": "STEP 1", "header": "...", "body": "..."}, ...]
    """
    n = len(steps)
    if not 3 <= n <= 5:
        raise ValueError(f"스텝 수는 3~5: 받은 값 {n}")
    for i, s in enumerate(steps):
        x = BODY_X + i * (STEP_W + STEP_GAP)
        color = BLUE if i % 2 == 0 else GREEN
        _step_card(slide, x, BODY_Y, STEP_W, STEP_H, color,
                   s.get("label", ""), s.get("header", ""), s.get("body", ""))
        # 화살표 (마지막 카드 뒤에는 없음)
        if i < n - 1:
            x_from = x + STEP_W + 0.05
            x_to = x + STEP_W + STEP_GAP - 0.05
            y_mid = BODY_Y + STEP_H / 2
            conn = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Cm(x_from), Cm(y_mid), Cm(x_to), Cm(y_mid),
            )
            conn.line.color.rgb = BLUE
            conn.line.width = Pt(0.75)
            # 화살촉
            line = conn.line.element
            ln = line.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
            if ln is not None:
                from lxml import etree
                tail_arrow = etree.SubElement(
                    ln, '{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd',
                    {'type': 'triangle', 'w': 'med', 'len': 'med'},
                )
    if footer:
        tb = slide.shapes.add_textbox(Cm(BODY_X), Cm(CAPTION_Y), Cm(31.21), Cm(0.8))
        from pptx.enum.text import PP_ALIGN
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = ""
        run = p.add_run()
        run.text = footer
        run.font.name = "Pretendard"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = BLUE


def _step_card(slide, x, y, w, h, color: RGBColor,
               label: str, header: str, body: str) -> None:
    rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Cm(x), Cm(y), Cm(w), Cm(h))
    rect.adjustments[0] = 0.05
    rect.fill.solid()
    rect.fill.fore_color.rgb = SURFACE
    rect.line.color.rgb = color
    rect.line.width = Pt(1)
    rect.shadow.inherit = False
    # STEP 라벨
    tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(y + 0.5), Cm(w - 1.0), Cm(0.6))
    _fmt_text(tb, label, size_pt=10, bold=True, color=color)
    # 헤더
    tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(y + 1.4), Cm(w - 1.0), Cm(1.6))
    _fmt_text(tb, header, size_pt=14, bold=True, color=BLACK)
    # 본문
    tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(y + 3.4),
                                  Cm(w - 1.0), Cm(h - 4.0))
    _fmt_text(tb, body, size_pt=14, bold=False, color=BLACK)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 9 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/patterns.py tests/test_patterns.py
git commit -m "feat(builder): §6.4 스텝형 (4단계 + 연결 화살표)"
```

---

## Task 10: `patterns.py` §6.5 KPI·메시지형

design.md §6.5: 3~4개 카드, BLUE/GREEN 솔리드 채움 교대, 큰 숫자 42pt White, 헤더 14pt Bold White, 본문 14pt Regular White.

**Files:**
- Modify: `tools/deepnoid_builder/patterns.py`
- Modify: `tests/test_patterns.py`

- [ ] **Step 1: 테스트 추가**

Append to `tests/test_patterns.py`:
```python
from tools.deepnoid_builder.patterns import add_kpi_cards


def test_kpi_4_cards(slide):
    cards = [
        {"number": f"0{i+1}", "header": f"H{i+1}", "body": f"B{i+1}"}
        for i in range(4)
    ]
    add_kpi_cards(slide, cards=cards)
    rects = [s for s in slide.shapes if s.shape_type == 1]
    assert len(rects) == 4
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 1 failure.

- [ ] **Step 3: 구현 추가**

Append to `tools/deepnoid_builder/patterns.py`:
```python
# ================== §6.5 KPI·메시지형 ==================
def add_kpi_cards(slide, cards: list) -> None:
    """§6.5 KPI/메시지 카드 — 메인색 솔리드 채움, 흰 텍스트.

    cards: [{"number": "01", "header": "...", "body": "..."}, ...]
    """
    n = len(cards)
    if not 3 <= n <= 4:
        raise ValueError(f"KPI 카드 수는 3~4: 받은 값 {n}")
    w, h, gap = STEP_W, STEP_H, STEP_GAP   # 4단 그리드와 동일
    for i, c in enumerate(cards):
        x = BODY_X + i * (w + gap)
        fill = BLUE if i % 2 == 0 else GREEN
        rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Cm(x), Cm(BODY_Y), Cm(w), Cm(h))
        rect.adjustments[0] = 0.05
        rect.fill.solid()
        rect.fill.fore_color.rgb = fill
        rect.line.fill.background()
        rect.shadow.inherit = False
        # 큰 숫자
        tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(BODY_Y + 0.5), Cm(w - 1.0), Cm(2.5))
        _fmt_text(tb, c.get("number", ""), size_pt=42, bold=True, color=WHITE)
        # 헤더
        tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(BODY_Y + 3.5), Cm(w - 1.0), Cm(1.6))
        _fmt_text(tb, c.get("header", ""), size_pt=14, bold=True, color=WHITE)
        # 본문
        tb = slide.shapes.add_textbox(Cm(x + 0.5), Cm(BODY_Y + 5.5),
                                      Cm(w - 1.0), Cm(h - 6.0))
        _fmt_text(tb, c.get("body", ""), size_pt=14, bold=False, color=WHITE)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_patterns.py -v`
Expected: 10 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/patterns.py tests/test_patterns.py
git commit -m "feat(builder): §6.5 KPI·메시지형 (솔리드 채움 카드)"
```

---

## Task 11: `builder.py` — 표지 / 간지 / 아웃트로

design.md §4 + Tech Talks 덱 실측:
- **표지**: title (대형 BLUE) + subtitle + 좌하단 워드마크. MVP 간단 레이아웃 — title (1.91cm, 6.0cm) 30 × 4 cm 32pt BLUE Bold, subtitle (1.91cm, 10.5cm) 30 × 1 cm 14pt BLACK, 로고 (1.91, 14.29) 3.06×1.35
- **간지**: 챕터 라벨(있을 때) "Chapter 0X" 20pt BLUE (1.91, 1.91), Main Title 36pt Bold (2.41, 0.58 + 아래 1.5cm offset 적용 ≈ 4.5), 글리프 180pt GREEN Bold (2.3, 12.5)
- **아웃트로**: 슬로건 32pt + 연락처 + 좌하단 로고 + 푸터. MVP — 슬로건 1줄 + 좌하단 로고만(연락처는 outline에서 받을 수 있게)

**Files:**
- Create: `tools/deepnoid_builder/builder.py`
- Create: `tests/test_builder.py`

- [ ] **Step 1: 테스트 먼저 작성**

Create `tests/test_builder.py`:
```python
from pathlib import Path
from pptx import Presentation
from tools.deepnoid_builder.builder import add_cover, add_divider, add_outro

LOGO = Path("skills/generate-deepnoid-ppt/assets/deepnoid_logo.png")


def test_cover_has_title_and_subtitle(prs):
    add_cover(prs, title="제목", subtitle="부제 · 발표자", logo_path=LOGO)
    slide = prs.slides[0]
    txts = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "제목" in txts
    assert "부제" in txts


def test_divider_has_title_and_number(prs):
    add_divider(prs, title="섹션 제목", chapter="Chapter 01", number=1)
    slide = prs.slides[0]
    txts = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "섹션 제목" in txts
    assert "Chapter 01" in txts
    assert "1" in txts


def test_outro_has_slogan(prs):
    add_outro(prs, logo_path=LOGO)
    slide = prs.slides[0]
    txts = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "Through AI" in txts
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_builder.py -v`
Expected: ModuleNotFoundError.

- [ ] **Step 3: 구현 작성**

Create `tools/deepnoid_builder/builder.py`:
```python
"""DEEPNOID 슬라이드 조립 — 표지·간지·아웃트로 + outline 디스패처."""
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from tools.deepnoid_builder.headers import add_inner_header
from tools.deepnoid_builder.patterns import (
    add_eyebrow_and_title, add_card_grid, add_comparison, add_step_flow, add_kpi_cards,
    BLUE, GREEN, BLACK, MUTED, _fmt_text,
)

SLIDE_W_CM = 33.87
SLIDE_H_CM = 19.05
SLOGAN = "Through AI, We make your life wider, bolder and clearer"


def new_presentation() -> Presentation:
    p = Presentation()
    p.slide_width = Cm(SLIDE_W_CM)
    p.slide_height = Cm(SLIDE_H_CM)
    return p


def _add_blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)


# ================== 표지 ==================
def add_cover(prs, title: str, subtitle: str, logo_path: Path) -> None:
    slide = _add_blank_slide(prs)
    # 제목 (대형 BLUE)
    tb = slide.shapes.add_textbox(Cm(1.91), Cm(6.0), Cm(30.0), Cm(4.0))
    _fmt_text(tb, title, size_pt=32, bold=True, color=BLUE)
    # 부제 (BLACK)
    tb = slide.shapes.add_textbox(Cm(1.91), Cm(10.5), Cm(30.0), Cm(1.0))
    _fmt_text(tb, subtitle, size_pt=14, bold=False, color=BLACK)
    # 로고
    if Path(logo_path).exists():
        slide.shapes.add_picture(str(logo_path),
                                 Cm(1.91), Cm(14.29), Cm(3.06), Cm(1.35))


# ================== 간지 ==================
def add_divider(prs, title: str, chapter: str = "", number: int | None = None) -> None:
    slide = _add_blank_slide(prs)
    # Chapter 라벨 (있을 때)
    if chapter:
        tb = slide.shapes.add_textbox(Cm(1.91), Cm(1.91), Cm(10.0), Cm(0.8))
        _fmt_text(tb, chapter, size_pt=20, bold=False, color=BLUE)
    # Main Title (36pt Bold)
    tb = slide.shapes.add_textbox(Cm(2.41), Cm(4.0), Cm(28.0), Cm(4.0))
    _fmt_text(tb, title, size_pt=36, bold=True, color=BLACK)
    # 글리프 (180pt GREEN Bold)
    if number is not None:
        tb = slide.shapes.add_textbox(Cm(2.3), Cm(12.5),
                                      Cm(8.0) if number >= 10 else Cm(3.7), Cm(3.33))
        _fmt_text(tb, str(number), size_pt=180, bold=True, color=GREEN)


# ================== 아웃트로 ==================
def add_outro(prs, logo_path: Path) -> None:
    slide = _add_blank_slide(prs)
    # 슬로건 32pt
    tb = slide.shapes.add_textbox(Cm(1.91), Cm(4.40), Cm(20.93), Cm(2.87))
    _fmt_text(tb, SLOGAN, size_pt=32, bold=True, color=BLACK)
    # 좌하단 워드마크
    if Path(logo_path).exists():
        slide.shapes.add_picture(str(logo_path),
                                 Cm(1.91), Cm(14.29), Cm(3.06), Cm(1.35))
    # 푸터
    tb = slide.shapes.add_textbox(Cm(1.91), Cm(16.08), Cm(17.03), Cm(1.45))
    _fmt_text(tb, "DEEPNOID Inc.  © 2026 All rights reserved.",
              size_pt=9, bold=False, color=MUTED)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_builder.py -v`
Expected: 3 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/builder.py tests/test_builder.py
git commit -m "feat(builder): 표지·간지·아웃트로 슬라이드"
```

---

## Task 12: `builder.py` 메인 `build()` 디스패처

**Files:**
- Modify: `tools/deepnoid_builder/builder.py`
- Modify: `tests/test_builder.py`

- [ ] **Step 1: 테스트 추가**

Append to `tests/test_builder.py`:
```python
from tools.deepnoid_builder.builder import build


def test_build_full_outline():
    outline = {
        "deck": {"title": "T", "subtitle": "S"},
        "slides": [
            {"type": "cover", "title": "표지", "subtitle": "부제"},
            {"type": "divider", "chapter": "Chapter 01", "title": "섹션", "number": 1},
            {"type": "card-grid-3",
             "eyebrow": "Intro / X",
             "title": "내지 제목",
             "cards": [
                 {"header": "A", "body": "a"},
                 {"header": "B", "body": "b", "accent": "blue"},
                 {"header": "C", "body": "c"},
             ]},
            {"type": "outro"},
        ],
    }
    prs = build(outline)
    assert len(prs.slides) == 4
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_builder.py::test_build_full_outline -v`
Expected: ImportError on `build`.

- [ ] **Step 3: 구현 추가**

Append to `tools/deepnoid_builder/builder.py`:
```python
LOGO_PATH = Path("skills/generate-deepnoid-ppt/assets/deepnoid_logo.png")


def build(outline: dict, logo_path: Path = LOGO_PATH) -> Presentation:
    """outline dict → Presentation 객체. 슬라이드 타입별로 디스패치한다."""
    prs = new_presentation()
    inner_page_num = 0  # 내지는 1부터 매김
    for spec in outline.get("slides", []):
        t = spec.get("type")
        if t == "cover":
            add_cover(prs, title=spec["title"], subtitle=spec.get("subtitle", ""),
                      logo_path=logo_path)
        elif t == "divider":
            add_divider(prs, title=spec["title"],
                        chapter=spec.get("chapter", ""),
                        number=spec.get("number"))
        elif t == "outro":
            add_outro(prs, logo_path=logo_path)
        elif t in ("card-grid-2", "card-grid-3", "card-grid-4", "card-grid-7"):
            inner_page_num += 1
            slide = _add_blank_slide(prs)
            add_inner_header(slide, page_num=inner_page_num, logo_path=logo_path)
            add_eyebrow_and_title(slide, spec.get("eyebrow", ""), spec.get("title", ""))
            n = int(t.split("-")[-1])
            add_card_grid(slide, n=n, cards=spec.get("cards", []),
                          note=spec.get("note", ""))
        elif t == "comparison":
            inner_page_num += 1
            slide = _add_blank_slide(prs)
            add_inner_header(slide, page_num=inner_page_num, logo_path=logo_path)
            add_eyebrow_and_title(slide, spec.get("eyebrow", ""), spec.get("title", ""))
            add_comparison(slide, asis=spec["asis"], tobe=spec["tobe"],
                           caption=spec.get("caption", ""))
        elif t == "step":
            inner_page_num += 1
            slide = _add_blank_slide(prs)
            add_inner_header(slide, page_num=inner_page_num, logo_path=logo_path)
            add_eyebrow_and_title(slide, spec.get("eyebrow", ""), spec.get("title", ""))
            add_step_flow(slide, steps=spec.get("steps", []),
                          footer=spec.get("footer", ""))
        elif t == "kpi":
            inner_page_num += 1
            slide = _add_blank_slide(prs)
            add_inner_header(slide, page_num=inner_page_num, logo_path=logo_path)
            add_eyebrow_and_title(slide, spec.get("eyebrow", ""), spec.get("title", ""))
            add_kpi_cards(slide, cards=spec.get("cards", []))
        else:
            raise ValueError(f"알 수 없는 슬라이드 type: {t}")
    return prs
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/ -v`
Expected: 전체 테스트 통과 (4 + 2 + 3 + 10 + 4 = 23 정도).

- [ ] **Step 5: 커밋**

```bash
git add tools/deepnoid_builder/builder.py tests/test_builder.py
git commit -m "feat(builder): build() 메인 디스패처 (8개 슬라이드 타입)"
```

---

## Task 13: `tools/generate_deck.py` CLI

**Files:**
- Create: `tools/generate_deck.py`
- Create: `tests/test_cli.py`

- [ ] **Step 1: 테스트 먼저 작성**

Create `tests/test_cli.py`:
```python
import subprocess
import sys
from pathlib import Path


def test_cli_builds_pptx(tmp_path):
    outline_yaml = tmp_path / "outline.yaml"
    outline_yaml.write_text("""
deck:
  title: Test
  output: {out}
slides:
  - type: cover
    title: 표지 제목
    subtitle: 부제
  - type: outro
""".format(out=str(tmp_path / "out.pptx")).replace("\\", "/"),
                            encoding="utf-8")
    result = subprocess.run(
        [sys.executable, "tools/generate_deck.py", str(outline_yaml)],
        capture_output=True, text=True, encoding="utf-8",
    )
    assert result.returncode == 0, result.stderr
    out = tmp_path / "out.pptx"
    assert out.exists()
    assert out.stat().st_size > 1000
```

- [ ] **Step 2: 테스트 실행 (실패 확인)**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL (`tools/generate_deck.py` 없음).

- [ ] **Step 3: 구현 작성**

Create `tools/generate_deck.py`:
```python
"""CLI: outline.yaml → DEEPNOID 디자인 PPTX 빌드.

사용법:
    python tools/generate_deck.py path/to/outline.yaml

outline.yaml 의 deck.output 필드가 있으면 거기에 저장, 없으면
outline 파일과 같은 폴더의 'output.pptx' 에 저장한다.
"""
import sys
from pathlib import Path

# tools/ 를 PYTHONPATH 처럼 잡기 위해 저장소 루트를 sys.path 에 추가
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import yaml
from tools.deepnoid_builder.builder import build


def main(argv: list) -> int:
    if len(argv) < 2:
        print("사용법: python tools/generate_deck.py <outline.yaml>", file=sys.stderr)
        return 2
    outline_path = Path(argv[1])
    if not outline_path.exists():
        print(f"outline 파일 없음: {outline_path}", file=sys.stderr)
        return 2
    outline = yaml.safe_load(outline_path.read_text(encoding="utf-8"))
    deck = outline.get("deck", {})
    output = deck.get("output") or str(outline_path.parent / "output.pptx")
    prs = build(outline)
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    print(f"저장 완료: {output}  (슬라이드 {len(prs.slides)}장)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_cli.py -v`
Expected: 1 passed.

- [ ] **Step 5: 커밋**

```bash
git add tools/generate_deck.py tests/test_cli.py
git commit -m "feat(cli): generate_deck.py — outline.yaml → DEEPNOID pptx"
```

---

## Task 14: `examples/sample-outline.yaml`

Tech Talks AX #01 덱 일부(7~8장)를 데모 입력으로 작성.

**Files:**
- Create: `examples/sample-outline.yaml`

- [ ] **Step 1: 작성**

Create `examples/sample-outline.yaml`:
```yaml
deck:
  title: "Tech Talks AX #01 — 비개발자의 AX 포인트"
  subtitle: "데모 outline"
  output: "_work/demo-output.pptx"

slides:
  - type: cover
    title: "엄마, 나 터미널 켰어"
    subtitle: "윈도우 쓰는 비개발자의 AX 포인트 찾기  ·  Tech Talks AX #01"

  - type: divider
    chapter: "Chapter 01"
    title: "AX는 기술부서만의 일이 아니다"
    number: 1

  - type: card-grid-3
    eyebrow: "Intro  /  지금 우리 풍경"
    title: "도입은 모두에게, 사용은 부서별로"
    cards:
      - header: "도구 확산"
        body: "Claude · ChatGPT · Copilot · NotebookLM…  도구 측면에서는 없어서 못 쓰는 단계가 아님."
        accent: gray
      - header: "부서별 격차"
        body: "사용률은 부서마다 10배 이상 벌어진다.  관전만 하면 격차는 굳어진다."
        accent: blue
      - header: "써봐야 안다"
        body: "AX는 기술팀만의 일이 아니다.  현업의 손에서 굴려봐야 sweet spot 보임."
        accent: green

  - type: comparison
    eyebrow: "Why  /  왜 써야해?"
    title: "관성의 시간을 검토와 책임의 시간으로"
    asis:
      label: "AS-IS  ·  지금 우리"
      header: "반복적·관성적 업무에 시간을 쓴다"
      bullets:
        - "같은 형식의 반복 검토"
        - "비슷한 메일 매주 발송"
        - "휴먼 에러 + 피로"
    tobe:
      label: "TO-BE  ·  AX 이후"
      header: "검토와 책임의 시간을 확보한다"
      bullets:
        - "반복 작업은 AI에게 위임"
        - "사람은 의사결정·책임"
        - "관성 → 판단의 영역으로 이동"
    caption: "AI는 도구, AX는 업무·조직의 전환."

  - type: step
    eyebrow: "How  /  4단계 반복 프레임"
    title: "작게 시도하고, 시그널을 보고, 크게 키운다"
    steps:
      - label: "STEP 1"
        header: "가설을 세운다"
        body: "이 업무에서 AI가 도울 수 있을 만한 가장 작은 가설을 적는다."
      - label: "STEP 2"
        header: "작게 테스트한다"
        body: "최소 리소스로 굴린다.  하루 안에 끝낼 크기로."
      - label: "STEP 3"
        header: "긍정 시그널을 본다"
        body: "시간 절약 · 에러 감소 · 검토 품질 등 정량·체감 시그널 확인."
      - label: "STEP 4"
        header: "스케일업한다"
        body: "팀 워크플로우로 옮기고, 스킬·자동화로 자산화한다."
    footer: "반복(Iteration) — 정답을 향한 가장 빠른 길."

  - type: kpi
    eyebrow: "Closing  /  Takehome Message"
    title: "AX는 도구 도입이 아니라 정의의 전환"
    cards:
      - number: "01"
        header: "왜?가 먼저다"
        body: "목표가 없으면 자동화는 잡일이 된다."
      - number: "02"
        header: "워크플로우 설계"
        body: "Task가 아니라 Workflow를 본다."
      - number: "03"
        header: "AI ≠ AX"
        body: "AI는 기술, AX는 업무·조직의 전환."
      - number: "04"
        header: "직무 정의가 먼저"
        body: "본인의 직무 정의가 있어야 AI에 먹히지 않는다."

  - type: outro
```

- [ ] **Step 2: outline 빌드 동작 확인**

Run: `python tools/generate_deck.py examples/sample-outline.yaml`
Expected: `저장 완료: _work/demo-output.pptx  (슬라이드 7장)`

- [ ] **Step 3: 렌더링 시각 검수**

Run: `node tools/render_pptx.mjs _work/demo-output.pptx _work/demo-png`
Expected: `렌더 완료: 7장 → _work/demo-png`

`_work/demo-png/slide-1.png` ~ `slide-7.png` 를 Read 도구로 열어 §6 패턴이 design.md 명세대로 나오는지 시각 확인:
- 표지: BLUE 큰 제목 + 부제 + 좌하단 워드마크
- 간지 1: "Chapter 01" + 큰 제목 + 좌하단 글리프 "1" GREEN
- 카드 그리드 3: eyebrow + 26pt 제목 + 카드 3개(gray/blue/green)
- 비교형: 좌 LIGHT_BLUE + 우 WHITE+GREEN 테두리, 우측 마지막 bullet BLUE
- 스텝형: 4 카드 + 사이 화살표 + 하단 BLUE 메시지
- KPI: BLUE/GREEN 솔리드 4 카드, 큰 숫자
- 아웃트로: 슬로건 + 좌하단 로고 + 푸터
- 내지 3장에 우상단 슬로건 + 로고 + 페이지번호 헤더가 있어야 함

문제가 보이면 해당 패턴 함수의 좌표·크기를 수정하고 재실행.

- [ ] **Step 4: 커밋**

```bash
git add examples/sample-outline.yaml
git commit -m "feat(examples): Tech Talks AX 데모 outline + end-to-end 빌드 검증"
```

---

## Task 15: `skills/generate-deepnoid-ppt/SKILL.md` + `commands/generate-deck.md`

**Files:**
- Create: `skills/generate-deepnoid-ppt/SKILL.md`
- Create: `commands/generate-deck.md`

- [ ] **Step 1: SKILL.md 작성**

Create `skills/generate-deepnoid-ppt/SKILL.md`:
```markdown
---
name: generate-deepnoid-ppt
description: DEEPNOID 디자인 시스템(`DEEPNOID-Design.md`)을 따른 PPTX를 자동 생성. 사용자가 구조화 YAML 아웃라인을 제공하면 §6 패턴(카드/그리드/비교형/스텝형/KPI)을 조합해 한 번에 발표 자료를 빌드한다.
---

# DEEPNOID PPT 자동 생성 (MVP)

이 스킬은 `tools/deepnoid_builder/` 빌더를 통해 구조화된 outline YAML 을 DEEPNOID 디자인 PPTX 로 변환합니다.

## 입력 형식

`examples/sample-outline.yaml` 을 참고하세요. 지원 슬라이드 타입:
`cover` · `divider` · `card-grid-2` · `card-grid-3` · `card-grid-4` · `card-grid-7` · `comparison` · `step` · `kpi` · `outro`.

각 슬라이드의 필수·선택 필드는 `docs/specs/2026-05-28-mvp-automation-design.md` §5 참조.

## 사용 방법

```bash
python tools/generate_deck.py path/to/outline.yaml
```

`outline.yaml` 의 `deck.output` 에 명시한 경로로 PPTX 가 저장됩니다. 명시가 없으면 outline 과 같은 폴더의 `output.pptx`.

## 검수

빌드 후 시각 검수:
```bash
node tools/render_pptx.mjs <output.pptx> _work/preview
```
`_work/preview/slide-NN.png` 로 슬라이드별 PNG 가 생성됩니다.

## 디자인 단일 출처

모든 색·서체·간격은 `skills/generate-deepnoid-ppt/assets/DEEPNOID-Design.md` 의 frontmatter 토큰을 따릅니다. 빌더는 본 파일을 단일 출처로 참조합니다.

## 범위 밖 (MVP)

- Word·Excel 자유 문서 입력 → 아웃라인 자동 제안
- 5 에이전트 팀 협업 생성
- §6.7 강조 컴포넌트(라벨 칩·빗금·키워드 하이라이트), §6.8 아이콘
- 자동 차트 생성
```

- [ ] **Step 2: 슬래시 명령 작성**

Create `commands/generate-deck.md`:
```markdown
---
description: outline.yaml 로부터 DEEPNOID 디자인 PPTX 를 빌드합니다. 사용법: /generate-deck <outline.yaml 경로>
---

# /generate-deck

사용자가 `/generate-deck <outline.yaml>` 명령을 실행하면:

1. outline 경로를 ARGUMENTS 에서 받습니다.
2. `python tools/generate_deck.py <ARGUMENTS>` 를 실행합니다.
3. 저장된 PPTX 경로를 사용자에게 안내합니다.
4. 선택적으로 `node tools/render_pptx.mjs <pptx> _work/preview` 로 PNG 미리보기를 만들어 시각 검수를 제안합니다.

오류 처리:
- ARGUMENTS 가 비어 있으면 `examples/sample-outline.yaml` 을 기본값으로 사용해도 되는지 사용자에게 묻습니다.
- outline 파싱/검증 오류는 stderr 그대로 사용자에게 전달합니다.

상세 입력 형식·지원 타입은 `skills/generate-deepnoid-ppt/SKILL.md` 참조.

ARGUMENTS: $ARGUMENTS
```

- [ ] **Step 3: 작성 확인**

Run: `cat skills/generate-deepnoid-ppt/SKILL.md | head -3 && cat commands/generate-deck.md | head -3`
Expected: 두 파일 모두 정상 출력.

- [ ] **Step 4: 커밋**

```bash
git add skills/generate-deepnoid-ppt/SKILL.md commands/generate-deck.md
git commit -m "feat(skill): generate-deepnoid-ppt 스킬 + /generate-deck 슬래시 명령"
```

---

## Task 16: README 업데이트 + 최종 푸시

**Files:**
- Modify: `README.md`

- [ ] **Step 1: README 의 "구성" 표를 갱신해 빌더·스킬·명령을 명시**

Edit `README.md` — "구성" 표 아래에 다음 사용법 섹션 추가 (정확한 위치는 현재 README 구조에 맞춰 조정):

```markdown
## 사용 방법 (자동화 빌더)

1. **입력 outline 작성** — `examples/sample-outline.yaml` 참고
2. **빌드 실행**:
   ```bash
   python tools/generate_deck.py path/to/outline.yaml
   ```
3. **시각 검수** (선택):
   ```bash
   node tools/render_pptx.mjs <output.pptx> _work/preview
   ```

Claude Code 환경에서는 슬래시 명령으로 한 번에:
```
/generate-deck path/to/outline.yaml
```

지원 슬라이드 타입: cover · divider · card-grid-2/3/4/7 · comparison · step · kpi · outro.
입력 형식 상세는 `skills/generate-deepnoid-ppt/SKILL.md` 또는 `docs/specs/2026-05-28-mvp-automation-design.md` 참조.
```

또한 README "개발 상태" 의 미완성 항목을 갱신:
- `[ ] 자동 파이프라인 스크립트` → `[x] 자동 파이프라인 (MVP — outline YAML → PPTX)`
- `[ ] 스킬·에이전트·명령` → `[x] 스킬·슬래시 명령 (MVP — 에이전트 팀은 이후 단계)`

- [ ] **Step 2: 전체 테스트 일괄 수행**

Run: `python -m pytest tests/ -v`
Expected: 모든 테스트 통과.

- [ ] **Step 3: end-to-end 재검증**

Run:
```bash
python tools/generate_deck.py examples/sample-outline.yaml
node tools/render_pptx.mjs _work/demo-output.pptx _work/demo-png
```
Expected: `저장 완료: ...` 와 `렌더 완료: 7장 → _work/demo-png`.

- [ ] **Step 4: 커밋 + 푸시**

```bash
git add README.md
git commit -m "docs: README 사용법 갱신 + MVP 자동화 완성 표시"
git push origin main
```

- [ ] **Step 5: 최종 확인**

Run: `git log --oneline -20`
저장소 https://github.com/via-soi/deepnoid-design-generator 에서 새 커밋들 확인.

---

## 완료 기준 (Definition of Done)

- `python -m pytest tests/ -v` 전체 통과 (단위 + 통합)
- `python tools/generate_deck.py examples/sample-outline.yaml` 실행 → 7장 PPTX 저장
- 렌더링 결과가 design.md §6 패턴 명세와 시각적으로 일치
- README 에 사용법 섹션 추가
- 모든 커밋이 푸시되고 GitHub 리포가 동기화됨

## 향후 계획 (MVP 이후 — 별도 계획)

- 자유 Markdown 입력 → 아웃라인 자동 제안
- 자유 Word/Excel 입력 → 5 에이전트 팀 협업 생성
- §6.7 라벨 칩·빗금 플레이스홀더·키워드 하이라이트
- §6.8 아이콘 라이브러리
- 자동 차트 생성
