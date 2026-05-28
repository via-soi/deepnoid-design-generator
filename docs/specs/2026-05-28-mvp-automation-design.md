# DEEPNOID Design Generator — MVP 자동화 설계

- **작성일:** 2026-05-28
- **상태:** 확정 (구현 계획 수립 대기)
- **저장소:** `deepnoid-design-generator` (GitHub, Public)
- **이전 설계:** `docs/specs/2026-05-16-deepnoid-ppt-generator-design.md` (템플릿 클론 + 5 에이전트 팀) — 본 설계로 갱신·축소

---

## 1. 배경 — 설계 갱신 이유

원본 설계(2026-05-16)는 **템플릿 클론 방식 + 5 에이전트 팀**(지휘자 + 콘텐츠 분석가 + 레이아웃 설계자 + 슬라이드 조립가 + 충실도 검수자)으로 설계됐다. 이후 운영 결정으로:

- 내지 사이드바를 제거한 PPT 템플릿(`deepnoid-ppt-template.pptx`)이 리포에서 삭제됨
- 디자인 시스템 명세 `DEEPNOID-Design.md`(v3.0)가 단일 출처로 확정됨 — §6 패턴별 좌표·크기·색·서체가 cm/pt 단위로 명시
- 리포에 남은 PPT 자산은 `docs/reference/`의 원본 57장 템플릿뿐이며, 이는 사이드바가 있는 구버전이라 클론 기반 생성에 적합하지 않음

따라서 자동화 방식은 **클론 → 명세 기반 빌드**로 전환되며, 5 에이전트 팀은 MVP에서 **단일 스킬**로 축소된다.

## 2. 목표와 비목표

**목표**
- 사용자가 구조화된 아웃라인(YAML)을 제공하면 DEEPNOID 디자인을 따른 PPTX를 자동 생성
- `DEEPNOID-Design.md`의 §6 핵심 패턴 5종 + 표지·간지·아웃트로를 코드로 재현
- 슬래시 명령 1개로 한 번에 빌드·렌더·산출

**비목표 (MVP 제외)**
- 자유 Word·Excel 문서 입력 → 아웃라인 자동 제안
- 5 에이전트 팀 협업 생성
- §6.7 강조 컴포넌트(라벨 칩·빗금 플레이스홀더·키워드 하이라이트)
- §6.8 아이콘 라이브러리
- 차트 자동 생성 (외부에서 이미지로 받음)

## 3. 핵심 결정사항

| 항목 | 결정 | 이유 |
|---|---|---|
| 입력 | **구조화 YAML 아웃라인** | 명시적·결정적. LLM 매핑 불필요. 사용자가 슬라이드 단위 제어 |
| 생성 방식 | **python-pptx로 명세 기반 빌드** | 클론 가능한 깨끗한 템플릿 없음. design.md가 cm/pt까지 명시 → 재현 가능 |
| 오케스트레이션 | **단일 스킬 (main 세션)** | 에이전트 팀 오버헤드 회피. 작은 스코프에 적합 |
| 진입점 | **`/generate-deck` 슬래시 명령** | 한 명령으로 빌드 완료 |
| 디자인 단일 출처 | **`DEEPNOID-Design.md`** | frontmatter 토큰(colors, typography, spacing, header, chapter-divider, outro)을 빌더가 로드 |

## 4. 아키텍처

```
사용자 → /generate-deck outline.yaml
         ↓
        SKILL.md (오케스트레이션)
         ↓
        tools/generate_deck.py (CLI)
         ↓
        deepnoid_builder/
         ├─ tokens.py      ← DEEPNOID-Design.md frontmatter 로드
         ├─ headers.py     ← 우상단 헤더 (슬로건+로고+페이지번호)
         ├─ patterns.py    ← §6.1~6.5 패턴 함수
         └─ builder.py     ← 표지·간지·아웃트로 + 슬라이드 조립
         ↓
        python-pptx → output.pptx
         ↓ (선택)
        tools/render_pptx.mjs → PNG 검수
```

**경계와 책임**
- `tokens.py` — `DEEPNOID-Design.md`의 frontmatter YAML을 파싱해 색상·서체·간격·헤더·간지·아웃트로 토큰을 dict로 반환
- `headers.py` — 우상단 헤더(슬로건 + 로고 + 페이지번호)를 슬라이드에 일관 배치
- `patterns.py` — §6.1~6.5 패턴마다 함수 1개. `add_card(slide, ...)`, `add_card_grid(slide, n, items)`, `add_comparison(slide, asis, tobe)`, `add_step_flow(slide, steps)`, `add_kpi_cards(slide, cards)`
- `builder.py` — 외부 진입점 `build(outline_dict) -> Presentation`. 슬라이드 타입별로 분기해 적절한 함수 조합

**의존성:** python-pptx, PyYAML, Pillow (이미 `requirements.txt` 포함 / PyYAML 추가 필요).

## 5. 입력 형식 — `outline.yaml`

```yaml
deck:
  title: "발표 제목"            # 아카이브 메타
  subtitle: "부제·발표자"
  output: "C:/path/output.pptx"

slides:
  - type: cover
    title: "표지 제목"
    subtitle: "부제·발표자·일자"

  - type: divider
    chapter: "Chapter 01"        # 선택, 없으면 생략
    title: "섹션 제목"
    number: 1                    # 좌하단 큰 글리프 숫자

  - type: card-grid-3            # card-grid-2|3|4|7 — 카드 수에 따라
    eyebrow: "Intro / 지금 우리 풍경"
    title: "내지 페이지 제목"
    cards:
      - { header: "...", body: "...", accent: blue }     # accent: blue|green|none
      - { header: "...", body: "...", accent: green }
      - { header: "...", body: "..." }                   # 기본: 그레이

  - type: comparison
    eyebrow: "Why / 왜 써야해?"
    title: "..."
    asis: { label: "AS-IS · ...", header: "...", bullets: ["...", "...", "..."] }
    tobe: { label: "TO-BE · ...", header: "...", bullets: ["...", "...", "..."] }
    caption: "하단 캡션 (선택)"

  - type: step
    eyebrow: "How / 4단계 프레임"
    title: "..."
    steps:
      - { label: "STEP 1", header: "...", body: "..." }
      # 3~5단계
    footer: "하단 통찰 문구 (선택)"

  - type: kpi
    eyebrow: "Closing / Takehome"
    title: "..."
    cards:
      - { number: "01", header: "...", body: "..." }   # 3~4장, BLUE/GREEN 교대 자동
      # ...

  - type: outro                  # 인자 없음 — design.md outro 토큰 그대로 사용
```

**지원 슬라이드 타입:** `cover`, `divider`, `card-grid-2`, `card-grid-3`, `card-grid-4`, `card-grid-7`, `comparison`, `step`, `kpi`, `outro`.

**검증 규칙 (빌더가 사전 검증):**
- `type`이 위 목록에 없으면 오류
- 필수 필드 누락(예: `cover.title`, `divider.title`)이면 오류
- `card-grid-N`에서 카드 수가 N과 다르면 경고(자동 절단 또는 빈 카드)
- `step.steps`는 3~5개

## 6. 파일 구성 (리포에 추가될 것)

| 경로 | 책임 | 신규/수정 |
|---|---|---|
| `commands/generate-deck.md` | `/generate-deck` 슬래시 명령 정의 | 신규 |
| `skills/generate-deepnoid-ppt/SKILL.md` | 스킬 오케스트레이션 지침 | 신규 |
| `skills/generate-deepnoid-ppt/assets/deepnoid_logo.png` | 헤더·아웃트로 로고 (원본 템플릿에서 추출) | 신규 |
| `tools/deepnoid_builder/__init__.py` | 패키지 진입점 | 신규 |
| `tools/deepnoid_builder/tokens.py` | DEEPNOID-Design.md 토큰 로더 | 신규 |
| `tools/deepnoid_builder/headers.py` | 우상단 헤더 빌더 | 신규 |
| `tools/deepnoid_builder/patterns.py` | §6.1~6.5 패턴 함수 | 신규 |
| `tools/deepnoid_builder/builder.py` | 슬라이드 조립 (cover/divider/outro/패턴 디스패치) | 신규 |
| `tools/generate_deck.py` | CLI 진입점: `python tools/generate_deck.py outline.yaml` | 신규 |
| `examples/sample-outline.yaml` | 입력 예시 (Tech Talks AX 데모) | 신규 |
| `requirements.txt` | `PyYAML` 추가 | 수정 |
| `README.md` | 사용법 안내 추가 | 수정 |

## 7. 동작 흐름

1. 사용자가 `examples/sample-outline.yaml`을 참고해 outline 작성
2. 셸에서 `/generate-deck path/to/outline.yaml` 실행 (또는 `python tools/generate_deck.py outline.yaml`)
3. SKILL.md가 빌더 CLI 호출
4. 빌더가 `DEEPNOID-Design.md` 토큰 로드 → 슬라이드별 패턴 함수 디스패치 → `python-pptx` 객체 조립 → 저장
5. (선택) `node tools/render_pptx.mjs output.pptx _work/preview`로 PNG 렌더해 시각 검수

## 8. 디자인 충실도 보장

`DEEPNOID-Design.md`의 cm/pt 좌표를 그대로 코드 상수로 옮기는 방식으로 보장한다. 클론 방식의 "1픽셀 보존"과 달리 **명세 1픽셀 보존**(코드가 명세와 일치)이 보장 기준이다.

- 색상: 토큰 dict에서 HEX 그대로 사용
- 서체: Pretendard 800/700/400 — 시스템에 미설치 시 빌더가 경고
- 좌표: cm 단위로 `Cm()` 변환. 모든 §6 패턴 함수는 design.md 표의 좌표·크기를 1:1 옮긴 상수 사용
- 간격: `spacing.margin-*`, `spacing.content-gap` 토큰 참조

## 9. 단위 분해 (Testability)

각 패턴 함수는 다음 두 가지를 만족하도록 설계한다:

1. **단독 호출 가능** — 빈 슬라이드 + 인자만으로 호출해 슬라이드 1장 생성 가능 (조립가/검수자 없이)
2. **순수 함수에 가깝게** — 슬라이드 객체 + 인자 dict를 받아 도형을 추가하고 슬라이드 객체를 반환. 외부 상태·전역 의존 최소화

검증:
- 패턴 함수별 더미 데이터로 1슬라이드 생성 → 렌더링 → 명세와 시각 비교
- end-to-end: `examples/sample-outline.yaml` → 8~10장 덱 → 렌더링 → 디자인 명세와 시각 비교

## 10. 실패 모드와 처리

| 실패 | 처리 |
|---|---|
| YAML 파싱 오류 | 라인·컬럼 메시지와 함께 즉시 중단 |
| 알 수 없는 `type` | 지원 목록을 보여주고 중단 |
| 필수 필드 누락 | 어떤 슬라이드의 어떤 필드인지 명시하고 중단 |
| 카드 수 불일치 | 경고 + 자동 절단/패딩 (절대 멈추지 않음) |
| Pretendard 미설치 | 경고만, 폴백 폰트로 진행 |
| 로고 PNG 누락 | 헤더에서 로고만 생략, 슬로건+페이지번호는 유지 |
| Python 의존성 누락 | `pip install -r requirements.txt` 안내 후 중단 |

## 11. 구현 시 결정할 항목 (계획 단계로 미룸)

- 카드 그리드 `card-grid-7`의 8번째 안내 카드 자동 채움 방식
- 비교형의 `caption` 위치·폭 (design.md에는 "슬라이드 하단 y=16.0cm"로만 명시)
- 토큰 로더의 캐싱 여부 (한 빌드에서 여러 번 호출되는 경우)
- 로고 PNG의 종횡비 유지·재색상화 금지 검증

## 12. 향후 확장 (MVP 이후)

- 자유 Markdown 입력 → 아웃라인 자동 제안 (중간 옵션)
- 자유 Word·Excel 입력 → 5 에이전트 팀 협업 생성 (풀스코프)
- §6.7 강조 컴포넌트, §6.8 아이콘
- 자동 차트 생성
- 시각 검수 자동화 (오버플로우·잔여 플레이스홀더 검출)
