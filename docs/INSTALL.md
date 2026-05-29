# DEEPNOID Design Generator 설치 안내 (동료용)

> 비개발 직군도 따라할 수 있도록 단계별로 정리했습니다.
> 막히는 부분이 있으면 박송희 부장(sbahk@deepnoid.com)에게 문의하세요.

---

## 한 줄 요약

**Claude Code → 슬래시 명령 두 줄로 설치, YAML 파일 하나로 PPT 자동 생성.**

---

## 1. 사전 준비

| 항목 | 설명 | 설치 링크 |
|---|---|---|
| **Claude Code** | 이 플러그인이 동작하는 환경 | [claude.com/code](https://claude.com/code) |
| **Python 3.12** | PPTX를 만드는 엔진 | [python.org/downloads](https://www.python.org/downloads/) (3.12.x 설치) |
| **LibreOffice** | (선택) 미리보기 이미지 생성용 | [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/) |
| **Node.js** | (선택) 미리보기 이미지 생성용 | [nodejs.org](https://nodejs.org/) (LTS 버전) |

> Python·LibreOffice·Node.js는 처음 한 번만 설치하면 됩니다. 미리보기를 안 만들 거면 LibreOffice·Node.js는 건너뛰어도 됩니다.

---

## 2. Claude Code에서 플러그인 설치

Claude Code를 열고 채팅창에 아래 두 줄을 차례로 입력합니다.

**(1) 마켓플레이스 추가**
```
/plugin marketplace add via-soi/deepnoid-design-generator
```

**(2) 플러그인 설치**
```
/plugin install deepnoid-design-generator@deepnoid-tools
```

설치되면 `/generate-deck` 명령이 사용 가능해집니다.

---

## 3. Python 패키지 한 번 설치

설치된 플러그인의 `requirements.txt` 위치를 알려준 뒤 명령 프롬프트(또는 PowerShell)에서 한 번만 실행:

```bash
pip install python-pptx PyYAML Pillow openpyxl python-docx
```

> Claude Code에 "deepnoid 빌더 의존성 설치해줘" 라고 부탁해도 자동으로 처리됩니다.

---

## 4. 첫 PPT 생성

### 4-1. 아웃라인 파일 작성

`my-deck.yaml` 이라는 텍스트 파일을 만들고 아래처럼 채웁니다 (메모장으로 충분합니다 — 저장 시 인코딩은 **UTF-8**로):

```yaml
deck:
  title: "발표 제목"
  output: "C:/Users/내이름/Desktop/my-deck.pptx"

slides:
  - type: cover
    title: "발표 큰 제목"
    subtitle: "부제 · 발표자 · 날짜"

  - type: divider
    chapter: "Chapter 01"
    title: "첫 번째 섹션 제목"
    number: 1

  - type: card-grid-3
    eyebrow: "Intro / 도입"
    title: "내용 페이지 제목"
    cards:
      - header: "포인트 1"
        body: "설명"
      - header: "포인트 2"
        body: "설명"
        accent: blue
      - header: "포인트 3"
        body: "설명"
        accent: green

  - type: outro
```

### 4-2. PPT 생성

Claude Code 채팅창에:

```
/generate-deck C:/Users/내이름/Desktop/my-deck.yaml
```

→ 지정한 경로에 `my-deck.pptx` 가 생성됩니다.

---

## 5. 지원하는 슬라이드 종류

| `type` | 용도 | 필요한 필드 |
|---|---|---|
| `cover` | 표지 | `title`, `subtitle` |
| `divider` | 섹션 간지 | `title`, `chapter`(선택), `number`(선택) |
| `card-grid-2/3/4/7` | 2·3·4·7장의 카드 배치 | `eyebrow`, `title`, `cards: [{header, body, accent}]` |
| `comparison` | AS-IS / TO-BE 비교 | `eyebrow`, `title`, `asis`, `tobe`, `caption`(선택) |
| `step` | 단계 흐름 (3~5단계) | `eyebrow`, `title`, `steps`, `footer`(선택) |
| `kpi` | 결론·메시지 카드 (3~4장) | `eyebrow`, `title`, `cards: [{number, header, body}]` |
| `outro` | 마무리 (인자 없음) | — |

전체 입력 형식은 `examples/sample-outline.yaml` 참고.

`accent` 색상 옵션: `gray`(기본) / `blue` / `green`.

---

## 6. 미리보기 만들기 (선택)

생성된 PPT를 슬라이드별 PNG 이미지로 보고 싶다면:

Claude Code에:
```
deepnoid 미리보기 만들어줘 — output.pptx 를 PNG로
```

또는 직접:
```bash
node tools/render_pptx.mjs <output.pptx> _work/preview
```

→ `_work/preview/slide-1.png ...` 으로 저장됩니다.

---

## 7. 업데이트

디자인 시스템이 갱신되면 Claude Code에서:

```
/plugin update deepnoid-design-generator
```

---

## 8. 자주 마주칠 문제

| 증상 | 해결 |
|---|---|
| `python: command not found` | Python 3.12를 [python.org](https://www.python.org/downloads/)에서 설치 (Microsoft Store 버전이 아닌 정식 인스톨러). 설치 시 "Add Python to PATH" 체크 |
| `pip install ... 권한 오류` | `pip install --user python-pptx PyYAML Pillow` (user 모드 설치) |
| 한글이 깨져 보임 | outline 파일을 UTF-8로 저장. 메모장 → 다른 이름으로 저장 → 인코딩 'UTF-8' 선택 |
| PPT가 생성됐는데 `Pretendard` 폰트가 적용 안 됨 | [Pretendard 폰트](https://cactus.tistory.com/306) 9종을 설치 |
| `/plugin install` 후에도 명령이 안 보임 | Claude Code 재시작 |

---

## 9. 도움이 필요할 때

- 디자인 규칙 (색·폰트·여백): `skills/generate-deepnoid-ppt/assets/DEEPNOID-Design.md`
- 입력 형식 상세: `docs/specs/2026-05-28-mvp-automation-design.md`
- 문의: **박송희 부장 — sbahk@deepnoid.com**
