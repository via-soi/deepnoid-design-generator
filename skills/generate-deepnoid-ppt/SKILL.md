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
