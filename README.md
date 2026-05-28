# DEEPNOID Design Generator

문서(Excel·Word 등)를 DEEPNOID 공식 템플릿 디자인에 맞춰 발표 자료·문서로
자동 변환하는 Claude Code 플러그인입니다.

> 설계 문서: `docs/specs/2026-05-16-deepnoid-ppt-generator-design.md`

## 포함 범위

이 저장소는 플러그인 프레임워크와 **DEEPNOID PPT 템플릿**을 함께 공개합니다.

| 항목 | 포함 여부 | 위치 |
|---|---|---|
| 플러그인 프레임워크 (코드·스크립트·설계 문서) | 포함 | `tools/`, `docs/`, `.claude-plugin/` |
| DEEPNOID 디자인 시스템 (컬러·타이포·레이아웃 규칙) | 포함 | `skills/generate-deepnoid-ppt/assets/DEEPNOID-Design.md` |
| 원본 PPT 템플릿 (57장) + 유형별 슬라이드 이미지 | 포함 | `docs/reference/` |
| DEEPNOID Word 템플릿 | 미포함 (비공개 별도 관리) | — |

## 동작 개념

- **PPT(.pptx) 산출물** — DEEPNOID PPT 템플릿을 기준으로 슬라이드를 생성.
  내지(內紙)는 양 끝 사이드바가 없는 최신 디자인을 사용.
- **Word(.docx) 산출물** — DEEPNOID Word 템플릿을 기준으로 문서를 생성
  (Word 템플릿은 비공개 자료로, 구동 시 별도 제공 필요).
- 원본 템플릿을 복제하고 텍스트·데이터만 교체해 디자인을 그대로 보존하는 방식.

## 사용 방법 (MVP 자동화 빌더)

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

지원 슬라이드 타입: `cover` · `divider` · `card-grid-2/3/4/7` · `comparison` · `step` · `kpi` · `outro`.
입력 형식 상세는 `skills/generate-deepnoid-ppt/SKILL.md` 또는 `docs/specs/2026-05-28-mvp-automation-design.md` 참조.

## 구성

| 경로 | 내용 |
|---|---|
| `.claude-plugin/plugin.json` | 플러그인 매니페스트 |
| `skills/generate-deepnoid-ppt/assets/` | DEEPNOID 디자인 시스템 명세(`DEEPNOID-Design.md`) |
| `tools/` | PPTX 추출·렌더링·매핑 보조 스크립트 |
| `docs/specs/` · `docs/plans/` | 설계 문서 · 구현 계획 |
| `docs/reference/` | 원본 PPT 템플릿 + 유형별 슬라이드 이미지 |
| `requirements.txt` | Python 런타임 의존성 |

## 개발 상태 (파일럿)

- [x] 리포 골격 — 플러그인 매니페스트·LICENSE·README
- [x] Python 런타임 의존성 정의 (`requirements.txt`)
- [x] 보조 스크립트 — 추출·렌더링·매핑
- [x] DEEPNOID PPT 템플릿 포함
- [x] 자동 파이프라인 (MVP — outline YAML → PPTX)
- [x] 스킬·슬래시 명령 (MVP — 에이전트 팀은 이후 단계)
