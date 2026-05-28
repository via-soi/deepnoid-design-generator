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
