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
