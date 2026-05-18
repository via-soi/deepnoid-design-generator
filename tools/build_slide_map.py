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
