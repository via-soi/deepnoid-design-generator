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
    print(f"추출: {name} ({w}x{h}) -> {DST}")


if __name__ == "__main__":
    main()
