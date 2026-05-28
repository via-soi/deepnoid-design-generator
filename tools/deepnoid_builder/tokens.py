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
