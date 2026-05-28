from pathlib import Path
from tools.deepnoid_builder.tokens import load_tokens

DESIGN_MD = Path("skills/generate-deepnoid-ppt/assets/DEEPNOID-Design.md")


def test_loads_main_colors():
    t = load_tokens(DESIGN_MD)
    assert t["colors"]["primary"] == "#0066FF"
    assert t["colors"]["green"] == "#34D0B3"
    assert t["colors"]["black"] == "#000000"


def test_loads_typography():
    t = load_tokens(DESIGN_MD)
    # body 토큰은 frontmatter 에 존재한다
    assert "fontSize" in t["typography"]["body"]
    assert "fontFamily" in t["typography"]["title"]


def test_loads_header_and_outro():
    t = load_tokens(DESIGN_MD)
    assert "logo-asset" in t["header"]
    assert "slogan-text" in t["header"]
    assert "slogan-pos" in t["outro"]


def test_raises_on_missing_frontmatter(tmp_path):
    bad = tmp_path / "no-fm.md"
    bad.write_text("just markdown, no frontmatter", encoding="utf-8")
    import pytest
    with pytest.raises(ValueError, match="frontmatter"):
        load_tokens(bad)
