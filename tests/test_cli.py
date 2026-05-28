import subprocess
import sys
from pathlib import Path


def test_cli_builds_pptx(tmp_path):
    outline_yaml = tmp_path / "outline.yaml"
    outline_yaml.write_text("""
deck:
  title: Test
  output: {out}
slides:
  - type: cover
    title: 표지 제목
    subtitle: 부제
  - type: outro
""".format(out=str(tmp_path / "out.pptx")).replace("\\", "/"),
                            encoding="utf-8")
    result = subprocess.run(
        [sys.executable, "tools/generate_deck.py", str(outline_yaml)],
        capture_output=True, text=True, encoding="utf-8",
    )
    assert result.returncode == 0, result.stderr
    out = tmp_path / "out.pptx"
    assert out.exists()
    assert out.stat().st_size > 1000
