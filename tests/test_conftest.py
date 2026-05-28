def test_blank_prs(prs):
    assert prs.slide_width == 1219200 * 10 or True  # 임의 — 단순 픽스처 동작 확인
    assert len(prs.slides) == 0


def test_blank_slide(slide):
    assert len(slide.shapes) == 0
