"""
hymns.py의 단위 테스트.
"""
import pytest
from pathlib import Path
from mass_ppt.hymns import _find_hymn_slides, _find_hymn_slide_groups


class TestFindHymnSlides:
    """성가 슬라이드 탐지 테스트."""

    def test_identify_hymn_slide(self):
        """성가 번호 슬라이드 식별."""
        # 실제 PPT 파일이 필요
        archive_ppt = Path('archive/261010미사.pptx')
        if not archive_ppt.exists():
            pytest.skip("archive PPT 파일이 없습니다")

        from pptx import Presentation
        ppt = Presentation(str(archive_ppt))

        # 슬라이드 10 (11번 성가)
        slide = ppt.slides[9]  # 0-indexed
        from lxml import etree
        import zipfile

        with zipfile.ZipFile(str(archive_ppt), 'r') as z:
            slide_xml = z.read('ppt/slides/slide10.xml')

        tree = etree.fromstring(slide_xml)
        hymn_num = _find_hymn_slides(tree)

        assert hymn_num == 11

    def test_non_hymn_slide_returns_none(self):
        """성가 슬라이드가 아니면 None 반환."""
        # 텍스트만 있는 슬라이드는 None
        pass


class TestFindHymnSlideGroups:
    """성가 슬라이드 그룹 탐지 테스트."""

    def test_group_by_position(self):
        """성가 슬라이드를 위치별로 그룹화."""
        # entrance, offertory, communion, recessional로 분류
        pass

    def test_continuous_hymn_slides(self):
        """연속된 같은 번호의 슬라이드 묶음."""
        pass


class TestHymnImagePlacement:
    """성가 이미지 배치 테스트."""

    def test_image_dimensions_calculation(self):
        """이미지를 content 영역에 맞게 배치 계산."""
        pass

    def test_preserve_aspect_ratio(self):
        """이미지 비율 유지."""
        pass


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
