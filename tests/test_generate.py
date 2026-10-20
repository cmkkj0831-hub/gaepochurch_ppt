"""
generate.py의 통합 테스트.
"""
import pytest
from pathlib import Path
from mass_ppt.generate import generate_from_local_hymns


class TestGenerateFromLocalHymns:
    """통합 생성 함수 테스트."""

    def test_generate_complete_ppt(self):
        """전체 PPT 생성."""
        archive_ppt = Path('archive/261017미사.pptx')
        if not archive_ppt.exists():
            pytest.skip("archive PPT 파일이 없습니다")

        with open(archive_ppt, 'rb') as f:
            base_ppt = f.read()

        # 테스트 데이터
        mass_url = "https://maria.catholic.or.kr/mi_pr/missa/missa.asp?menu=missa&missaid=16874&gomonth=2026-10-18&missatype=DA"
        hymn_numbers = {
            'entrance': 58,
            'offertory': 106,
            'communion': [128, 129],
            'recessional': 61
        }

        # 생성
        result_bytes, report = generate_from_local_hymns(
            base_ppt,
            mass_url,
            hymn_numbers
        )

        # 검증
        assert isinstance(result_bytes, bytes)
        assert len(result_bytes) > 0
        assert 'updated_texts' in report
        assert 'inserted_hymns' in report
        assert 'warnings' in report

    def test_missing_hymn_image_warning(self):
        """없는 성가 이미지에 대한 경고."""
        pass

    def test_communion_slide_positioning(self):
        """성체성가 슬라이드 배치 검증."""
        pass


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
