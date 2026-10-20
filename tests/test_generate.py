from mass_ppt.generate import generate
from tests.conftest import run_table

FULL = {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}


def test_generate_end_to_end(ppt_1010, fields_1018, get_images):
    out, report = generate(ppt_1010, fields_1018, FULL, get_images)
    assert report['slide_count'] == 62
    assert [h['num'] for h in report['hymns']] == [58, 106, 128, 129, 61]
    assert report['warnings'] == []
    texts = {v[0] for v in run_table(out).values()}
    assert '2026.10.17' in texts
    assert '◎ 주님은 민족들의 눈앞에 당신 정의를 드러내셨네.' in texts


def test_generate_tolerates_none_everywhere(ppt_1010, get_images):
    out, report = generate(ppt_1010, {}, {'entrance': None, 'communion': []}, get_images)
    assert report['slide_count'] == 53
    assert report['warnings']
