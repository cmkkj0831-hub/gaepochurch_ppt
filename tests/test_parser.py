import pytest
from bs4 import BeautifulSoup

from mass_ppt.parser import FIELD_KEYS, make_cover_title, parse_html
from tests.conftest import FIXTURES


def load(day):
    return (FIXTURES / f'missa_{day}.html').read_text(encoding='utf-8')


def test_1018_expected_values():
    f = parse_html(load('2026-10-18'))
    assert f['cover_title'] == '연중 제 29 주일(전교 주일)'
    assert f['cover_date'] == '2026.10.17'
    assert (f['first_reading_intro'], f['first_reading_ref']) == ('▥ 이사야서의 말씀입니다.', '2,1-5')
    assert f['psalm_refrain'] == '◎ 주님은 민족들의 눈앞에 당신 정의를 드러내셨네.'
    assert (f['second_reading_intro'], f['second_reading_ref']) == ('▥ 사도 바오로의 로마서 말씀입니다.', '10,9-18')
    assert f['gospel_acclamation'] == ('○ 주님이 말씀하신다. 너희는 가서 모든 민족들을 가르쳐라. '
                                       '내가 세상 끝 날까지 언제나 너희와 함께 있으리라.')
    assert (f['gospel_intro'], f['gospel_ref']) == ('✠ 마태오가 전한 거룩한 복음입니다.', '28,16-20')
    assert f['communion_antiphon'] == ('주님이 말씀하신다. 내가 너희에게 명령한 것을 모든 민족들에게 '
                                       '가르쳐 지키게 하여라. 내가 세상 끝 날까지 언제나 너희와 함께 있으리라.')


def test_1011_matches_261010_template_text():
    f = parse_html(load('2026-10-11'))
    assert f['cover_title'] == '연중 제 28 주일(군인 주일)'
    assert f['cover_date'] == '2026.10.10'
    assert f['first_reading_ref'] == '25,6-10ㄱ'
    assert f['psalm_refrain'] == '◎ 저는 오래오래 주님 집에 사오리다.'
    assert (f['second_reading_intro'], f['second_reading_ref']) == ('▥ 사도 바오로의 필리피서 말씀입니다.', '4,12-14.19-20')
    assert f['gospel_ref'] == '22,1-14'
    assert f['communion_antiphon'] == '부자들도 궁색해져 굶주리게 되지만, 주님을 찾는 이에게는 좋은 것뿐이리라.'


def test_1004_no_feast_name():
    f = parse_html(load('2026-10-04'))
    assert f['cover_title'] == '연중 제 27 주일'
    assert f['cover_date'] == '2026.10.03'
    assert f['first_reading_ref'] == '5,1-7'
    assert all(f[k] for k in FIELD_KEYS)


@pytest.mark.parametrize('raw, expected', [
    ('(녹) 연중 제29주일, 민족들의 복음화를 위한 미사(전교 주일)', '연중 제 29 주일(전교 주일)'),
    ('(녹)연중 제28주일(군인 주일)', '연중 제 28 주일(군인 주일)'),
    ('(녹) 연중 제27주일', '연중 제 27 주일'),
    (None, None),
])
def test_make_cover_title(raw, expected):
    assert make_cover_title(raw) == expected


@pytest.mark.parametrize('section, keys', [
    ('화답송', ['psalm_refrain']),
    ('복음 환호송', ['gospel_acclamation']),
    ('제2독서', ['second_reading_intro', 'second_reading_ref']),
    ('영성체송', ['communion_antiphon']),
])
def test_missing_section_gives_none_not_other_values(section, keys):
    soup = BeautifulSoup(load('2026-10-18'), 'html.parser')
    h3 = soup.find('h3', string=section)
    h3.find_next_sibling().decompose()
    h3.decompose()
    f = parse_html(str(soup))
    for k in keys:
        assert f[k] is None
    full = parse_html(load('2026-10-18'))
    for k in FIELD_KEYS:
        if k not in keys:
            assert f[k] == full[k]


def test_page_wide_symbols_outside_sections_are_ignored():
    html = load('2026-10-18').replace(
        '<form name="contents">',
        '<div class="calendar">○ 달력 ◎ 메뉴 ▥ 가짜서 말씀입니다.99,1-2</div><form name="contents">')
    assert parse_html(html) == parse_html(load('2026-10-18'))
