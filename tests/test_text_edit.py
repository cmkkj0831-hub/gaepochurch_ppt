from mass_ppt.pptx_package import Package, qn
from mass_ppt.text_edit import apply_text_fields, find_text_slides
from tests.conftest import run_table

EXPECTED_CHANGES = {
    1: [('2026.10.10', '2026.10.17'),
        (' 연중 제 28 주일(군인 주일)', ' 연중 제 29 주일(전교 주일)')],
    27: [('25,6-10ㄱ', '2,1-5')],
    28: [('◎ 저는 오래오래 주님 집에 사오리다.', '◎ 주님은 민족들의 눈앞에 당신 정의를 드러내셨네.')],
    29: [('▥ 사도 바오로의 필리피서 말씀입니다.', '▥ 사도 바오로의 로마서 말씀입니다.'),
         ('4,12-14.19-20', '10,9-18')],
    30: [('○ 우리 주 예수 그리스도의 아버지, 저희 마음의 눈을 밝혀 주시어 부르심을 받은 저희의 희망을 알게 하여 주소서.',
          '○ 주님이 말씀하신다. 너희는 가서 모든 민족들을 가르쳐라. 내가 세상 끝 날까지 언제나 너희와 함께 있으리라.')],
    31: [('22,1-14', '28,16-20')],
    56: [('부자들도 궁색해져 굶주리게 되지만, 주님을 찾는 이에게는 좋은 것뿐이리라.',
          '주님이 말씀하신다. 내가 너희에게 명령한 것을 모든 민족들에게 가르쳐 지키게 하여라. 내가 세상 끝 날까지 언제나 너희와 함께 있으리라.')],
}


def test_identifies_each_text_slide_once(ppt_1010):
    assert find_text_slides(Package(ppt_1010)) == {
        '제1독서': [27], '화답송': [28], '제2독서': [29], '복음환호송': [30], '복음': [31], '영성체송': [56]}


def test_only_target_runs_change(ppt_1010, fields_1018):
    out, report = apply_text_fields(ppt_1010, fields_1018)
    before, after = run_table(ppt_1010), run_table(out)
    assert before.keys() == after.keys()
    diffs = {}
    for key in before:
        assert before[key][1] == after[key][1], f'서식 변경: {key}'
        if before[key][0] != after[key][0]:
            diffs.setdefault(key[0], []).append((before[key][0], after[key][0]))
    assert {k: sorted(v) for k, v in diffs.items()} == {k: sorted(v) for k, v in EXPECTED_CHANGES.items()}
    assert report['warnings'] == []


def test_gospel_and_title_runs_keep_leading_space(ppt_1010, fields_1018):
    out, _ = apply_text_fields(ppt_1010, fields_1018)
    texts = [v[0] for k, v in run_table(out).items() if k[0] == 31]
    assert ' ✠ 마태오가 전한 거룩한 복음입니다.' in texts
    assert '복음' in texts


def test_none_value_is_skipped_with_warning(ppt_1010, fields_1018):
    fields = dict(fields_1018, psalm_refrain=None)
    out, report = apply_text_fields(ppt_1010, fields)
    assert any('화답송' in w for w in report['warnings'])
    t = [v[0] for k, v in run_table(out).items() if k[0] == 28]
    assert '◎ 저는 오래오래 주님 집에 사오리다.' in t


def _with_slide_title_changed(ppt_bytes, slide_no, old, new):
    pkg = Package(ppt_bytes)
    part = pkg.slide_parts()[slide_no - 1]
    root = pkg.xml(part)
    for t in root.iter(qn('a:t')):
        if t.text == old:
            t.text = new
    pkg.write_xml(part, root)
    return pkg.save()


def test_duplicate_title_slide_is_not_edited(ppt_1010, fields_1018):
    dup = _with_slide_title_changed(ppt_1010, 30, '복 음 환 호 송', '화답송')
    out, report = apply_text_fields(dup, fields_1018)
    assert any('"화답송" 슬라이드가 2개' in w for w in report['warnings'])
    assert any('"복음환호송" 슬라이드가 0개' in w for w in report['warnings'])
    before, after = run_table(dup), run_table(out)
    assert all(before[k] == after[k] for k in before if k[0] in (28, 30))


def test_body_word_does_not_make_gospel_slide(ppt_1010):
    found = find_text_slides(Package(ppt_1010))
    assert found['복음'] == [31]
