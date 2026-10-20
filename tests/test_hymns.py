import io
import posixpath

import pytest
from pptx import Presentation

from mass_ppt.hymns import HymnTemplateError, find_hymn_bundles, hymn_number, place_image, replace_hymns
from mass_ppt.pptx_package import CONTENT_TYPES, NS, Package, qn, resolve
from mass_ppt.storage import hymn_page_files

FULL = {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}


def numbers(ppt_bytes):
    pkg = Package(ppt_bytes)
    return [n for p in pkg.slide_parts() if (n := hymn_number(pkg.xml(p))) is not None]


def test_hymn_slide_detection(ppt_1010):
    pkg = Package(ppt_1010)
    parts = pkg.slide_parts()
    assert hymn_number(pkg.xml(parts[9])) == 11
    assert hymn_number(pkg.xml(parts[8])) is None   # 그림 + "미사가 곧 시작합니다"
    assert hymn_number(pkg.xml(parts[59])) is None  # 그림 2장 (성가 아님)
    assert hymn_number(pkg.xml(parts[12])) is None  # 그림만


def test_bundles_in_template(ppt_1010):
    assert find_hymn_bundles(Package(ppt_1010)) == [[10, 11], [33, 34], [57, 58, 59], [61, 62]]


def test_place_image_matches_template_for_wide_image():
    assert place_image(1104, 644) == (1567543, 330200, 10624457, 6197599)


def test_place_image_tall_image_fills_height_centered():
    x, y, cx, cy = place_image(872, 1440)
    assert (y, cy) == (0, 6858000)
    assert cx == 6858000 * 872 // 1440
    assert x == 1567543 + (10624457 - cx) // 2


def test_hymn_page_files_order_and_mixed_ext():
    names = ['101-2.png', '101-3.jpg', '101-1.png', '1010.png', '128.png', '58-1.PNG']
    assert hymn_page_files(names, 101) == ['101-1.png', '101-2.png', '101-3.jpg']
    assert hymn_page_files(names, 128) == ['128.png']
    assert hymn_page_files(names, 58) == ['58-1.PNG']
    assert hymn_page_files(names, 999) == []


def assert_package_ok(ppt_bytes):
    pkg = Package(ppt_bytes)
    assert pkg.order[0] == CONTENT_TYPES
    ct = pkg.xml(CONTENT_TYPES)
    overrides = {o.get('PartName')[1:] for o in ct.findall('ct:Override', NS)}
    defaults = {d.get('Extension').lower() for d in ct.findall('ct:Default', NS)}
    used = set()
    for name in pkg.parts:
        if name.endswith('.rels'):
            src = name.replace('_rels/', '').removesuffix('.rels')
            for _, _, target, ext in pkg.rels(src):
                if not ext:
                    assert target in pkg.parts, f'{name} -> {target} 없음'
                    used.add(target)
    for name in pkg.parts:
        if name.startswith('ppt/slides/slide'):
            assert name in overrides, name
        if name.startswith('ppt/media/'):
            assert name in used, f'고아 media {name}'
            assert posixpath.splitext(name)[1][1:].lower() in defaults
    for o in overrides:
        assert o in pkg.parts, f'Override 대상 없음 {o}'
    pres = pkg.xml('ppt/presentation.xml')
    ids = [s.get('id') for s in pres.find('p:sldIdLst', NS)]
    sec = [s.get('id') for s in pres.iter('{http://schemas.microsoft.com/office/powerpoint/2010/main}sldId')]
    assert sorted(sec) == sorted(ids)
    assert len(Presentation(io.BytesIO(ppt_bytes)).slides) == len(ids)


def test_full_replacement(ppt_1010, get_images):
    out, report = replace_hymns(ppt_1010, FULL, get_images)
    assert report['slide_count'] == 62
    assert report['warnings'] == []
    assert numbers(out) == [58, 58, 106, 106, 128, 129, 129, 61, 61]
    assert_package_ok(out)

    src, dst = Package(ppt_1010), Package(out)
    src_parts, dst_parts = src.slide_parts(), dst.slide_parts()
    hymn_src = {i for b in find_hymn_bundles(src) for i in b}
    hymn_dst = {i for b in find_hymn_bundles(dst) for i in b}
    keep_src = [p for i, p in enumerate(src_parts, 1) if i not in hymn_src]
    keep_dst = [p for i, p in enumerate(dst_parts, 1) if i not in hymn_dst]
    assert len(keep_src) == len(keep_dst) == 53
    for a, b in zip(keep_src, keep_dst):
        assert src.read(a) == dst.read(b)


def test_inserted_slide_layout_and_image_untouched(ppt_1010, get_images):
    out, _ = replace_hymns(ppt_1010, FULL, get_images)
    pkg = Package(out)
    part = pkg.slide_parts()[9]
    root = pkg.xml(part)
    pic = root.find('.//p:pic', NS)
    off, ext = pic.find('.//a:off', NS), pic.find('.//a:ext', NS)
    assert (off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy')) == ('1567543', '330200', '10624457', '6197599')
    assert pic.find('.//a:srcRect', NS) is None
    sp = root.find('.//p:sp', NS)
    sp_off, sp_ext = sp.find('.//a:off', NS), sp.find('.//a:ext', NS)
    assert (sp_off.get('x'), sp_off.get('y'), sp_ext.get('cx'), sp_ext.get('cy')) == ('0', '0', '1567543', '1015663')
    rpr = sp.find('.//a:rPr', NS)
    assert rpr.get('sz') == '6000' and rpr.get('b') == '1'
    assert rpr.find('.//a:outerShdw', NS) is not None
    embed = pic.find('.//a:blip', NS).get(qn('r:embed'))
    target = next(t for rid, _, t, _ in pkg.rels(part) if rid == embed)
    assert pkg.read(target) == get_images(58)[0]


@pytest.mark.parametrize('communion, count, nums', [
    ([128], 60, [58, 58, 106, 106, 128, 61, 61]),
    ([129], 61, [58, 58, 106, 106, 129, 129, 61, 61]),
])
def test_single_communion_hymn(ppt_1010, get_images, communion, count, nums):
    out, report = replace_hymns(ppt_1010, dict(FULL, communion=communion), get_images)
    assert report['slide_count'] == count
    assert numbers(out) == nums
    assert_package_ok(out)


def test_missing_image_is_skipped_with_warning(ppt_1010, get_images):
    out, report = replace_hymns(ppt_1010, dict(FULL, communion=[128, 999]), get_images)
    assert any('999' in w for w in report['warnings'])
    assert numbers(out) == [58, 58, 106, 106, 128, 61, 61]
    assert_package_ok(out)


def _move_slide(ppt_bytes, from_no, after_no):
    pkg = Package(ppt_bytes)
    pres = pkg.xml('ppt/presentation.xml')
    lst = pres.find('p:sldIdLst', NS)
    items = list(lst)
    el, anchor = items[from_no - 1], items[after_no - 1]
    lst.remove(el)
    anchor.addnext(el)
    pkg.write_xml('ppt/presentation.xml', pres)
    return pkg.save()


def test_communion_antiphon_is_moved_back(ppt_1010, get_images):
    moved = _move_slide(ppt_1010, 56, 54)  # 영성체송을 기도문(55) 앞으로
    out, report = replace_hymns(moved, FULL, get_images)
    assert any('영성체송 슬라이드를' in w for w in report['warnings'])
    pkg = Package(out)
    titles = [''.join(t.text or '' for t in pkg.xml(p).iter(qn('a:t'))) for p in pkg.slide_parts()]
    i = next(i for i, t in enumerate(titles) if '합당치 않사오나' in t)
    assert '영 성 체 송' in titles[i + 1]
    assert numbers(out) == [58, 58, 106, 106, 128, 129, 129, 61, 61]
    assert_package_ok(out)


def test_wrong_bundle_count_raises(ppt_1010, get_images):
    broken = _move_slide(ppt_1010, 61, 40)  # 파견 묶음을 쪼개 묶음 5개로
    with pytest.raises(HymnTemplateError):
        replace_hymns(broken, FULL, get_images)
