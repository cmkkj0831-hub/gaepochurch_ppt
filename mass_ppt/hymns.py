"""성가 번호 슬라이드 묶음(입당/봉헌/성체/파견)을 새 성가 이미지로 교체한다."""
import copy
import io
import re
from typing import Callable, List

from lxml import etree
from PIL import Image

from .pptx_package import (
    CONTENT_TYPES, NS, REL_SLIDE, Package, compact, qn, relative, rels_path,
    shape_text, slide_title, sp_tree_children,
)

ROLES = ('entrance', 'offertory', 'communion', 'recessional')
ROLE_LABELS = {'entrance': '입당', 'offertory': '봉헌', 'communion': '성체', 'recessional': '파견'}
NUMBER_BOX_W = 1567543
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
REL_NOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
IMAGE_TYPES = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg'}


class HymnTemplateError(Exception):
    pass


def hymn_number(slide_root):
    """성가 번호 슬라이드(그림 1 + 숫자뿐인 텍스트 1, 그 외 도형 없음)면 번호, 아니면 None."""
    kids = sp_tree_children(slide_root)
    names = sorted(etree.QName(k).localname for k in kids)
    if names != ['pic', 'sp']:
        return None
    sp = next(k for k in kids if etree.QName(k).localname == 'sp')
    text = shape_text(sp).strip()
    return int(text) if text.isdigit() else None


def find_hymn_bundles(pkg: Package) -> List[List[int]]:
    """연속된 성가 번호 슬라이드 묶음 (슬라이드 번호는 1부터)."""
    bundles, cur = [], []
    for i, part in enumerate(pkg.slide_parts(), 1):
        if hymn_number(pkg.xml(part)) is not None:
            cur.append(i)
        elif cur:
            bundles.append(cur)
            cur = []
    if cur:
        bundles.append(cur)
    return bundles


def place_image(img_w: int, img_h: int, slide_w: int = 12192000, slide_h: int = 6858000):
    """이미지 영역(x=번호칸 오른쪽~끝, y=0~끝)에 비율 유지 최대 크기로 가운데 배치. (x, y, cx, cy)"""
    ax, ay, aw, ah = NUMBER_BOX_W, 0, slide_w - NUMBER_BOX_W, slide_h
    if img_w * ah >= img_h * aw:
        cx, cy = aw, aw * img_h // img_w
    else:
        cx, cy = ah * img_w // img_h, ah
    return ax + (aw - cx) // 2, ay + (ah - cy) // 2, cx, cy


def _has_text(root, needle):
    return any(needle in compact(shape_text(sp)) for sp in root.iter(qn('p:sp')))


class _Deck:
    """presentation.xml의 슬라이드 목록과 관련 part들을 함께 고친다."""

    def __init__(self, pkg: Package):
        self.pkg = pkg
        self.pres = pkg.xml('ppt/presentation.xml')
        self.pres_rels = pkg.xml(rels_path('ppt/presentation.xml'))
        self.ct = pkg.xml(CONTENT_TYPES)
        self.sld_lst = self.pres.find('p:sldIdLst', NS)
        rmap = {r.get('Id'): r.get('Target') for r in self.pres_rels}
        self.part_of = {s.get('id'): 'ppt/' + rmap[s.get(qn('r:id'))] for s in self.sld_lst}
        self.replaced = {}
        self.deleted = set()

    def ids(self):
        return [s.get('id') for s in self.sld_lst]

    def parts(self):
        return [self.part_of[i] for i in self.ids()]

    def move_after(self, sid, after_sid):
        el = next(s for s in self.sld_lst if s.get('id') == sid)
        anchor = next(s for s in self.sld_lst if s.get('id') == after_sid)
        self.sld_lst.remove(el)
        anchor.addnext(el)

    def _next_slide_name(self):
        nums = [int(m.group(1)) for n in self.pkg.parts
                if (m := re.fullmatch(r'ppt/slides/slide(\d+)\.xml', n))]
        return f'ppt/slides/slide{max(nums) + 1}.xml'

    def _next_media_name(self, ext):
        n = 1
        while any(p.startswith(f'ppt/media/hymn{n}.') for p in self.pkg.parts):
            n += 1
        return f'ppt/media/hymn{n}.{ext}'

    def _next_rid(self):
        used = {r.get('Id') for r in self.pres_rels}
        n = 1
        while f'rId{n}' in used:
            n += 1
        return f'rId{n}'

    def add_slide_like(self, template_part, number, image_bytes, after_el):
        pkg = self.pkg
        with Image.open(io.BytesIO(image_bytes)) as im:
            fmt, (w, h) = (im.format or '').lower(), im.size
        ext = {'png': 'png', 'jpeg': 'jpg'}.get(fmt)
        if ext is None:
            raise HymnTemplateError(f'{number}번 성가 이미지 형식({fmt or "알 수 없음"})은 지원하지 않습니다. png/jpg만 가능합니다.')
        media = self._next_media_name(ext)
        pkg.write(media, image_bytes)

        slide = pkg.xml(template_part)
        kids = sp_tree_children(slide)
        sp = next(k for k in kids if etree.QName(k).localname == 'sp')
        pic = next(k for k in kids if etree.QName(k).localname == 'pic')
        runs = [r for r in sp.iter(qn('a:r'))]
        runs[0].find('a:t', NS).text = str(number)
        for r in runs[1:]:
            r.find('a:t', NS).text = ''

        sld_sz = self.pres.find('p:sldSz', NS)
        x, y, cx, cy = place_image(w, h, int(sld_sz.get('cx')), int(sld_sz.get('cy')))
        xfrm = pic.find('p:spPr/a:xfrm', NS)
        xfrm.find('a:off', NS).attrib.update({'x': str(x), 'y': str(y)})
        xfrm.find('a:ext', NS).attrib.update({'cx': str(cx), 'cy': str(cy)})
        blip_fill = pic.find('p:blipFill', NS)
        for src in blip_fill.findall('a:srcRect', NS):
            blip_fill.remove(src)
        if blip_fill.find('a:stretch', NS) is None:
            etree.SubElement(blip_fill, qn('a:stretch'))
        stretch = blip_fill.find('a:stretch', NS)
        if stretch.find('a:fillRect', NS) is None:
            etree.SubElement(stretch, qn('a:fillRect'))
        embed = blip_fill.find('a:blip', NS).get(qn('r:embed'))

        new_part = self._next_slide_name()
        rels = pkg.xml(rels_path(template_part))
        for r in list(rels):
            if r.get('Type') == REL_NOTES:
                rels.remove(r)
            elif r.get('Id') == embed:
                r.set('Target', relative(new_part, media))
        pkg.write_xml(new_part, slide)
        pkg.write_xml(rels_path(new_part), rels)

        ctype = next(o.get('ContentType') for o in self.ct.findall('ct:Override', NS)
                     if o.get('PartName') == '/' + template_part)
        etree.SubElement(self.ct, qn('ct:Override'), PartName='/' + new_part, ContentType=ctype)

        rid = self._next_rid()
        etree.SubElement(self.pres_rels, '{%s}Relationship' % NS['rel'],
                         Id=rid, Type=REL_SLIDE, Target=relative('ppt/presentation.xml', new_part))
        new_id = str(max(int(i) for i in list(self.part_of) + self.ids()) + 1)
        el = etree.Element(qn('p:sldId'), id=new_id)
        el.set(qn('r:id'), rid)
        if after_el is None:
            self.sld_lst.insert(0, el)
        else:
            after_el.addnext(el)
        self.part_of[new_id] = new_part
        return el

    def delete_slide(self, sid):
        pkg = self.pkg
        el = next(s for s in self.sld_lst if s.get('id') == sid)
        rid = el.get(qn('r:id'))
        self.sld_lst.remove(el)
        for r in list(self.pres_rels):
            if r.get('Id') == rid:
                self.pres_rels.remove(r)
        part = self.part_of[sid]
        for _, rtype, target, ext in pkg.rels(part):
            if rtype == REL_NOTES and not ext:
                pkg.delete(rels_path(target))
                pkg.delete(target)
                self._drop_override(target)
        pkg.delete(rels_path(part))
        pkg.delete(part)
        self._drop_override(part)
        self.deleted.add(sid)

    def _drop_override(self, part):
        for o in self.ct.findall('ct:Override', NS):
            if o.get('PartName') == '/' + part:
                self.ct.remove(o)

    def _sync_sections(self):
        order = {sid: i for i, sid in enumerate(self.ids())}
        for sec_lst in self.pres.iter('{%s}sldIdLst' % P14):
            new_ids = []
            for s in sec_lst:
                sid = s.get('id')
                new_ids.extend(self.replaced.get(sid, []))
                if sid not in self.deleted:
                    new_ids.append(sid)
            new_ids.sort(key=lambda i: order.get(i, len(order)))
            for s in list(sec_lst):
                sec_lst.remove(s)
            for sid in new_ids:
                etree.SubElement(sec_lst, '{%s}sldId' % P14, id=sid)

    def _ensure_image_defaults(self):
        have = {d.get('Extension').lower() for d in self.ct.findall('ct:Default', NS)}
        for ext, ctype in IMAGE_TYPES.items():
            if ext not in have:
                d = etree.Element(qn('ct:Default'), Extension=ext, ContentType=ctype)
                self.ct.insert(0, d)

    def _drop_orphan_media(self):
        used = set()
        for name in list(self.pkg.parts):
            if name.endswith('.rels'):
                src = name.replace('_rels/', '').removesuffix('.rels')
                for _, _, target, ext in self.pkg.rels(src):
                    if not ext:
                        used.add(target)
        for name in list(self.pkg.parts):
            if name.startswith('ppt/media/') and name not in used:
                self.pkg.delete(name)

    def _update_app_counts(self):
        name = 'docProps/app.xml'
        if name not in self.pkg.parts:
            return
        app = self.pkg.xml(name)
        slides = app.find('ep:Slides', NS)
        if slides is not None:
            slides.text = str(len(self.ids()))
        notes = app.find('ep:Notes', NS)
        if notes is not None:
            notes.text = str(sum(1 for n in self.pkg.parts
                                 if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml', n)))
        self.pkg.write_xml(name, app)

    def finish(self):
        self._sync_sections()
        self._ensure_image_defaults()
        self.pkg.write_xml('ppt/presentation.xml', self.pres)
        self.pkg.write_xml(rels_path('ppt/presentation.xml'), self.pres_rels)
        self.pkg.write_xml(CONTENT_TYPES, self.ct)
        self._drop_orphan_media()
        self._update_app_counts()


def _fix_communion_order(deck: _Deck, warnings: list):
    parts = deck.parts()
    roots = [deck.pkg.xml(p) for p in parts]
    prayer = [i for i, r in enumerate(roots) if _has_text(r, '합당치않사오나')]
    antiphon = [i for i, r in enumerate(roots) if slide_title(r, ('영성체송',))]
    if len(prayer) != 1 or len(antiphon) != 1:
        raise HymnTemplateError(
            f'영성체 기도문 슬라이드({len(prayer)}개)나 영성체송 슬라이드({len(antiphon)}개)를 '
            '하나씩 찾지 못했습니다. 기준 PPT를 확인해 주세요.')
    if antiphon[0] != prayer[0] + 1:
        ids = deck.ids()
        deck.move_after(ids[antiphon[0]], ids[prayer[0]])
        warnings.append('영성체송 슬라이드를 영성체 기도문("합당치 않사오나") 바로 다음으로 옮겼습니다.')


def _bundles_by_ids(deck: _Deck):
    bundles, cur = [], []
    for sid in deck.ids():
        if hymn_number(deck.pkg.xml(deck.part_of[sid])) is not None:
            cur.append(sid)
        elif cur:
            bundles.append(cur)
            cur = []
    if cur:
        bundles.append(cur)
    return bundles


def _validate_bundles(deck: _Deck, bundles):
    if len(bundles) != 4:
        raise HymnTemplateError(
            f'성가 번호 슬라이드 묶음이 {len(bundles)}개입니다. 입당/봉헌/성체/파견 4개여야 합니다. '
            '기준 PPT의 성가 슬라이드(그림 1장 + 번호 칸)를 확인해 주세요.')
    ids = deck.ids()
    after_offertory = ids.index(bundles[1][-1]) + 1
    if after_offertory >= len(ids) or not slide_title(
            deck.pkg.xml(deck.part_of[ids[after_offertory]]), ('예물준비기도',)):
        raise HymnTemplateError('봉헌 성가 묶음 바로 다음에 "예물 준비 기도" 슬라이드가 없습니다.')
    before_communion = ids.index(bundles[2][0]) - 1
    if before_communion < 0 or not slide_title(
            deck.pkg.xml(deck.part_of[ids[before_communion]]), ('영성체송',)):
        raise HymnTemplateError('성체 성가 묶음 바로 앞에 영성체송 슬라이드가 없습니다.')


def _role_numbers(hymn_numbers: dict, role: str) -> list:
    v = hymn_numbers.get(role)
    if v is None:
        return []
    vals = v if isinstance(v, (list, tuple)) else [v]
    out = []
    for x in vals:
        if x is None or str(x).strip() in ('', '0'):
            continue
        out.append(int(str(x).strip()))
    return out


def replace_hymns(ppt_bytes: bytes, hymn_numbers: dict, get_hymn_images: Callable[[int], list]):
    """hymn_numbers = {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}"""
    pkg = Package(ppt_bytes)
    deck = _Deck(pkg)
    report = {'hymns': [], 'warnings': []}

    _fix_communion_order(deck, report['warnings'])
    bundles = _bundles_by_ids(deck)
    _validate_bundles(deck, bundles)

    for role, bundle in zip(ROLES, bundles):
        template_part = deck.part_of[bundle[0]]
        template_xml = pkg.read(template_part)
        template_rels = pkg.read(rels_path(template_part))
        prev = bundle[0]
        prev_idx = deck.ids().index(prev)
        anchor = deck.sld_lst[prev_idx - 1] if prev_idx > 0 else None

        tmp_part = 'ppt/slides/_hymn_template.xml'
        pkg.write(tmp_part, template_xml)
        pkg.write(rels_path(tmp_part), template_rels)
        etree.SubElement(deck.ct, qn('ct:Override'), PartName='/' + tmp_part,
                         ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')

        for sid in bundle:
            deck.delete_slide(sid)

        new_ids = []
        for num in _role_numbers(hymn_numbers, role):
            images = get_hymn_images(num) or []
            if not images:
                report['warnings'].append(f'{ROLE_LABELS[role]} {num}번: 성가 이미지가 없어 넣지 않았습니다.')
                continue
            for img in images:
                anchor = deck.add_slide_like(tmp_part, num, img, anchor)
                new_ids.append(anchor.get('id'))
            report['hymns'].append({'role': ROLE_LABELS[role], 'num': num, 'pages': len(images)})
        deck.replaced[bundle[0]] = new_ids

        pkg.delete(tmp_part)
        pkg.delete(rels_path(tmp_part))
        deck._drop_override(tmp_part)

    deck.finish()
    report['slide_count'] = len(deck.ids())
    return pkg.save(), report
