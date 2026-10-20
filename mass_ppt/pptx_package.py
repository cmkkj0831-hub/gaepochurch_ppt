"""PPTX(zip) 패키지를 메모리에서 읽고 원래 구조를 유지한 채 다시 쓴다."""
import io
import posixpath
import zipfile

from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
}
CONTENT_TYPES = '[Content_Types].xml'
REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'


def qn(tag: str) -> str:
    prefix, local = tag.split(':')
    return '{%s}%s' % (NS[prefix], local)


def rels_path(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, '_rels', f + '.rels')


def resolve(source_part: str, target: str) -> str:
    if target.startswith('/'):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relative(source_part: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(source_part))


def parse_xml(data: bytes):
    return etree.fromstring(data)


def to_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


class Package:
    def __init__(self, data: bytes):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            self.infos = {i.filename: i for i in z.infolist()}
            self.order = [i.filename for i in z.infolist()]
            self.parts = {n: z.read(n) for n in self.order}

    def read(self, name: str) -> bytes:
        return self.parts[name]

    def xml(self, name: str):
        return parse_xml(self.parts[name])

    def write(self, name: str, data: bytes):
        if name not in self.parts:
            self.order.append(name)
        self.parts[name] = data

    def write_xml(self, name: str, root):
        self.write(name, to_xml(root))

    def delete(self, name: str):
        if name in self.parts:
            del self.parts[name]
            self.order.remove(name)
            self.infos.pop(name, None)

    def rels(self, part: str):
        """part의 관계 목록: [(rId, type, 절대 경로 또는 외부 URL, external)]"""
        rp = rels_path(part)
        if rp not in self.parts:
            return []
        out = []
        for r in self.xml(rp):
            ext = r.get('TargetMode') == 'External'
            target = r.get('Target') if ext else resolve(part, r.get('Target'))
            out.append((r.get('Id'), r.get('Type'), target, ext))
        return out

    def slide_parts(self) -> list:
        """presentation.xml의 sldIdLst 순서대로 슬라이드 part 경로."""
        pres = self.xml('ppt/presentation.xml')
        rmap = {rid: t for rid, _, t, _ in self.rels('ppt/presentation.xml')}
        lst = pres.find('p:sldIdLst', NS)
        return [rmap[s.get(qn('r:id'))] for s in lst] if lst is not None else []

    def save(self) -> bytes:
        names = [n for n in self.order if n in self.parts]
        if CONTENT_TYPES in names:
            names.remove(CONTENT_TYPES)
            names.insert(0, CONTENT_TYPES)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w') as z:
            for n in names:
                old = self.infos.get(n)
                info = zipfile.ZipInfo(n, date_time=old.date_time if old else (1980, 1, 1, 0, 0, 0))
                info.compress_type = old.compress_type if old else zipfile.ZIP_DEFLATED
                info.external_attr = old.external_attr if old else 0o600 << 16
                z.writestr(info, self.parts[n])
        return buf.getvalue()


def shape_text(sp) -> str:
    return ''.join(t.text or '' for t in sp.iter(qn('a:t')))


def compact(text: str) -> str:
    return ''.join(text.split())


def sp_tree_children(slide_root) -> list:
    tree = slide_root.find('p:cSld/p:spTree', NS)
    return [c for c in tree if etree.QName(c).localname not in ('nvGrpSpPr', 'grpSpPr', 'extLst')]


def slide_title(slide_root, titles) -> str:
    """슬라이드 안 sp 중 공백 제거 텍스트가 titles 중 하나와 정확히 같은 것을 반환."""
    for sp in slide_root.iter(qn('p:sp')):
        t = compact(shape_text(sp))
        if t in titles:
            return t
    return ''
