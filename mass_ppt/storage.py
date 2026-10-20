"""
최종본 PPT와 성가 이미지 저장소.
- LocalStorage: 로컬 archive/, songs/ 폴더
- GitHubStorage: 최종본은 GitHub Releases(태그 archive) 첨부파일, 성가는 저장소 songs/ 폴더
"""
import base64
import re
import threading
import time
from abc import ABC, abstractmethod
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import requests

ARCHIVE_NAME_RE = re.compile(r'^(\d{6})미사\.pptx$')
SONG_NAME_RE = re.compile(r'^(\d+)(?:-(\d+))?\.(png|jpe?g)$', re.IGNORECASE)
KEEP_ARCHIVES = 12
PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'


class StorageError(Exception):
    """화면에 그대로 보여줄 수 있는 한국어 메시지를 담는다."""


class GitHubAuthError(StorageError):
    pass


AUTH_MESSAGE = ('GitHub 연결 오류: 토큰이 만료되었거나 권한이 없습니다. '
                '관리자에게 토큰 갱신을 요청하세요.')


@dataclass
class FileInfo:
    name: str
    size: int
    modified_time: Optional[datetime] = None
    date: Optional[str] = None
    uploader: Optional[str] = None


def archive_date(filename: str) -> Optional[str]:
    m = ARCHIVE_NAME_RE.match(filename or '')
    return m.group(1) if m else None


def hymn_page_files(names: List[str], hymn_num: int) -> List[str]:
    """파일명 목록에서 성가 번호의 이미지들을 쪽 순서로. 쪽 번호 파일이 있으면 그것만 사용."""
    single, pages = [], []
    for name in names:
        m = SONG_NAME_RE.match(name)
        if not m or int(m.group(1)) != hymn_num:
            continue
        if m.group(2):
            pages.append((int(m.group(2)), name))
        else:
            single.append(name)
    if pages:
        return [n for _, n in sorted(pages)]
    return single[:1]


class Storage(ABC):
    @abstractmethod
    def list_archive(self) -> List[FileInfo]:
        """최종본 목록, 최신 날짜 순."""

    @abstractmethod
    def read_archive(self, filename: str) -> Optional[bytes]:
        pass

    @abstractmethod
    def save_archive(self, filename: str, ppt_bytes: bytes, uploader: Optional[str] = None) -> None:
        """실패하면 StorageError."""

    @abstractmethod
    def delete_archive(self, filename: str) -> None:
        pass

    @abstractmethod
    def list_songs(self) -> List[FileInfo]:
        pass

    @abstractmethod
    def read_song_file(self, filename: str) -> Optional[bytes]:
        pass

    @abstractmethod
    def change_songs(self, add: Dict[str, bytes], delete: Iterable[str], message: str) -> None:
        """songs/ 파일 추가·교체·삭제를 한 번에. 실패하면 StorageError."""

    def clear_cache(self):
        pass

    def get_latest_archive(self) -> Optional[Tuple[str, bytes]]:
        files = self.list_archive()
        if not files:
            return None
        data = self.read_archive(files[0].name)
        return (files[0].name, data) if data else None

    def hymn_files(self, hymn_num: int) -> List[str]:
        return hymn_page_files([f.name for f in self.list_songs()], hymn_num)

    def get_hymn_images(self, hymn_num: int) -> List[bytes]:
        """해당 번호의 모든 장을 순서대로 (num.png 또는 num-1, num-2 …)."""
        out = []
        for name in self.hymn_files(hymn_num):
            data = self.read_song_file(name)
            if data:
                out.append(data)
        return out


def _check_archive_name(filename: str):
    if not archive_date(filename):
        raise StorageError(f'파일명 "{filename}"이 YYMMDD미사.pptx 형식이 아닙니다.')


def _check_song_name(filename: str):
    if Path(filename).name != filename or not SONG_NAME_RE.match(filename):
        raise StorageError(f'성가 파일명 "{filename}"이 규칙(58.png, 58-1.png)에 맞지 않습니다.')


class LocalStorage(Storage):
    def __init__(self, archive_dir: Path = None, songs_dir: Path = None):
        root = Path(__file__).parent.parent
        self.archive_dir = Path(archive_dir or root / 'archive')
        self.songs_dir = Path(songs_dir or root / 'songs')
        self.archive_dir.mkdir(exist_ok=True)
        self.songs_dir.mkdir(exist_ok=True)

    def list_archive(self) -> List[FileInfo]:
        files = []
        for p in self.archive_dir.glob('*.pptx'):
            d = archive_date(p.name)
            if d:
                st = p.stat()
                files.append(FileInfo(p.name, st.st_size, datetime.fromtimestamp(st.st_mtime), d))
        return sorted(files, key=lambda f: f.date, reverse=True)

    def read_archive(self, filename: str) -> Optional[bytes]:
        if not archive_date(filename):
            return None
        try:
            return (self.archive_dir / filename).read_bytes()
        except OSError:
            return None

    def save_archive(self, filename: str, ppt_bytes: bytes, uploader: Optional[str] = None) -> None:
        _check_archive_name(filename)
        (self.archive_dir / filename).write_bytes(ppt_bytes)

    def delete_archive(self, filename: str) -> None:
        _check_archive_name(filename)
        (self.archive_dir / filename).unlink(missing_ok=True)

    def list_songs(self) -> List[FileInfo]:
        files = []
        for p in self.songs_dir.iterdir():
            if p.is_file() and p.suffix.lower() in ('.png', '.jpg', '.jpeg'):
                st = p.stat()
                files.append(FileInfo(p.name, st.st_size, datetime.fromtimestamp(st.st_mtime)))
        return sorted(files, key=lambda f: f.name)

    def read_song_file(self, filename: str) -> Optional[bytes]:
        try:
            return (self.songs_dir / Path(filename).name).read_bytes()
        except OSError:
            return None

    def change_songs(self, add: Dict[str, bytes], delete: Iterable[str], message: str) -> None:
        for name in list(delete):
            _check_song_name(name)
            (self.songs_dir / name).unlink(missing_ok=True)
        for name, data in add.items():
            _check_song_name(name)
            (self.songs_dir / name).write_bytes(data)


class _TTLCache:
    def __init__(self, ttl: float):
        self.ttl, self.data, self.lock = ttl, {}, threading.Lock()

    def get(self, key):
        with self.lock:
            hit = self.data.get(key)
            if hit and time.monotonic() - hit[0] < self.ttl:
                return hit[1]
        return None

    def put(self, key, value):
        with self.lock:
            self.data[key] = (time.monotonic(), value)

    def clear(self):
        with self.lock:
            self.data.clear()


class _LRUBytes:
    """내용이 바뀌지 않는 키(blob sha, asset id)용. 큰 PPT는 몇 개만 보관."""

    def __init__(self, max_items: int):
        self.max, self.data, self.lock = max_items, OrderedDict(), threading.Lock()

    def get(self, key):
        with self.lock:
            if key in self.data:
                self.data.move_to_end(key)
                return self.data[key]
        return None

    def put(self, key, value):
        with self.lock:
            self.data[key] = value
            self.data.move_to_end(key)
            while len(self.data) > self.max:
                self.data.popitem(last=False)


class GitHubStorage(Storage):
    API = 'https://api.github.com'
    UPLOADS = 'https://uploads.github.com'
    RELEASE_TAG = 'archive'

    def __init__(self, token: str, repo: str, branch: str = 'main', ttl: float = 30.0):
        self.repo, self.branch = repo, branch
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Bearer {token}',
            'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2022-11-28',
        })
        self.cache = _TTLCache(ttl)
        self.songs_bytes = _LRUBytes(200)
        self.archive_bytes = _LRUBytes(3)
        self.write_lock = threading.Lock()

    # ---- HTTP ----
    def _req(self, method, url, ok=(200, 201, 204), allow_404=False, **kw):
        if url.startswith('/'):
            url = f'{self.API}/repos/{self.repo}{url}'
        kw.setdefault('timeout', 60)
        try:
            resp = self.session.request(method, url, **kw)
        except requests.RequestException as e:
            raise StorageError(f'GitHub에 연결할 수 없습니다. 인터넷 연결을 확인하고 다시 시도해 주세요. ({e.__class__.__name__})') from e
        if resp.status_code == 401:
            raise GitHubAuthError(AUTH_MESSAGE)
        if resp.status_code == 403:
            if resp.headers.get('x-ratelimit-remaining') == '0':
                raise StorageError('GitHub 사용량 한도에 걸렸습니다. 잠시 후 다시 시도해 주세요.')
            raise GitHubAuthError(AUTH_MESSAGE)
        if resp.status_code == 404 and allow_404:
            return None
        if resp.status_code not in ok:
            raise StorageError(f'GitHub 요청이 실패했습니다 ({resp.status_code}). 잠시 후 다시 시도해 주세요.')
        return resp

    def clear_cache(self):
        self.cache.clear()

    # ---- 최종본: Releases ----
    @staticmethod
    def _asset_name(filename: str) -> str:
        return f'{archive_date(filename)}_missa.pptx'

    @staticmethod
    def _display_name(asset_name: str) -> Optional[str]:
        m = re.fullmatch(r'(\d{6})_missa\.pptx', asset_name)
        return f'{m.group(1)}미사.pptx' if m else None

    def _release(self, create=False):
        rel = self.cache.get('release')
        if rel:
            return rel
        resp = self._req('GET', f'/releases/tags/{self.RELEASE_TAG}', allow_404=True)
        if resp is None:
            if not create:
                return None
            resp = self._req('POST', '/releases', json={
                'tag_name': self.RELEASE_TAG, 'target_commitish': self.branch, 'name': '최종본 PPT',
                'body': f'매주 최종본 PPT를 보관합니다. 최근 {KEEP_ARCHIVES}주만 유지됩니다.'})
        rel = resp.json()
        self.cache.put('release', rel)
        return rel

    def _assets(self, fresh=False):
        if not fresh:
            hit = self.cache.get('assets')
            if hit is not None:
                return hit
        rel = self._release()
        assets = []
        if rel:
            page = 1
            while True:
                batch = self._req('GET', f'/releases/{rel["id"]}/assets',
                                  params={'per_page': 100, 'page': page}).json()
                assets.extend(batch)
                if len(batch) < 100:
                    break
                page += 1
        self.cache.put('assets', assets)
        return assets

    def list_archive(self) -> List[FileInfo]:
        files = []
        for a in self._assets():
            name = self._display_name(a['name'])
            if not name:
                continue
            label = a.get('label') or ''
            uploader = label.split(':', 1)[1].strip() if label.startswith('올린 사람:') else None
            modified = datetime.fromisoformat(a['updated_at'].replace('Z', '+00:00')) if a.get('updated_at') else None
            files.append(FileInfo(name, a.get('size', 0), modified, archive_date(name), uploader or None))
        return sorted(files, key=lambda f: f.date, reverse=True)

    def read_archive(self, filename: str) -> Optional[bytes]:
        if not archive_date(filename):
            return None
        asset = next((a for a in self._assets() if a['name'] == self._asset_name(filename)), None)
        if asset is None:
            return None
        hit = self.archive_bytes.get(asset['id'])
        if hit is not None:
            return hit
        data = self._req('GET', f'/releases/assets/{asset["id"]}',
                         headers={'Accept': 'application/octet-stream'}, timeout=180).content
        self.archive_bytes.put(asset['id'], data)
        return data

    def save_archive(self, filename: str, ppt_bytes: bytes, uploader: Optional[str] = None) -> None:
        _check_archive_name(filename)
        asset_name = self._asset_name(filename)
        params = {'name': asset_name}
        if uploader and uploader.strip():
            params['label'] = f'올린 사람: {uploader.strip()}'
        with self.write_lock:
            rel = self._release(create=True)
            for attempt in range(3):
                for a in self._assets(fresh=True):
                    if a['name'] == asset_name:
                        self._req('DELETE', f'/releases/assets/{a["id"]}', allow_404=True)
                resp = self._req('POST', f'{self.UPLOADS}/repos/{self.repo}/releases/{rel["id"]}/assets',
                                 ok=(201, 422), params=params, data=ppt_bytes, timeout=300,
                                 headers={'Content-Type': PPTX_MIME})
                if resp.status_code == 201:
                    break
                time.sleep(1 + attempt)
            else:
                raise StorageError('같은 이름의 최종본이 동시에 올라가고 있어 저장하지 못했습니다. 다시 시도해 주세요.')
            self._prune_archives()
            self.clear_cache()

    def _prune_archives(self):
        ours = [(self._display_name(a['name']), a) for a in self._assets(fresh=True)]
        ours = sorted([(archive_date(n), a) for n, a in ours if n], key=lambda x: x[0], reverse=True)
        for _, a in ours[KEEP_ARCHIVES:]:
            self._req('DELETE', f'/releases/assets/{a["id"]}', allow_404=True)

    def delete_archive(self, filename: str) -> None:
        _check_archive_name(filename)
        with self.write_lock:
            for a in self._assets(fresh=True):
                if a['name'] == self._asset_name(filename):
                    self._req('DELETE', f'/releases/assets/{a["id"]}', allow_404=True)
            self.clear_cache()

    # ---- 성가: 저장소 songs/ ----
    def _songs_tree(self):
        hit = self.cache.get('songs_tree')
        if hit is not None:
            return hit
        tree = self._req('GET', f'/git/trees/{self.branch}', params={'recursive': '1'}).json()
        if tree.get('truncated'):
            raise StorageError('저장소 파일이 너무 많아 성가 목록을 모두 읽지 못했습니다.')
        songs = {item['path'][len('songs/'):]: item for item in tree.get('tree', [])
                 if item['type'] == 'blob' and item['path'].startswith('songs/')
                 and '/' not in item['path'][len('songs/'):]}
        self.cache.put('songs_tree', songs)
        return songs

    def list_songs(self) -> List[FileInfo]:
        return sorted((FileInfo(n, it.get('size', 0)) for n, it in self._songs_tree().items()
                       if Path(n).suffix.lower() in ('.png', '.jpg', '.jpeg')), key=lambda f: f.name)

    def read_song_file(self, filename: str) -> Optional[bytes]:
        item = self._songs_tree().get(Path(filename).name)
        if item is None:
            return None
        hit = self.songs_bytes.get(item['sha'])
        if hit is not None:
            return hit
        data = self._req('GET', f'/git/blobs/{item["sha"]}',
                         headers={'Accept': 'application/vnd.github.raw+json'}).content
        self.songs_bytes.put(item['sha'], data)
        return data

    def change_songs(self, add: Dict[str, bytes], delete: Iterable[str], message: str) -> None:
        delete = [n for n in delete if n not in add]
        for name in list(add) + delete:
            _check_song_name(name)
        if not add and not delete:
            return
        with self.write_lock:
            blobs = {}
            for name, data in add.items():
                blobs[name] = self._req('POST', '/git/blobs', json={
                    'content': base64.b64encode(data).decode(), 'encoding': 'base64'}).json()['sha']
            for attempt in range(4):
                head = self._req('GET', f'/git/ref/heads/{self.branch}').json()['object']['sha']
                base_tree = self._req('GET', f'/git/commits/{head}').json()['tree']['sha']
                existing = {n for n in self._songs_tree_at(head)}
                entries = [{'path': f'songs/{n}', 'mode': '100644', 'type': 'blob', 'sha': s}
                           for n, s in blobs.items()]
                entries += [{'path': f'songs/{n}', 'mode': '100644', 'type': 'blob', 'sha': None}
                            for n in delete if n in existing]
                if not entries:
                    break
                tree = self._req('POST', '/git/trees', json={'base_tree': base_tree, 'tree': entries}).json()['sha']
                commit = self._req('POST', '/git/commits', json={
                    'message': message, 'tree': tree, 'parents': [head]}).json()['sha']
                resp = self._req('PATCH', f'/git/refs/heads/{self.branch}', ok=(200, 422),
                                 json={'sha': commit, 'force': False})
                if resp.status_code == 200:
                    break
                time.sleep(0.5 * (attempt + 1))
            else:
                raise StorageError('다른 사람이 동시에 성가를 바꾸고 있어 저장하지 못했습니다. 다시 시도해 주세요.')
            self.clear_cache()

    def _songs_tree_at(self, commit_sha):
        tree = self._req('GET', f'/git/trees/{commit_sha}', params={'recursive': '1'}).json()
        return {i['path'][len('songs/'):] for i in tree.get('tree', [])
                if i['type'] == 'blob' and i['path'].startswith('songs/')}
