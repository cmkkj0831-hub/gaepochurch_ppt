import re

import pytest
import requests

from mass_ppt.storage import (AUTH_MESSAGE, KEEP_ARCHIVES, GitHubAuthError, GitHubStorage, LocalStorage,
                              StorageError)


class FakeResponse:
    def __init__(self, status, body=None, headers=None, content=b''):
        self.status_code, self._body, self.headers = status, body, headers or {}
        self.content = content

    def json(self):
        return self._body


class FakeGitHub:
    """Releases·Git Data API 일부를 흉내 내는 가짜 서버."""

    def __init__(self):
        self.assets = {}
        self.next_id = 1
        self.release = None
        self.calls = []
        self.fail_ref_updates = 0
        self.head = 'c0'
        self.files = {'songs/1-1.png': 'b0'}

    def request(self, method, url, params=None, json=None, data=None, headers=None, timeout=None):
        self.calls.append((method, url))
        path = re.sub(r'^https://(api|uploads)\.github\.com/repos/[^/]+/[^/]+', '', url)
        if method == 'GET' and path == '/releases/tags/archive':
            return FakeResponse(200, self.release) if self.release else FakeResponse(404)
        if method == 'POST' and path == '/releases':
            self.release = {'id': 7}
            return FakeResponse(201, self.release)
        if method == 'GET' and path == '/releases/7/assets':
            return FakeResponse(200, list(self.assets.values()))
        if method == 'POST' and path == '/releases/7/assets':
            if params['name'] in [a['name'] for a in self.assets.values()]:
                return FakeResponse(422)
            aid = self.next_id
            self.next_id += 1
            self.assets[aid] = {'id': aid, 'name': params['name'], 'label': params.get('label'),
                                'size': len(data), 'updated_at': '2026-10-20T00:00:00Z', 'data': data}
            return FakeResponse(201, self.assets[aid])
        m = re.fullmatch(r'/releases/assets/(\d+)', path)
        if m:
            aid = int(m.group(1))
            if method == 'DELETE':
                self.assets.pop(aid, None)
                return FakeResponse(204)
            return FakeResponse(200, content=self.assets[aid]['data'])
        if method == 'GET' and path == '/git/ref/heads/main':
            return FakeResponse(200, {'object': {'sha': self.head}})
        if method == 'GET' and path.startswith('/git/commits/'):
            return FakeResponse(200, {'tree': {'sha': 't-' + path.split('/')[-1]}})
        if method == 'GET' and path.startswith('/git/trees/'):
            return FakeResponse(200, {'tree': [{'path': p, 'type': 'blob', 'sha': s, 'size': 1}
                                               for p, s in self.files.items()]})
        if method == 'POST' and path == '/git/blobs':
            return FakeResponse(201, {'sha': 'blob-new'})
        if method == 'POST' and path == '/git/trees':
            self.last_tree = json
            return FakeResponse(201, {'sha': 'tree-new'})
        if method == 'POST' and path == '/git/commits':
            self.last_commit = json
            return FakeResponse(201, {'sha': 'commit-new'})
        if method == 'PATCH' and path == '/git/refs/heads/main':
            if self.fail_ref_updates:
                self.fail_ref_updates -= 1
                self.head = 'c-other'
                return FakeResponse(422)
            self.head = json['sha']
            return FakeResponse(200, {})
        raise AssertionError(f'unexpected {method} {url}')


@pytest.fixture
def gh():
    fake = FakeGitHub()
    store = GitHubStorage('x', 'o/r', ttl=0)
    store.session.request = fake.request
    return store, fake


@pytest.mark.parametrize('status, headers, auth', [
    (401, {}, True),
    (403, {}, True),
    (403, {'x-ratelimit-remaining': '0'}, False),
])
def test_auth_errors_are_raised_not_hidden(gh, status, headers, auth):
    store, _ = gh
    store.session.request = lambda *a, **k: FakeResponse(status, {}, headers)
    with pytest.raises(StorageError) as ei:
        store.list_archive()
    assert isinstance(ei.value, GitHubAuthError) == auth
    if auth:
        assert str(ei.value) == AUTH_MESSAGE
    with pytest.raises(StorageError):
        store.list_songs()


def test_network_error_is_korean_storage_error(gh):
    store, _ = gh

    def boom(*a, **k):
        raise requests.ConnectionError('down')
    store.session.request = boom
    with pytest.raises(StorageError, match='GitHub에 연결할 수 없습니다'):
        store.list_archive()


def test_no_release_means_empty_list(gh):
    store, _ = gh
    assert store.list_archive() == []


def test_save_names_label_replace_and_list(gh):
    store, fake = gh
    store.save_archive('261017미사.pptx', b'v1', '김선생')
    store.save_archive('261017미사.pptx', b'v2', None)
    assert [a['name'] for a in fake.assets.values()] == ['261017_missa.pptx']
    files = store.list_archive()
    assert [(f.name, f.date, f.uploader) for f in files] == [('261017미사.pptx', '261017', None)]
    assert store.read_archive('261017미사.pptx') == b'v2'
    assert store.get_latest_archive() == ('261017미사.pptx', b'v2')


def test_uploader_label_shown(gh):
    store, _ = gh
    store.save_archive('261024미사.pptx', b'x', '김선생')
    assert store.list_archive()[0].uploader == '김선생'


def test_keeps_latest_12_by_filename_date(gh):
    store, fake = gh
    dates = [f'26{m:02d}{d:02d}' for m, d in [(1, 3), (1, 10), (1, 17), (1, 24), (1, 31), (2, 7), (2, 14),
                                              (2, 21), (2, 28), (3, 7), (3, 14), (3, 21), (3, 28), (4, 4)]]
    for d in reversed(dates):
        store.save_archive(f'{d}미사.pptx', d.encode())
    kept = [f.date for f in store.list_archive()]
    assert len(kept) == KEEP_ARCHIVES
    assert kept == sorted(dates, reverse=True)[:KEEP_ARCHIVES]


def test_bad_archive_name_rejected(gh):
    store, _ = gh
    with pytest.raises(StorageError):
        store.save_archive('미사.pptx', b'x')


def test_song_commit_retries_from_new_head(gh):
    store, fake = gh
    fake.fail_ref_updates = 1
    store.change_songs({'999-1.png': b'a'}, ['1-1.png'], '성가 교체: 999')
    assert fake.head == 'commit-new'
    assert fake.last_commit['parents'] == ['c-other']
    entries = {e['path']: e['sha'] for e in fake.last_tree['tree']}
    assert entries == {'songs/999-1.png': 'blob-new', 'songs/1-1.png': None}
    assert fake.last_tree['base_tree'] == 't-c-other'


def test_song_name_rules(gh):
    store, _ = gh
    with pytest.raises(StorageError):
        store.change_songs({'../evil.png': b'a'}, [], 'x')
    with pytest.raises(StorageError):
        store.change_songs({'abc.png': b'a'}, [], 'x')


def test_local_change_songs(tmp_path):
    store = LocalStorage(tmp_path / 'a', tmp_path / 's')
    store.change_songs({'999-1.png': b'1', '999-2.png': b'2'}, [], 'add')
    assert store.hymn_files(999) == ['999-1.png', '999-2.png']
    store.change_songs({'999.png': b'x'}, store.hymn_files(999), 'replace')
    assert store.hymn_files(999) == ['999.png']
    assert store.get_hymn_images(999) == [b'x']
    store.change_songs({}, ['999.png'], 'delete')
    assert store.hymn_files(999) == []
