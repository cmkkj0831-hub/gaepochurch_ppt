"""
PPT와 성가 이미지를 저장하고 불러오는 저장소 인터페이스.
로컬 및 GitHub 구현 제공.
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, List, Tuple, Dict
from dataclasses import dataclass
import re
from datetime import datetime
import base64
import requests
import time


@dataclass
class FileInfo:
    """저장된 파일 정보."""
    name: str
    size: int
    modified_time: datetime
    date: Optional[str] = None  # YYMMDD (PPT인 경우)
    uploader: Optional[str] = None  # 올린 사람 이름


class Storage(ABC):
    """저장소 인터페이스."""

    @abstractmethod
    def list_archive(self) -> List[FileInfo]:
        """최종본 PPT 목록 (날짜순 역순)."""
        pass

    @abstractmethod
    def get_latest_archive(self) -> Optional[Tuple[str, bytes]]:
        """가장 최근 최종본 (filename, bytes)."""
        pass

    @abstractmethod
    def read_archive(self, filename: str) -> Optional[bytes]:
        """최종본 읽기."""
        pass

    @abstractmethod
    def save_archive(self, filename: str, ppt_bytes: bytes, uploader: Optional[str] = None) -> bool:
        """최종본 저장 (YYMMDD미사.pptx 형식 검증)."""
        pass

    @abstractmethod
    def delete_archive(self, filename: str) -> bool:
        """최종본 삭제."""
        pass

    @abstractmethod
    def list_songs(self) -> List[FileInfo]:
        """성가 이미지 목록."""
        pass

    @abstractmethod
    def read_song(self, hymn_num: int, page: int = 1) -> Optional[bytes]:
        """성가 이미지 읽기 (hymn_num-page.png/jpg 또는 hymn_num.png)."""
        pass

    @abstractmethod
    def get_hymn_image_getter(self) -> callable:
        """generate.py에 전달할 이미지 getter 함수."""
        pass


class LocalStorage(Storage):
    """로컬 폴더 기반 저장소."""

    def __init__(self, archive_dir: Path = None, songs_dir: Path = None):
        """
        Args:
            archive_dir: 최종본 PPT 폴더 (기본: ./archive/)
            songs_dir: 성가 이미지 폴더 (기본: ./songs/)
        """
        if archive_dir is None:
            archive_dir = Path(__file__).parent.parent / 'archive'
        if songs_dir is None:
            songs_dir = Path(__file__).parent.parent / 'songs'

        self.archive_dir = Path(archive_dir)
        self.songs_dir = Path(songs_dir)

        self.archive_dir.mkdir(exist_ok=True)
        self.songs_dir.mkdir(exist_ok=True)

    def _parse_ppt_filename(self, filename: str) -> Optional[str]:
        """YYMMDD미사.pptx에서 YYMMDD 추출."""
        match = re.match(r'^(\d{6})미사\.pptx$', filename)
        return match.group(1) if match else None

    def _validate_ppt_filename(self, filename: str) -> bool:
        """파일명이 규칙을 따르는지 검증."""
        return self._parse_ppt_filename(filename) is not None

    def _date_key(self, date_str: str) -> int:
        """YYMMDD를 정렬 키로 변환 (역순)."""
        return -int(date_str)

    def list_archive(self) -> List[FileInfo]:
        """최종본 목록 (날짜순 역순)."""
        files = []
        for ppt_file in self.archive_dir.glob('*.pptx'):
            date_str = self._parse_ppt_filename(ppt_file.name)
            if date_str:
                stat = ppt_file.stat()
                files.append(FileInfo(
                    name=ppt_file.name,
                    size=stat.st_size,
                    modified_time=datetime.fromtimestamp(stat.st_mtime),
                    date=date_str
                ))

        # 날짜순 역순 정렬
        files.sort(key=lambda f: self._date_key(f.date))
        return files

    def get_latest_archive(self) -> Optional[Tuple[str, bytes]]:
        """가장 최근 최종본."""
        files = self.list_archive()
        if not files:
            return None

        latest = files[0]
        path = self.archive_dir / latest.name
        try:
            return latest.name, path.read_bytes()
        except (IOError, OSError):
            return None

    def read_archive(self, filename: str) -> Optional[bytes]:
        """최종본 읽기."""
        if not self._validate_ppt_filename(filename):
            return None

        path = self.archive_dir / filename
        try:
            return path.read_bytes()
        except (IOError, OSError):
            return None

    def save_archive(self, filename: str, ppt_bytes: bytes, uploader: Optional[str] = None) -> bool:
        """최종본 저장."""
        if not self._validate_ppt_filename(filename):
            return False

        path = self.archive_dir / filename
        try:
            path.write_bytes(ppt_bytes)
            return True
        except (IOError, OSError):
            return False

    def delete_archive(self, filename: str) -> bool:
        """최종본 삭제."""
        if not self._validate_ppt_filename(filename):
            return False

        path = self.archive_dir / filename
        try:
            path.unlink()
            return True
        except (IOError, OSError):
            return False

    def list_songs(self) -> List[FileInfo]:
        """성가 이미지 목록."""
        files = []
        for img_file in self.songs_dir.glob('*'):
            if img_file.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                stat = img_file.stat()
                files.append(FileInfo(
                    name=img_file.name,
                    size=stat.st_size,
                    modified_time=datetime.fromtimestamp(stat.st_mtime)
                ))

        return sorted(files, key=lambda f: f.name)

    def read_song(self, hymn_num: int, page: int = 1) -> Optional[bytes]:
        """성가 이미지 읽기."""
        # hymn_num-page.png/jpg 또는 hymn_num.png
        for ext in ['.png', '.jpg', '.jpeg']:
            # 여러 장
            if page > 1:
                path = self.songs_dir / f'{hymn_num}-{page}{ext}'
                if path.exists():
                    try:
                        return path.read_bytes()
                    except (IOError, OSError):
                        pass
            else:
                # 한 장짜리
                path = self.songs_dir / f'{hymn_num}{ext}'
                if path.exists():
                    try:
                        return path.read_bytes()
                    except (IOError, OSError):
                        pass

                # 여러 장의 첫 번째
                path = self.songs_dir / f'{hymn_num}-1{ext}'
                if path.exists():
                    try:
                        return path.read_bytes()
                    except (IOError, OSError):
                        pass

        return None

    def get_hymn_image_getter(self) -> callable:
        """generate.py에 전달할 이미지 getter."""
        def getter(hymn_num: int) -> Optional[bytes]:
            # 첫 번째 이미지 반환
            return self.read_song(hymn_num, page=1)

        return getter


class GitHubStorage(Storage):
    """GitHub 저장소 기반 저장소 (공개 저장소, Fine-grained token)."""

    def __init__(self, token: str, repo: str, branch: str = "main"):
        """
        Args:
            token: GitHub Fine-grained Personal Access Token (Contents r/w 권한)
            repo: 저장소 (예: "cmkkj0831-hub/gaepochurch_ppt")
            branch: 브랜치 (기본: main)
        """
        self.token = token
        self.repo = repo
        self.branch = branch
        self.base_url = f"https://api.github.com/repos/{repo}"
        self.headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/json"
        }

    def _get(self, endpoint: str, params: dict = None) -> Optional[dict]:
        """GET 요청."""
        try:
            resp = requests.get(
                f"{self.base_url}{endpoint}",
                headers=self.headers,
                params=params or {},
                timeout=10
            )
            resp.raise_for_status()
            return resp.json()
        except Exception:
            return None

    def _put(self, endpoint: str, data: dict) -> Optional[dict]:
        """PUT 요청."""
        try:
            resp = requests.put(
                f"{self.base_url}{endpoint}",
                headers=self.headers,
                json=data,
                timeout=10
            )
            resp.raise_for_status()
            return resp.json()
        except Exception:
            return None

    def _delete(self, endpoint: str, data: dict) -> Optional[dict]:
        """DELETE 요청."""
        try:
            resp = requests.delete(
                f"{self.base_url}{endpoint}",
                headers=self.headers,
                json=data,
                timeout=10
            )
            resp.raise_for_status()
            return resp.json()
        except Exception:
            return None

    def _get_file(self, path: str) -> Optional[dict]:
        """파일 메타정보 조회."""
        return self._get(f"/contents/{path}", {"ref": self.branch})

    def _put_file(self, path: str, content: bytes, message: str, sha: str = None) -> bool:
        """파일 저장 (Contents API - 1MB 이상은 실패 가능)."""
        data = {
            "message": message,
            "content": base64.b64encode(content).decode(),
            "branch": self.branch
        }
        if sha:
            data["sha"] = sha

        return self._put(f"/contents/{path}", data) is not None

    def _get_large_file(self, path: str) -> Optional[bytes]:
        """큰 파일 다운로드 (Git Data API 사용)."""
        # 먼저 tree에서 blob SHA 찾기
        tree_data = self._get(f"/git/trees/{self.branch}?recursive=1")
        if not tree_data:
            return None

        blob_sha = None
        for item in tree_data.get('tree', []):
            if item['path'] == path:
                blob_sha = item['sha']
                break

        if not blob_sha:
            return None

        # blob 내용 읽기
        blob_data = self._get(f"/git/blobs/{blob_sha}")
        if not blob_data:
            return None

        try:
            return base64.b64decode(blob_data['content'])
        except Exception:
            return None

    def _put_large_file(self, path: str, content: bytes, message: str) -> bool:
        """큰 파일 저장 (Git Data API)."""
        try:
            # 1. blob 생성
            blob_resp = requests.post(
                f"{self.base_url}/git/blobs",
                headers=self.headers,
                json={
                    "content": base64.b64encode(content).decode(),
                    "encoding": "base64"
                },
                timeout=10
            )
            blob_resp.raise_for_status()
            blob_sha = blob_resp.json()['sha']

            # 2. 현재 tree 가져오기
            ref_resp = self._get(f"/git/refs/heads/{self.branch}")
            if not ref_resp:
                return False

            commit_sha = ref_resp['object']['sha']
            commit_resp = self._get(f"/git/commits/{commit_sha}")
            if not commit_resp:
                return False

            tree_sha = commit_resp['tree']['sha']
            tree_resp = self._get(f"/git/trees/{tree_sha}")
            if not tree_resp:
                return False

            # 3. 새 tree 항목 생성
            tree_items = tree_resp.get('tree', [])
            # 기존 파일 제거
            tree_items = [t for t in tree_items if t['path'] != path]
            # 새 파일 추가
            tree_items.append({
                "path": path,
                "mode": "100644",
                "type": "blob",
                "sha": blob_sha
            })

            # 4. 새 tree 생성
            new_tree_resp = requests.post(
                f"{self.base_url}/git/trees",
                headers=self.headers,
                json={"tree": tree_items},
                timeout=10
            )
            new_tree_resp.raise_for_status()
            new_tree_sha = new_tree_resp.json()['sha']

            # 5. 새 commit 생성
            new_commit_resp = requests.post(
                f"{self.base_url}/git/commits",
                headers=self.headers,
                json={
                    "message": message,
                    "tree": new_tree_sha,
                    "parents": [commit_sha]
                },
                timeout=10
            )
            new_commit_resp.raise_for_status()
            new_commit_sha = new_commit_resp.json()['sha']

            # 6. ref 업데이트 (충돌 재시도)
            for attempt in range(3):
                ref_update_resp = requests.patch(
                    f"{self.base_url}/git/refs/heads/{self.branch}",
                    headers=self.headers,
                    json={"sha": new_commit_sha},
                    timeout=10
                )

                if ref_update_resp.status_code == 200:
                    return True
                elif ref_update_resp.status_code == 422:
                    # 충돌: 최신 commit 다시 받아서 재시도
                    if attempt < 2:
                        time.sleep(1)
                        continue

                ref_update_resp.raise_for_status()

            return False
        except Exception:
            return False

    def list_archive(self) -> List[FileInfo]:
        """최종본 목록 (GitHub)."""
        files = []

        # GitHub API로 archive/ 폴더 조회
        tree_data = self._get(f"/git/trees/{self.branch}?recursive=1")
        if not tree_data:
            return files

        for item in tree_data.get('tree', []):
            if item['path'].startswith('archive/') and item['path'].endswith('.pptx'):
                filename = item['path'].split('/')[-1]
                date_str = self._parse_ppt_filename(filename)
                if date_str:
                    # 파일 메타정보 (크기, 수정시간)
                    file_info = self._get_file(item['path'])
                    if file_info:
                        try:
                            import datetime as dt
                            modified_time = dt.datetime.fromisoformat(
                                file_info.get('commit', {}).get('author', {}).get('date', '').replace('Z', '+00:00')
                            )
                        except Exception:
                            modified_time = datetime.now()

                        files.append(FileInfo(
                            name=filename,
                            size=file_info.get('size', 0),
                            modified_time=modified_time,
                            date=date_str
                        ))

        # 날짜순 역순 정렬
        files.sort(key=lambda f: -int(f.date))
        return files

    def get_latest_archive(self) -> Optional[Tuple[str, bytes]]:
        """가장 최근 최종본."""
        files = self.list_archive()
        if not files:
            return None

        latest = files[0]
        path = f"archive/{latest.name}"

        # 큰 파일이므로 Git Data API 사용
        content = self._get_large_file(path)
        if content:
            return latest.name, content

        return None

    def read_archive(self, filename: str) -> Optional[bytes]:
        """최종본 읽기."""
        if not self._validate_ppt_filename(filename):
            return None

        path = f"archive/{filename}"
        return self._get_large_file(path)

    def save_archive(self, filename: str, ppt_bytes: bytes, uploader: Optional[str] = None) -> bool:
        """최종본 저장."""
        if not self._validate_ppt_filename(filename):
            return False

        path = f"archive/{filename}"
        message = f"최종본: {filename}"
        if uploader:
            message += f" (올린 사람: {uploader})"

        return self._put_large_file(path, ppt_bytes, message)

    def delete_archive(self, filename: str) -> bool:
        """최종본 삭제."""
        if not self._validate_ppt_filename(filename):
            return False

        path = f"archive/{filename}"
        file_info = self._get_file(path)
        if not file_info:
            return False

        data = {
            "message": f"최종본 삭제: {filename}",
            "sha": file_info.get('sha'),
            "branch": self.branch
        }

        return self._delete(f"/contents/{path}", data) is not None

    def list_songs(self) -> List[FileInfo]:
        """성가 이미지 목록."""
        files = []

        tree_data = self._get(f"/git/trees/{self.branch}?recursive=1")
        if not tree_data:
            return files

        for item in tree_data.get('tree', []):
            if item['path'].startswith('songs/') and item['path'].split('.')[-1].lower() in ['png', 'jpg', 'jpeg']:
                filename = item['path'].split('/')[-1]
                files.append(FileInfo(
                    name=filename,
                    size=item.get('size', 0),
                    modified_time=datetime.now()
                ))

        return sorted(files, key=lambda f: f.name)

    def read_song(self, hymn_num: int, page: int = 1) -> Optional[bytes]:
        """성가 이미지 읽기."""
        for ext in ['.png', '.jpg', '.jpeg']:
            # 한 장짜리
            path = f"songs/{hymn_num}{ext}"
            content = self._get_large_file(path)
            if content:
                return content

            # 여러 장
            path = f"songs/{hymn_num}-{page}{ext}"
            content = self._get_large_file(path)
            if content:
                return content

        return None

    def get_hymn_image_getter(self) -> callable:
        """generate.py에 전달할 이미지 getter."""
        def getter(hymn_num: int) -> Optional[bytes]:
            return self.read_song(hymn_num, page=1)

        return getter

    def _parse_ppt_filename(self, filename: str) -> Optional[str]:
        """YYMMDD미사.pptx에서 YYMMDD 추출."""
        match = re.match(r'^(\d{6})미사\.pptx$', filename)
        return match.group(1) if match else None

    def _validate_ppt_filename(self, filename: str) -> bool:
        """파일명이 규칙을 따르는지 검증."""
        return self._parse_ppt_filename(filename) is not None
