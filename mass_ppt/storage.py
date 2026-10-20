"""
PPT와 성가 이미지를 저장하고 불러오는 저장소 인터페이스.
나중에 GitHub 구현으로 교체 가능.
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, List, Tuple, Dict
from dataclasses import dataclass
import re
from datetime import datetime


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
