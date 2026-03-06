"""
cache_service.py
TTL 기반 인메모리 캐시

스레드 안전성:
  - threading.Lock 으로 읽기/쓰기 보호.
  - 멀티스레드 환경(백그라운드 프리페치 + UI 스레드)에서 안전하게 사용 가능.

메모리 수명:
  - TTL 만료 항목은 get() 호출 시 지연 삭제(lazy eviction).
  - clear() 로 전체 삭제 가능.
"""
import threading
import time
from typing import Dict, List, Optional, Tuple

from models import SpecRecord

_DEFAULT_TTL = 300  # seconds


class CacheEntry:
    __slots__ = ("record", "expires_at")

    def __init__(self, record: SpecRecord, ttl: float) -> None:
        self.record = record
        self.expires_at = time.monotonic() + ttl


class CacheService:
    """SpecRecord 인메모리 캐시 (short_code → SpecRecord, TTL 지원)."""

    def __init__(self, ttl_seconds: float = _DEFAULT_TTL) -> None:
        self._ttl = ttl_seconds
        self._store: Dict[str, CacheEntry] = {}
        self._lock = threading.Lock()
        self._fully_loaded = False

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get(self, short_code: str) -> Optional[SpecRecord]:
        """단축어로 캐시 조회. 만료 또는 미존재 시 None."""
        with self._lock:
            entry = self._store.get(short_code)
            if entry is None:
                return None
            if time.monotonic() > entry.expires_at:
                del self._store[short_code]
                return None
            return entry.record

    def put(self, record: SpecRecord) -> None:
        """단일 레코드 캐시 저장."""
        with self._lock:
            self._store[record.short_code] = CacheEntry(record, self._ttl)

    def put_all(self, records: List[SpecRecord]) -> None:
        """레코드 목록 일괄 저장. 기존 캐시는 유지하고 덮어씀."""
        now = time.monotonic()
        with self._lock:
            for rec in records:
                self._store[rec.short_code] = CacheEntry(rec, self._ttl)
            self._fully_loaded = True

    def clear(self) -> None:
        """전체 캐시 삭제."""
        with self._lock:
            self._store.clear()
            self._fully_loaded = False

    def build_lookup(self) -> Dict[str, SpecRecord]:
        """
        현재 캐시 전체를 {short_code: SpecRecord} dict로 반환.
        단축어 일괄 치환 시 사용.
        만료 항목은 제외.
        """
        now = time.monotonic()
        with self._lock:
            return {
                k: v.record
                for k, v in self._store.items()
                if now <= v.expires_at
            }

    def search(self, keyword: str) -> List[SpecRecord]:
        """
        캐시 내에서 keyword로 ShortCode 또는 SpecName 부분 검색.
        SharePoint 미연결 상태에서 오프라인 검색에 사용.
        """
        kw = keyword.lower()
        now = time.monotonic()
        results: List[SpecRecord] = []
        with self._lock:
            for entry in self._store.values():
                if now > entry.expires_at:
                    continue
                rec = entry.record
                if kw in rec.short_code.lower() or kw in rec.spec_name.lower():
                    results.append(rec)
        return results

    @property
    def count(self) -> int:
        """현재 유효한 캐시 항목 수."""
        now = time.monotonic()
        with self._lock:
            return sum(1 for v in self._store.values() if now <= v.expires_at)

    @property
    def is_fully_loaded(self) -> bool:
        """put_all() 로 전체 목록이 로드된 상태인지 여부."""
        return self._fully_loaded

    def set_ttl(self, ttl_seconds: float) -> None:
        """TTL 변경 (기존 항목에는 소급 미적용)."""
        self._ttl = ttl_seconds
