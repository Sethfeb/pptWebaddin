"""
sharepoint_service.py
SharePoint REST API v1 연동 서비스

인증 전략:
  - requests_ntlm 없이 requests 기본 auth=None 으로 시도 (SharePoint Online은 쿠키/토큰 필요)
  - Windows 통합 인증(NTLM/Kerberos)은 requests-negotiate-sspi 패키지 필요
  - 현재는 Windows 자격증명 자동 전달을 위해 requests_ntlm 또는 HttpNegotiateAuth 사용
  - 패키지 미설치 시 NoAuth 모드로 폴백하고 경고 출력

스레드 안전성:
  - Session 객체는 인스턴스 단위로 소유. 멀티스레드 공유 금지.
  - 호출 측에서 스레드별 인스턴스 생성 권장.

메모리/수명:
  - requests.Session은 with 문 또는 명시적 close() 로 해제.
  - 본 클래스는 close() 메서드 제공.
"""
import re
from typing import List, Optional, Tuple

import requests

from models import SpecRecord

_LIST_NAME = "EquipmentSpecs"
_SELECT_FIELDS = "EquipID,ShortCode,SpecName,SpecValue,Unit,Revision"
_TIMEOUT_SEC = 15


def _make_session() -> Tuple[requests.Session, bool]:
    """
    Windows 통합 인증 세션 생성 시도.
    성공 시 (session, True), 실패 시 (기본 세션, False) 반환.
    """
    session = requests.Session()
    session.headers.update({
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
    })

    # requests-negotiate-sspi (NTLM/Kerberos 자동 선택) 우선 시도
    try:
        from requests_negotiate_sspi import HttpNegotiateAuth  # type: ignore
        session.auth = HttpNegotiateAuth()
        return session, True
    except ImportError:
        pass

    # requests-ntlm 폴백
    try:
        import sspi  # noqa: F401  (pywin32 포함)
        import requests_ntlm  # type: ignore
        session.auth = requests_ntlm.HttpNtlmAuth("", "", send_single_token=True)
        return session, True
    except (ImportError, Exception):
        pass

    # 인증 없음 (SharePoint On-Prem 익명 허용 환경 또는 테스트용)
    return session, False


def _odata_escape(value: str) -> str:
    """OData 문자열 값에서 작은따옴표 이스케이프."""
    return value.replace("'", "''")


def _parse_items(json_data: dict) -> List[SpecRecord]:
    """
    SharePoint REST API odata=verbose 응답에서 SpecRecord 목록 추출.
    응답 구조: {"d": {"results": [{"EquipID": ..., ...}, ...]}}
    """
    try:
        raw_items = json_data["d"]["results"]
    except (KeyError, TypeError):
        return []

    records: List[SpecRecord] = []
    for item in raw_items:
        records.append(SpecRecord(
            equip_id=item.get("EquipID") or "",
            short_code=item.get("ShortCode") or "",
            spec_name=item.get("SpecName") or "",
            spec_value=item.get("SpecValue") or "",
            unit=item.get("Unit") or "",
            revision=int(item.get("Revision") or 0),
        ))
    return records


class SharePointService:
    """SharePoint EquipmentSpecs 리스트 조회 서비스."""

    def __init__(self, site_url: str) -> None:
        """
        Parameters
        ----------
        site_url : str
            SharePoint 사이트 루트 URL
            예) https://contoso.sharepoint.com/sites/factory
        """
        self._site_url = site_url.rstrip("/")
        self._session, self._auth_ok = _make_session()
        if not self._auth_ok:
            print("[EquipSpec][WARN] Windows 통합 인증 패키지 없음 - 익명 요청으로 진행")

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_by_short_code(self, short_code: str) -> Optional[SpecRecord]:
        """단축어 1건으로 사양 레코드 조회. 없으면 None."""
        escaped = _odata_escape(short_code)
        endpoint = (
            f"{self._site_url}/_api/web/lists/getbytitle('{_LIST_NAME}')"
            f"/items?$select={_SELECT_FIELDS}"
            f"&$filter=ShortCode eq '{escaped}'"
            f"&$top=1"
        )
        items = self._get_items(endpoint)
        return items[0] if items else None

    def get_all(self) -> List[SpecRecord]:
        """전체 사양 목록 반환 (캐시 프리페치용, 최대 5000건)."""
        endpoint = (
            f"{self._site_url}/_api/web/lists/getbytitle('{_LIST_NAME}')"
            f"/items?$select={_SELECT_FIELDS}"
            f"&$orderby=EquipID,ShortCode"
            f"&$top=5000"
        )
        return self._get_items(endpoint)

    def search(self, keyword: str) -> List[SpecRecord]:
        """
        키워드로 ShortCode 또는 SpecName 부분 검색 (최대 100건).
        OData v3 substringof 함수 사용 (SharePoint REST API v1 지원).
        """
        escaped = _odata_escape(keyword)
        endpoint = (
            f"{self._site_url}/_api/web/lists/getbytitle('{_LIST_NAME}')"
            f"/items?$select={_SELECT_FIELDS}"
            f"&$filter=substringof('{escaped}',ShortCode) or "
            f"substringof('{escaped}',SpecName)"
            f"&$top=100"
        )
        return self._get_items(endpoint)

    def close(self) -> None:
        """HTTP 세션 해제."""
        self._session.close()

    def __enter__(self) -> "SharePointService":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _get_items(self, endpoint: str) -> List[SpecRecord]:
        """GET 요청 실행 후 SpecRecord 목록 반환. 오류 시 빈 목록."""
        try:
            resp = self._session.get(endpoint, timeout=_TIMEOUT_SEC)
            resp.raise_for_status()
            return _parse_items(resp.json())
        except requests.exceptions.RequestException as exc:
            print(f"[EquipSpec][ERROR] SharePoint request failed: {exc}")
            return []
        except (ValueError, KeyError) as exc:
            print(f"[EquipSpec][ERROR] JSON parse failed: {exc}")
            return []
