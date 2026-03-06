"""
sharepoint_service.py
SharePoint REST API v1 연동 서비스

인증 전략:
  - MSAL OAuth 2.0 Bearer 토큰 사용 (SharePoint Online 표준)
  - AuthService.get_token() 으로 획득한 access_token을 Authorization 헤더에 전달
  - 토큰 만료 시 자동 갱신 (MSAL 캐시)

스레드 안전성:
  - Session 객체는 인스턴스 단위로 소유. 멀티스레드 공유 금지.
  - 호출 측에서 스레드별 인스턴스 생성 권장.

메모리/수명:
  - requests.Session은 close() 로 해제.
  - 본 클래스는 close() 메서드 제공.
"""
from typing import Callable, List, Optional

import requests

from models import SpecRecord

_LIST_NAME = "EquipmentSpecs"
_SELECT_FIELDS = "EquipID,ShortCode,SpecName,SpecValue,Unit,Revision"
_TIMEOUT_SEC = 15


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
    """SharePoint EquipmentSpecs 리스트 조회 서비스 (MSAL OAuth 인증)."""

    def __init__(
        self,
        site_url: str,
        token_provider: Callable[[], str],
    ) -> None:
        """
        Parameters
        ----------
        site_url : str
            SharePoint 사이트 루트 URL
            예) https://ati5344.sharepoint.com/sites/atimarketing
        token_provider : () -> str
            호출 시마다 유효한 access_token 문자열을 반환하는 콜백.
            AuthService.get_token() 을 래핑하여 전달.
        """
        self._site_url = site_url.rstrip("/")
        self._token_provider = token_provider
        self._last_error: str = ""
        self._session = requests.Session()
        self._session.headers.update({
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        })

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

    @property
    def last_error(self) -> str:
        """마지막 요청 오류 메시지. 정상이면 빈 문자열."""
        return self._last_error

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

    def _get_auth_header(self) -> Optional[str]:
        """token_provider 호출하여 Bearer 토큰 헤더값 반환. 실패 시 None."""
        try:
            token = self._token_provider()
            return f"Bearer {token}"
        except Exception as exc:
            self._last_error = f"인증 토큰 획득 실패: {exc}"
            print(f"[EquipSpec][ERROR] {self._last_error}")
            return None

    def _get_items(self, endpoint: str) -> List[SpecRecord]:
        """GET 요청 실행 후 SpecRecord 목록 반환. 오류 시 빈 목록."""
        self._last_error = ""

        auth_header = self._get_auth_header()
        if auth_header is None:
            return []

        try:
            resp = self._session.get(
                endpoint,
                headers={"Authorization": auth_header},
                timeout=_TIMEOUT_SEC,
            )
            resp.raise_for_status()
            return _parse_items(resp.json())
        except requests.exceptions.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else "?"
            if status == 401:
                self._last_error = "인증 만료 (HTTP 401) — 로그아웃 후 다시 로그인하세요."
            elif status == 403:
                self._last_error = (
                    "접근 거부 (HTTP 403) — 리스트 읽기 권한이 없습니다.\n"
                    "SharePoint 사이트 관리자에게 'Sites.Read.All' 권한을 요청하세요."
                )
            elif status == 404:
                self._last_error = (
                    f"리스트를 찾을 수 없습니다 (HTTP 404).\n"
                    f"사이트 URL이 올바른지, '{_LIST_NAME}' 리스트가 존재하는지 확인하세요."
                )
            else:
                self._last_error = f"HTTP 오류 {status}: {exc}"
            print(f"[EquipSpec][ERROR] {self._last_error}")
            return []
        except requests.exceptions.ConnectionError as exc:
            self._last_error = f"연결 실패 — URL이 올바른지, 네트워크를 확인하세요.\n({exc})"
            print(f"[EquipSpec][ERROR] {self._last_error}")
            return []
        except requests.exceptions.Timeout:
            self._last_error = f"요청 시간 초과 ({_TIMEOUT_SEC}초) — 네트워크 상태를 확인하세요."
            print(f"[EquipSpec][ERROR] {self._last_error}")
            return []
        except requests.exceptions.RequestException as exc:
            self._last_error = f"요청 오류: {exc}"
            print(f"[EquipSpec][ERROR] {self._last_error}")
            return []
        except (ValueError, KeyError) as exc:
            self._last_error = f"응답 파싱 오류: {exc}"
            print(f"[EquipSpec][ERROR] {self._last_error}")
            return []
