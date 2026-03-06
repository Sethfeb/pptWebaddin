"""
auth_service.py
Microsoft OAuth 2.0 인증 서비스 (MSAL Device Code Flow)

인증 흐름:
  1. acquire_token() 호출
  2. 캐시에 유효한 토큰 있으면 즉시 반환 (silent)
  3. 없으면 Device Code Flow 시작:
     - 콜백(on_device_flow)으로 UI에 로그인 URL + 코드 전달
     - 사용자가 브라우저에서 로그인 완료
     - 토큰 수신 후 캐시 저장
  4. 이후 호출 시 캐시에서 자동 복원 / 만료 시 refresh_token으로 자동 갱신

토큰 캐시:
  - 위치: %APPDATA%/EquipSpecTool/token_cache.bin
  - msal.SerializableTokenCache 사용 (암호화 없음, 로컬 전용)
  - [추정] 민감 환경에서는 DPAPI 암호화 추가 권장

스레드 안전성:
  - acquire_token()은 블로킹 호출 (Device Code 대기 최대 15분).
  - 반드시 백그라운드 스레드에서 호출하고 결과를 after()로 UI에 전달.

권한 범위:
  - Sites.Read.All: SharePoint 리스트 읽기
  - offline_access: refresh_token 발급 (자동 갱신)
"""
import json
import os
from typing import Callable, Dict, Optional

import msal  # type: ignore

# Azure CLI 공개 클라이언트 ID
# 출처: https://learn.microsoft.com/en-us/cli/azure/authenticate-azure-cli
# Azure CLI는 SharePoint 리소스 스코프에 대해 사전 승인되어 있어 AADSTS65002 미발생
_CLIENT_ID = "04b07795-8542-4c45-a7e5-921ad3d5b33d"  # Microsoft Azure CLI

# SharePoint 리소스 직접 스코프 — 테넌트별로 동적으로 구성
# _SCOPES는 AuthService.__init__ 에서 site_url 기반으로 설정됨
# 형식: https://{tenant}.sharepoint.com/AllSites.Read
# offline_access, profile, openid 는 MSAL이 자동 추가하는 예약 스코프 — 직접 명시 금지
_SCOPES_TEMPLATE = "{sharepoint_host}/AllSites.Read"

_APPDATA = os.environ.get("APPDATA", os.path.expanduser("~"))
_CACHE_DIR = os.path.join(_APPDATA, "EquipSpecTool")
_CACHE_FILE = os.path.join(_CACHE_DIR, "token_cache.bin")

# Device Code Flow 타임아웃 (초) — MSAL 기본값 15분
_DEVICE_FLOW_TIMEOUT = 300


class AuthError(Exception):
    """인증 실패 시 발생."""


class AuthService:
    """
    MSAL 기반 Microsoft OAuth 인증 서비스.

    Parameters
    ----------
    site_url : str
        SharePoint 사이트 루트 URL
        예) https://ati5344.sharepoint.com/sites/atimarketing
    """

    def __init__(self, site_url: str) -> None:
        """
        Parameters
        ----------
        site_url : str
            SharePoint 사이트 루트 URL
            예) https://ati5344.sharepoint.com/sites/atimarketing
        """
        self._tenant_id = self._normalize_tenant(site_url)
        self._sharepoint_host = self._extract_host(site_url)
        # SharePoint 리소스 직접 스코프 (테넌트별 동적 구성)
        self._scopes = [_SCOPES_TEMPLATE.format(sharepoint_host=self._sharepoint_host)]
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        self._app = msal.PublicClientApplication(
            client_id=_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            token_cache=self._cache,
        )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def acquire_token_silent(self) -> Optional[str]:
        """
        캐시에서 토큰 조회 (네트워크 요청 없음).
        유효한 토큰 있으면 access_token 반환, 없으면 None.
        """
        accounts = self._app.get_accounts()
        if not accounts:
            return None
        result = self._app.acquire_token_silent(self._scopes, account=accounts[0])
        if result and "access_token" in result:
            self._save_cache()
            return result["access_token"]
        return None

    def acquire_token_device_flow(
        self,
        on_device_flow: Callable[[str, str], None],
    ) -> str:
        """
        Device Code Flow로 토큰 획득.

        Parameters
        ----------
        on_device_flow : (url, user_code) -> None
            로그인 URL과 코드를 UI에 전달하는 콜백.
            예) lambda url, code: show_login_dialog(url, code)

        반환값: access_token 문자열
        예외: AuthError (실패 또는 타임아웃)
        """
        flow = self._app.initiate_device_flow(scopes=self._scopes)
        if "user_code" not in flow:
            raise AuthError(f"Device Flow 시작 실패: {flow.get('error_description', '알 수 없는 오류')}")

        # UI 콜백으로 로그인 안내 전달
        on_device_flow(
            flow["verification_uri"],
            flow["user_code"],
        )

        # 사용자 로그인 대기 (블로킹 — exit_condition 없이 호출하면 완료까지 대기)
        result = self._app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            err = result.get("error_description") or result.get("error") or "알 수 없는 오류"
            raise AuthError(f"로그인 실패: {err}")

        self._save_cache()
        return result["access_token"]

    def get_token(self, on_device_flow: Callable[[str, str], None]) -> str:
        """
        캐시 우선 → 없으면 Device Code Flow.
        항상 유효한 access_token 반환.
        """
        token = self.acquire_token_silent()
        if token:
            return token
        return self.acquire_token_device_flow(on_device_flow)

    def logout(self) -> None:
        """캐시 삭제 (로그아웃)."""
        accounts = self._app.get_accounts()
        for account in accounts:
            self._app.remove_account(account)
        self._cache.deserialize("{}")
        if os.path.isfile(_CACHE_FILE):
            os.remove(_CACHE_FILE)

    @property
    def is_logged_in(self) -> bool:
        """캐시에 계정이 있으면 True (토큰 유효성 미검증)."""
        return bool(self._app.get_accounts())

    @property
    def account_name(self) -> str:
        """로그인된 계정명. 없으면 빈 문자열."""
        accounts = self._app.get_accounts()
        if accounts:
            return accounts[0].get("username", "")
        return ""

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    @staticmethod
    def _extract_host(site_url: str) -> str:
        """
        SharePoint 사이트 URL에서 호스트 부분만 추출.
        예) https://ati5344.sharepoint.com/sites/atimarketing
              → https://ati5344.sharepoint.com
        """
        url = site_url.strip()
        if "://" in url:
            scheme, rest = url.split("://", 1)
            host = rest.split("/")[0]
            return f"{scheme}://{host}"
        return url

    @staticmethod
    def _normalize_tenant(tenant_id: str) -> str:
        """
        SharePoint URL 또는 테넌트 도메인에서 테넌트 식별자 추출.
        예)
          "https://ati5344.sharepoint.com/sites/atimarketing"
            → "ati5344.onmicrosoft.com"
          "ati5344.sharepoint.com"
            → "ati5344.onmicrosoft.com"
          "ati5344.onmicrosoft.com" (이미 정규화됨)
            → "ati5344.onmicrosoft.com"
          GUID 형식 → 그대로 반환
        """
        t = tenant_id.strip().lower()
        # GUID 형식이면 그대로
        if len(t) == 36 and t.count("-") == 4:
            return t
        # URL에서 호스트 추출
        if "://" in t:
            t = t.split("://", 1)[1].split("/")[0]
        # sharepoint.com 도메인이면 테넌트명 추출
        if t.endswith(".sharepoint.com"):
            tenant_name = t.replace(".sharepoint.com", "")
            return f"{tenant_name}.onmicrosoft.com"
        # 이미 onmicrosoft.com 형식이거나 커스텀 도메인
        return t

    def _load_cache(self) -> None:
        if os.path.isfile(_CACHE_FILE):
            try:
                with open(_CACHE_FILE, "r", encoding="utf-8") as f:
                    self._cache.deserialize(f.read())
            except (OSError, ValueError):
                pass

    def _save_cache(self) -> None:
        if self._cache.has_state_changed:
            os.makedirs(_CACHE_DIR, exist_ok=True)
            with open(_CACHE_FILE, "w", encoding="utf-8") as f:
                f.write(self._cache.serialize())
