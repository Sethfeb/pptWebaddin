"""
test_services.py
단위 테스트 (unittest)

테스트 범위:
  - models.SpecRecord
  - settings (임시 파일 사용)
  - services.cache_service.CacheService
  - services.sharepoint_service._parse_items, _odata_escape
  - services.ppt_service (COM 없이 is_running=False 경로)

하드웨어/COM 의존 테스트는 Mock으로 격리.
"""
import os
import sys
import time
import unittest
from unittest.mock import MagicMock, patch, patch as mock_patch

# src-py 경로를 sys.path에 추가
_SRC = os.path.join(os.path.dirname(__file__), "..", "src-py")
sys.path.insert(0, os.path.abspath(_SRC))

from models import SpecRecord
from services.cache_service import CacheService
from services.sharepoint_service import _odata_escape, _parse_items


# ──────────────────────────────────────────────────────────────────────
# SpecRecord 테스트
# ──────────────────────────────────────────────────────────────────────

class TestSpecRecord(unittest.TestCase):
    def test_display_value_with_unit(self) -> None:
        rec = SpecRecord(spec_value="1500", unit="rpm")
        self.assertEqual(rec.display_value, "1500 rpm")

    def test_display_value_no_unit(self) -> None:
        rec = SpecRecord(spec_value="ON", unit="")
        self.assertEqual(rec.display_value, "ON")

    def test_display_value_whitespace_unit(self) -> None:
        rec = SpecRecord(spec_value="100", unit="  ")
        self.assertEqual(rec.display_value, "100")

    def test_str(self) -> None:
        rec = SpecRecord(equip_id="C001", short_code="!c001_spd",
                         spec_name="속도", spec_value="1500", unit="rpm", revision=2)
        s = str(rec)
        self.assertIn("C001", s)
        self.assertIn("!c001_spd", s)
        self.assertIn("rev.2", s)


# ──────────────────────────────────────────────────────────────────────
# settings 테스트
# ──────────────────────────────────────────────────────────────────────

class TestSettings(unittest.TestCase):
    def setUp(self) -> None:
        import tempfile
        self._tmp = tempfile.mkdtemp()
        # settings 모듈의 경로를 임시 디렉터리로 패치
        import settings as cfg
        self._cfg = cfg
        self._orig_dir = cfg._SETTINGS_DIR
        self._orig_file = cfg._SETTINGS_FILE
        cfg._SETTINGS_DIR = self._tmp
        cfg._SETTINGS_FILE = os.path.join(self._tmp, "settings.json")

    def tearDown(self) -> None:
        self._cfg._SETTINGS_DIR = self._orig_dir
        self._cfg._SETTINGS_FILE = self._orig_file

    def test_load_defaults_when_no_file(self) -> None:
        data = self._cfg.load()
        self.assertEqual(data["sharepoint_url"], "")
        self.assertEqual(data["cache_ttl_seconds"], 300)

    def test_save_and_load(self) -> None:
        self._cfg.save({"sharepoint_url": "https://example.com", "cache_ttl_seconds": 60,
                        "prefetch_on_start": False, "shortcode_prefix": "!"})
        data = self._cfg.load()
        self.assertEqual(data["sharepoint_url"], "https://example.com")
        self.assertEqual(data["cache_ttl_seconds"], 60)

    def test_is_configured_false(self) -> None:
        self.assertFalse(self._cfg.is_configured())

    def test_is_configured_true(self) -> None:
        self._cfg.set_value("sharepoint_url", "https://example.com")
        self.assertTrue(self._cfg.is_configured())


# ──────────────────────────────────────────────────────────────────────
# CacheService 테스트
# ──────────────────────────────────────────────────────────────────────

class TestCacheService(unittest.TestCase):
    def _make_rec(self, code: str, value: str = "v") -> SpecRecord:
        return SpecRecord(equip_id="E1", short_code=code,
                          spec_name="N", spec_value=value, unit="", revision=1)

    def test_put_and_get(self) -> None:
        cache = CacheService(ttl_seconds=60)
        rec = self._make_rec("!abc")
        cache.put(rec)
        result = cache.get("!abc")
        self.assertIsNotNone(result)
        assert result is not None
        self.assertEqual(result.short_code, "!abc")

    def test_get_miss(self) -> None:
        cache = CacheService(ttl_seconds=60)
        self.assertIsNone(cache.get("!notexist"))

    def test_ttl_expiry(self) -> None:
        cache = CacheService(ttl_seconds=0.05)
        cache.put(self._make_rec("!exp"))
        time.sleep(0.1)
        self.assertIsNone(cache.get("!exp"))

    def test_put_all_sets_fully_loaded(self) -> None:
        cache = CacheService()
        self.assertFalse(cache.is_fully_loaded)
        cache.put_all([self._make_rec("!a"), self._make_rec("!b")])
        self.assertTrue(cache.is_fully_loaded)
        self.assertEqual(cache.count, 2)

    def test_clear(self) -> None:
        cache = CacheService()
        cache.put_all([self._make_rec("!x")])
        cache.clear()
        self.assertEqual(cache.count, 0)
        self.assertFalse(cache.is_fully_loaded)

    def test_search(self) -> None:
        cache = CacheService()
        cache.put(SpecRecord(equip_id="E1", short_code="!conv_speed",
                             spec_name="컨베이어 속도", spec_value="1500", unit="rpm"))
        cache.put(SpecRecord(equip_id="E2", short_code="!pump_flow",
                             spec_name="펌프 유량", spec_value="200", unit="L/min"))
        results = cache.search("conv")
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].short_code, "!conv_speed")

    def test_build_lookup(self) -> None:
        cache = CacheService()
        cache.put(self._make_rec("!k1", "v1"))
        cache.put(self._make_rec("!k2", "v2"))
        lookup = cache.build_lookup()
        self.assertIn("!k1", lookup)
        self.assertIn("!k2", lookup)


# ──────────────────────────────────────────────────────────────────────
# SharePointService 유틸 함수 테스트
# ──────────────────────────────────────────────────────────────────────

class TestSharePointUtils(unittest.TestCase):
    def test_odata_escape_single_quote(self) -> None:
        self.assertEqual(_odata_escape("O'Brien"), "O''Brien")

    def test_odata_escape_no_quote(self) -> None:
        self.assertEqual(_odata_escape("normal"), "normal")

    def test_parse_items_valid(self) -> None:
        data = {
            "d": {
                "results": [
                    {"EquipID": "C001", "ShortCode": "!c001_spd",
                     "SpecName": "속도", "SpecValue": "1500", "Unit": "rpm", "Revision": 1},
                    {"EquipID": "C002", "ShortCode": "!c002_temp",
                     "SpecName": "온도", "SpecValue": "80", "Unit": "°C", "Revision": 2},
                ]
            }
        }
        items = _parse_items(data)
        self.assertEqual(len(items), 2)
        self.assertEqual(items[0].equip_id, "C001")
        self.assertEqual(items[1].revision, 2)

    def test_parse_items_empty_results(self) -> None:
        data = {"d": {"results": []}}
        self.assertEqual(_parse_items(data), [])

    def test_parse_items_malformed(self) -> None:
        self.assertEqual(_parse_items({}), [])
        self.assertEqual(_parse_items({"d": {}}), [])


# ──────────────────────────────────────────────────────────────────────
# AuthService 테스트 (MSAL 격리)
# ──────────────────────────────────────────────────────────────────────

class TestAuthService(unittest.TestCase):
    def test_normalize_tenant_sharepoint_url(self) -> None:
        from services.auth_service import AuthService
        result = AuthService._normalize_tenant(
            "https://ati5344.sharepoint.com/sites/atimarketing"
        )
        self.assertEqual(result, "ati5344.onmicrosoft.com")

    def test_normalize_tenant_domain(self) -> None:
        from services.auth_service import AuthService
        result = AuthService._normalize_tenant("ati5344.sharepoint.com")
        self.assertEqual(result, "ati5344.onmicrosoft.com")

    def test_normalize_tenant_already_normalized(self) -> None:
        from services.auth_service import AuthService
        result = AuthService._normalize_tenant("ati5344.onmicrosoft.com")
        self.assertEqual(result, "ati5344.onmicrosoft.com")

    def test_normalize_tenant_guid(self) -> None:
        from services.auth_service import AuthService
        guid = "12345678-1234-1234-1234-123456789012"
        self.assertEqual(AuthService._normalize_tenant(guid), guid)

    def test_acquire_token_silent_no_account(self) -> None:
        """캐시에 계정 없으면 None 반환."""
        from services.auth_service import AuthService
        with patch("msal.PublicClientApplication") as MockApp:
            mock_app_inst = MagicMock()
            mock_app_inst.get_accounts.return_value = []
            MockApp.return_value = mock_app_inst
            svc = AuthService.__new__(AuthService)
            svc._app = mock_app_inst
            svc._cache = MagicMock()
            result = svc.acquire_token_silent()
            self.assertIsNone(result)


# ──────────────────────────────────────────────────────────────────────
# SharePointService (token_provider 기반) 테스트
# ──────────────────────────────────────────────────────────────────────

class TestSharePointServiceWithToken(unittest.TestCase):
    def _make_service(self, token: str = "fake_token") -> object:
        from services.sharepoint_service import SharePointService
        return SharePointService(
            site_url="https://ati5344.sharepoint.com/sites/atimarketing",
            token_provider=lambda: token,
        )

    def test_token_provider_called_on_request(self) -> None:
        """요청 시 token_provider가 호출되어 Authorization 헤더에 포함되는지 확인."""
        from services.sharepoint_service import SharePointService
        called = []
        def provider() -> str:
            called.append(True)
            return "test_token"
        svc = SharePointService(
            site_url="https://example.sharepoint.com/sites/test",
            token_provider=provider,
        )
        with patch.object(svc._session, "get") as mock_get:
            mock_resp = MagicMock()
            mock_resp.raise_for_status.return_value = None
            mock_resp.json.return_value = {"d": {"results": []}}
            mock_get.return_value = mock_resp
            svc.search("test")
        self.assertTrue(called)
        call_kwargs = mock_get.call_args[1]
        self.assertIn("Authorization", call_kwargs.get("headers", {}))
        self.assertEqual(call_kwargs["headers"]["Authorization"], "Bearer test_token")

    def test_token_provider_failure_returns_empty(self) -> None:
        """token_provider 예외 시 빈 목록 반환."""
        from services.sharepoint_service import SharePointService
        svc = SharePointService(
            site_url="https://example.sharepoint.com/sites/test",
            token_provider=lambda: (_ for _ in ()).throw(Exception("auth failed")),
        )
        result = svc.search("test")
        self.assertEqual(result, [])
        self.assertIn("인증 토큰 획득 실패", svc.last_error)


# ──────────────────────────────────────────────────────────────────────
# SettingsWindow URL 검증 테스트
# ──────────────────────────────────────────────────────────────────────

class TestUrlValidation(unittest.TestCase):
    def setUp(self) -> None:
        from views.settings_window import _validate_sharepoint_url
        self._validate = _validate_sharepoint_url

    def test_valid_sites_url(self) -> None:
        self.assertEqual(self._validate("https://ati5344.sharepoint.com/sites/atimarketing"), "")

    def test_valid_teams_url(self) -> None:
        self.assertEqual(self._validate("https://ati5344.sharepoint.com/teams/myteam"), "")

    def test_empty_url_allowed(self) -> None:
        self.assertEqual(self._validate(""), "")

    def test_share_link_blocked(self) -> None:
        url = "https://ati5344.sharepoint.com/:l:/s/atimarketing/JAC99AA5NEN0TqtYXt9pFPKhAeDLHmytJ6msR9WFCLJFlqs?e=NCyXyE"
        err = self._validate(url)
        self.assertIn("공유 링크", err)

    def test_http_blocked(self) -> None:
        err = self._validate("http://ati5344.sharepoint.com/sites/atimarketing")
        self.assertIn("https://", err)

    def test_wrong_path_blocked(self) -> None:
        err = self._validate("https://ati5344.sharepoint.com/")
        self.assertNotEqual(err, "")


# ──────────────────────────────────────────────────────────────────────
# PptService 테스트 (COM 격리)
# ──────────────────────────────────────────────────────────────────────

class TestPptService(unittest.TestCase):
    def test_is_running_false_when_no_ppt(self) -> None:
        """COM GetActiveObject 실패 시 False 반환 확인."""
        from services.ppt_service import PptService
        with patch("win32com.client.GetActiveObject", side_effect=Exception("not running")):
            self.assertFalse(PptService.is_running())

    def test_insert_text_returns_false_when_no_ppt(self) -> None:
        from services.ppt_service import PptService
        with patch("win32com.client.GetActiveObject", side_effect=Exception("not running")):
            result = PptService.insert_text("test")
            self.assertFalse(result)

    def test_replace_shortcodes_raises_when_no_ppt(self) -> None:
        from services.ppt_service import PptNotRunningError, PptService
        with patch("win32com.client.GetActiveObject", side_effect=Exception("not running")):
            with self.assertRaises(PptNotRunningError):
                PptService.replace_shortcodes({})

    def test_replace_shortcodes_with_mock_ppt(self) -> None:
        """Mock PPT COM 객체로 치환 로직 검증."""
        from services.ppt_service import PptService

        mock_app = MagicMock()
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_tf = MagicMock()
        mock_chars = MagicMock()

        mock_app.ActivePresentation.Slides = [mock_slide]
        mock_slide.Shapes = [mock_shape]
        mock_shape.HasTextFrame = True
        mock_shape.TextFrame.TextRange = mock_tf
        mock_tf.Text = "속도는 !conv_speed 입니다."
        mock_tf.Characters.return_value = mock_chars

        lookup = {
            "!conv_speed": SpecRecord(spec_value="1500", unit="rpm")
        }

        with patch("win32com.client.GetActiveObject", return_value=mock_app):
            count, not_found = PptService.replace_shortcodes(lookup, scope="all_slides")

        self.assertEqual(count, 1)
        self.assertEqual(not_found, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
