"""
settings.py
JSON 파일 기반 설정 관리
저장 위치: %APPDATA%/EquipSpecTool/settings.json
"""
import json
import os
from typing import Any, Dict


_APPDATA = os.environ.get("APPDATA", os.path.expanduser("~"))
_SETTINGS_DIR = os.path.join(_APPDATA, "EquipSpecTool")
_SETTINGS_FILE = os.path.join(_SETTINGS_DIR, "settings.json")

_DEFAULTS: Dict[str, Any] = {
    "sharepoint_url": "",
    "cache_ttl_seconds": 300,
    "prefetch_on_start": True,
    "shortcode_prefix": "!",
}


def _ensure_dir() -> None:
    os.makedirs(_SETTINGS_DIR, exist_ok=True)


def load() -> Dict[str, Any]:
    """설정 파일을 읽어 dict로 반환. 없으면 기본값 반환."""
    _ensure_dir()
    if not os.path.isfile(_SETTINGS_FILE):
        return dict(_DEFAULTS)
    try:
        with open(_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data: Dict[str, Any] = json.load(f)
        merged = dict(_DEFAULTS)
        merged.update(data)
        return merged
    except (json.JSONDecodeError, OSError):
        return dict(_DEFAULTS)


def save(data: Dict[str, Any]) -> None:
    """설정 dict를 파일에 저장."""
    _ensure_dir()
    with open(_SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get(key: str) -> Any:
    """단일 키 값 조회."""
    return load().get(key, _DEFAULTS.get(key))


def set_value(key: str, value: Any) -> None:
    """단일 키 값 저장."""
    data = load()
    data[key] = value
    save(data)


def is_configured() -> bool:
    """SharePoint URL이 설정되어 있는지 확인."""
    url: str = get("sharepoint_url")
    return bool(url and url.strip())
