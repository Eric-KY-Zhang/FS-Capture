from __future__ import annotations

import importlib
import os
import sys
import tomllib
from pathlib import Path
from typing import Any

import tomli_w
from pydantic import BaseModel, Field, model_validator


def project_root() -> Path:
    """Resolve project root whether running from source or PyInstaller bundle."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


class PathsCfg(BaseModel):
    output_dir: str = "output"
    cache_dir: str = "data/cache"
    log_dir: str = "data/logs"


class ConcurrencyCfg(BaseModel):
    max_workers: int = 4


class RateLimitsCfg(BaseModel):
    cninfo: float = 5.0
    hkexnews: float = 3.0
    sec: float = 8.0
    dart: float = 5.0
    akshare: float = 4.0
    twse: float = 2.0


class SECCfg(BaseModel):
    user_agent: str = "FS Capture (kaiyu199602@gmail.com)"


class DARTCfg(BaseModel):
    api_key: str = ""


class UICfg(BaseModel):
    theme: str = "light"
    language: str = "zh_CN"
    window_width: int = 1280
    window_height: int = 820


class Settings(BaseModel):
    paths: PathsCfg = Field(default_factory=PathsCfg)
    concurrency: ConcurrencyCfg = Field(default_factory=ConcurrencyCfg)
    rate_limits: RateLimitsCfg = Field(default_factory=RateLimitsCfg)
    sec: SECCfg = Field(default_factory=SECCfg)
    dart: DARTCfg = Field(default_factory=DARTCfg)
    ui: UICfg = Field(default_factory=UICfg)

    @model_validator(mode="before")
    @classmethod
    def _normalize_dart_key(cls, data: Any) -> Any:
        if not isinstance(data, dict):
            return data

        normalized = dict(data)
        dart = dict(normalized.get("dart") or {})
        opendart = normalized.get("opendart") or {}

        dart_key = str(dart.get("api_key") or "").strip()
        alias_key = ""
        if isinstance(opendart, dict):
            alias_key = str(opendart.get("api_key") or "").strip()
        env_key = os.environ.get("OPENDART_API_KEY", "").strip()

        if dart_key:
            dart["api_key"] = dart_key
            normalized["dart"] = dart
        else:
            fallback_key = alias_key or env_key
            if fallback_key:
                dart["api_key"] = fallback_key
                normalized["dart"] = dart

        return normalized

    def output_path(self) -> Path:
        return self._resolve(self.paths.output_dir)

    def cache_path(self) -> Path:
        return self._resolve(self.paths.cache_dir)

    def log_path(self) -> Path:
        return self._resolve(self.paths.log_dir)

    @staticmethod
    def _resolve(p: str) -> Path:
        path = Path(p)
        if not path.is_absolute():
            path = project_root() / path
        path.mkdir(parents=True, exist_ok=True)
        return path


def config_path() -> Path:
    return project_root() / "config.toml"


def load_settings() -> Settings:
    cfg_file = config_path()
    if not cfg_file.exists():
        s = Settings()
        save_settings(s)
        return Settings.model_validate({})
    with cfg_file.open("rb") as f:
        data: dict[str, Any] = tomllib.load(f)
    return Settings.model_validate(data)


def save_settings(settings: Settings) -> None:
    cfg_file = config_path()
    cfg_file.parent.mkdir(parents=True, exist_ok=True)
    with cfg_file.open("wb") as f:
        tomli_w.dump(settings.model_dump(), f)


def invalidate_dart_client_cache() -> bool:
    """Clear the cached KR OpenDartReader client if that plugin is available."""
    try:
        module = importlib.import_module("plugins.kr.name_resolver")
    except Exception:
        return False

    reset_client = getattr(module, "reset_dart_client", None)
    if callable(reset_client):
        try:
            reset_client()
        except Exception:
            return False
        return True

    cached_factory = getattr(module, "_dart", None)
    cache_clear = getattr(cached_factory, "cache_clear", None)
    if not callable(cache_clear):
        cached_factory = getattr(module, "_dart_for_key", None)
        cache_clear = getattr(cached_factory, "cache_clear", None)
        if not callable(cache_clear):
            return False

    try:
        cache_clear()
    except Exception:
        return False
    return True
