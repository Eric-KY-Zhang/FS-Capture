from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field

if sys.version_info >= (3, 11):
    import tomllib
else:
    import tomli as tomllib  # type: ignore

import tomli_w


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
        return s
    with cfg_file.open("rb") as f:
        data: dict[str, Any] = tomllib.load(f)
    return Settings.model_validate(data)


def save_settings(settings: Settings) -> None:
    cfg_file = config_path()
    cfg_file.parent.mkdir(parents=True, exist_ok=True)
    with cfg_file.open("wb") as f:
        tomli_w.dump(settings.model_dump(), f)
