"""用户偏好（字号、奖项方案）持久化到 %APPDATA%\\AwardStudio\\settings.json。"""

from __future__ import annotations

import json
import os
from pathlib import Path

APP_NAME = "AwardStudio"


def settings_path() -> Path:
    base = os.environ.get("APPDATA")
    root = Path(base) if base else Path.home() / ".config"
    return root / APP_NAME / "settings.json"


def load() -> dict:
    path = settings_path()
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save(data: dict) -> None:
    path = settings_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_suffix(".tmp")
        temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.replace(path)
    except OSError:
        # 配置写不进去不影响主流程
        pass
