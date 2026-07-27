import hashlib
from pathlib import Path
from typing import Optional

from aippt.logger import logger

ICONIFY_API = "https://api.iconify.design"

_DEFAULT_SET = "lucide"
_DEFAULT_SIZE = 128
_DEFAULT_COLOR = "currentColor"
_DEFAULT_STROKE = 1.5

_HAS_REQUESTS = True
try:
    import requests  # noqa: F401
except ImportError:
    _HAS_REQUESTS = False


def fetch_icon(
    name: str,
    icon_set: str = _DEFAULT_SET,
    size: int = _DEFAULT_SIZE,
    color: Optional[str] = None,
    stroke_width: float = _DEFAULT_STROKE,
    cache_dir: str = "assets/cache/icons",
) -> Optional[str]:
    if ":" in name:
        icon_set, name = name.split(":", 1)
    if not name:
        return None

    cache_path = Path(cache_dir)
    cache_path.mkdir(parents=True, exist_ok=True)

    key = f"{icon_set}_{name}_{size}_{color or 'none'}_{stroke_width}"
    hash_key = hashlib.md5(key.encode()).hexdigest()
    local_svg = cache_path / f"{hash_key}.svg"
    local_png = cache_path / f"{hash_key}.png"

    if local_png.exists():
        return str(local_png)
    if local_svg.exists():
        return str(local_svg)

    if not _HAS_REQUESTS:
        logger.warning("requests 未安装，无法获取图标: %s:%s", icon_set, name)
        return None

    try:
        import requests
        url = f"{ICONIFY_API}/{icon_set}/{name}.svg"
        params = {"height": size}
        if color:
            params["color"] = color
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        local_svg.write_bytes(resp.content)
        logger.info("Iconify cached: %s → %s", key, local_svg)
        return str(local_svg)
    except Exception as e:
        logger.warning("Iconify fetch failed [%s:%s]: %s", icon_set, name, e)
        return None


def fetch_icon_batch(
    queries: list[dict],
    cache_dir: str = "assets/cache/icons",
) -> dict[str, str]:
    results = {}
    for q in queries:
        path = fetch_icon(
            name=q.get("query", ""),
            icon_set=q.get("set", _DEFAULT_SET),
            size=q.get("size", _DEFAULT_SIZE),
            color=q.get("color"),
            cache_dir=cache_dir,
        )
        if path:
            results[q.get("slot", q.get("query", ""))] = path
    return results
