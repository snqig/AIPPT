from pathlib import Path
from typing import Any, Optional

from aippt.logger import logger


class AssetResult:
    __slots__ = ("slot", "local_path", "width", "height", "source", "asset_type")

    def __init__(self, slot: str, local_path: str, width: int = 0, height: int = 0,
                 source: str = "", asset_type: str = "icon") -> None:
        self.slot = slot
        self.local_path = local_path
        self.width = width
        self.height = height
        self.source = source
        self.asset_type = asset_type


def fetch_assets(
    plan: list[dict],
    cache_dir: str = "assets/cache",
    enable_photos: bool = False,
) -> list[AssetResult]:
    results: list[AssetResult] = []
    icons_dir = str(Path(cache_dir) / "icons")

    from aippt.assets.iconify_fetcher import fetch_icon

    for asset_spec in plan:
        asset_type = asset_spec.get("type", "icon")

        if asset_type == "icon":
            query = asset_spec.get("query", "")
            if ":" in query:
                parts = query.split(":", 1)
                icon_set, name = parts[0], parts[1]
            elif "/" in query:
                icon_set, name = query.split("/", 1)
            else:
                icon_set = asset_spec.get("set", "lucide")
                name = query

            local_path = fetch_icon(
                name=name,
                icon_set=icon_set,
                size=asset_spec.get("size", 128),
                cache_dir=icons_dir,
            )
            if local_path:
                results.append(AssetResult(
                    slot=asset_spec["slot"],
                    local_path=local_path,
                    asset_type="icon",
                    source=f"iconify:{icon_set}/{name}",
                ))

        elif asset_type == "photo":
            if not enable_photos:
                logger.debug("Photos disabled, skipping: %s", asset_spec.get("query"))
                continue
            path = _fetch_photo(asset_spec, cache_dir)
            if path:
                results.append(AssetResult(
                    slot=asset_spec["slot"],
                    local_path=path,
                    asset_type="photo",
                    source=asset_spec.get("query", ""),
                ))

    return results


def _fetch_photo(spec: dict, cache_dir: str) -> Optional[str]:
    query = spec.get("query", "")
    if not query:
        return None
    logger.info("Photo fetch stub: %s (enable with UNSPLASH_ACCESS_KEY)", query)
    return None


def get_image_size(local_path: str) -> tuple[int, int]:
    try:
        from PIL import Image
        with Image.open(local_path) as img:
            return img.size
    except Exception:
        return (0, 0)
