"""System-font glyph outlines for the MathJax typography adapter.

This is the macOS counterpart of the Windows plugin's GDI+
``SystemFontOutlineService``. It returns the same ``FontRunOutline`` fields:
baseline-relative SVG path data in 1000 units per em with the y axis pointing
up, plus advance, ink height/depth and horizontal ink bounds in em units.
Fallback is explicit and ordered; Qt's implicit per-glyph substitution is
never used because each candidate must cover the whole run.
"""

from __future__ import annotations

import threading
from collections.abc import Iterable
from typing import Any

EM = 1000
REGULAR, BOLD, ITALIC = 0, 1, 2
FALLBACK_FAMILIES = (
    "Times New Roman",
    "PingFang SC",
    "Songti SC",
    "Apple Symbols",
    "Menlo",
)
MAX_RUN_CHARACTERS = 2048


class OutlineError(ValueError):
    pass


class SystemFontOutlines:
    def __init__(self, fallback_families: Iterable[str] = FALLBACK_FAMILIES) -> None:
        self.fallback_families = tuple(fallback_families)
        self._lock = threading.Lock()
        self._families: dict[str, str] | None = None
        self._cache: dict[tuple[str, str, int], dict[str, Any]] = {}

    def families(self) -> list[str]:
        # Names starting with "." are private macOS UI fonts, not user choices.
        return sorted(
            {name for name in self._installed().values() if not name.startswith(".")},
            key=str.casefold,
        )

    def shape_runs(self, runs: list[dict[str, Any]]) -> dict[str, Any]:
        """Shape adapter runs, keyed by the adapter's opaque run key."""

        result: dict[str, Any] = {}
        for run in runs:
            key, text, family, style = (
                run.get("key"),
                run.get("text"),
                run.get("family"),
                run.get("style"),
            )
            if (
                not isinstance(key, str)
                or not isinstance(text, str)
                or not isinstance(family, str)
            ):
                raise OutlineError("文字段数据无效。")
            if (
                isinstance(style, bool)
                or not isinstance(style, int)
                or style & ~(BOLD | ITALIC)
            ):
                raise OutlineError("文字段样式无效。")
            result[key] = self.shape(text, family, style)
        return result

    def shape(self, text: str, family: str, style: int) -> dict[str, Any]:
        if len(text) > MAX_RUN_CHARACTERS:
            raise OutlineError("文字段过长。")
        for character in text:
            if (
                ord(character) < 32
                or 0xD800 <= ord(character) <= 0xDFFF
                or ord(character) > 0xFFFF
            ):
                raise OutlineError(
                    "系统字体轮廓暂不支持辅助平面字符或控制字符；请使用受支持的数学字形或显式换行结构。"
                )
        if not family.strip():
            raise OutlineError("字体家族不能为空。")
        cache_key = (text, family.strip(), style)
        with self._lock:
            cached = self._cache.get(cache_key)
            if cached is None:
                cached = self._shape(text, family.strip(), style)
                if len(self._cache) > 4096:
                    self._cache.clear()
                self._cache[cache_key] = cached
            return cached

    def _installed(self) -> dict[str, str]:
        if self._families is None:
            from PyQt6.QtGui import QFontDatabase

            self._families = {
                name.casefold(): name for name in QFontDatabase.families()
            }
        return self._families

    def _shape(self, text: str, requested: str, style: int) -> dict[str, Any]:
        from PyQt6.QtGui import QFont, QRawFont

        installed = self._installed()
        failures: list[str] = []
        seen: set[str] = set()
        for candidate in (requested, *self.fallback_families):
            name = installed.get(candidate.casefold())
            if candidate.casefold() in seen:
                continue
            seen.add(candidate.casefold())
            if name is None:
                failures.append(f"{candidate}：未安装")
                continue
            font = QFont(name)
            font.setPixelSize(EM)
            font.setBold(bool(style & BOLD))
            font.setItalic(bool(style & ITALIC))
            font.setKerning(True)
            font.setStyleStrategy(QFont.StyleStrategy.NoFontMerging)
            raw = QRawFont.fromFont(font)
            missing = [
                f"U+{ord(c):04X}"
                for c in text
                if not c.isspace() and not raw.supportsCharacter(c)
            ]
            if missing:
                failures.append(f"{name}：缺少 {' '.join(missing)}")
                continue
            warning = f"{'；'.join(failures)}；已使用 {name}" if failures else None
            return _outline(font, text, requested, name, style, warning)
        raise OutlineError("无法为文字段选择字体：" + "；".join(failures))


def _outline(
    font, text: str, requested: str, actual: str, style: int, warning: str | None
) -> dict[str, Any]:
    from PyQt6.QtCore import QPointF
    from PyQt6.QtGui import QFontMetricsF, QPainterPath, QTransform

    advance = QFontMetricsF(font).horizontalAdvance(text) / EM if text else 0.0
    path = QPainterPath()
    if text:
        path.addText(QPointF(0, 0), font, text)
    path = QTransform.fromScale(1, -1).map(path)
    base = {
        "Text": text,
        "RequestedFamily": requested,
        "ActualFamily": actual,
        "Style": style,
        "Advance": advance,
        "Warning": warning,
    }
    if path.elementCount() == 0:
        return {
            **base,
            "Height": 0.0,
            "Depth": 0.0,
            "Left": 0.0,
            "Right": 0.0,
            "Path": "",
        }
    bounds = path.boundingRect()
    return {
        **base,
        "Height": max(0.0, bounds.bottom()) / EM,
        "Depth": max(0.0, -bounds.top()) / EM,
        "Left": bounds.left() / EM,
        "Right": bounds.right() / EM,
        "Path": _svg_path(path),
    }


def _svg_path(path) -> str:
    from PyQt6.QtGui import QPainterPath

    kinds = QPainterPath.ElementType
    parts: list[str] = []
    count = path.elementCount()
    index = 0
    subpath_start = None
    while index < count:
        element = path.elementAt(index)
        if element.type == kinds.MoveToElement:
            if subpath_start is not None:
                parts.append("Z")
            parts.append(f"M{_n(element.x)} {_n(element.y)}")
            subpath_start = (element.x, element.y)
            index += 1
        elif element.type == kinds.LineToElement:
            parts.append(f"L{_n(element.x)} {_n(element.y)}")
            index += 1
        elif element.type == kinds.CurveToElement:
            if index + 2 >= count:
                raise OutlineError("字形曲线路径不完整。")
            second, end = path.elementAt(index + 1), path.elementAt(index + 2)
            parts.append(
                f"C{_n(element.x)} {_n(element.y)} {_n(second.x)} {_n(second.y)} {_n(end.x)} {_n(end.y)}"
            )
            index += 3
        else:
            raise OutlineError("字形路径类型无效。")
    if subpath_start is not None:
        parts.append("Z")
    return "".join(parts)


def _n(value: float) -> str:
    text = f"{value:.3f}".rstrip("0").rstrip(".")
    return "0" if text in {"", "-0"} else text
