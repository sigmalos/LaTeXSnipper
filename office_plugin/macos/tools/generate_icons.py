"""Regenerate the macOS add-in ribbon icons.

Office web add-ins cannot use the built-in ``imageMso`` icons of the Windows
ribbon, so the commands ship small drawn PNGs instead. Run from the repository
root with a Python that has Pillow::

    python office_plugin/macos/tools/generate_icons.py
"""

from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[3]
OUTPUT = ROOT / "office_plugin/macos/web/icons"
SIZES = (16, 32, 80)
SCALE = 8
INK = (52, 58, 64, 255)
ACCENT = (37, 99, 235, 255)
DANGER = (200, 45, 45, 255)
SERIF = (
    "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
    "/Library/Fonts/Times New Roman Bold.ttf",
    "C:/Windows/Fonts/timesbd.ttf",
)


def font(size: int) -> ImageFont.ImageFont:
    for candidate in SERIF:
        if Path(candidate).is_file():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def canvas(size: int):
    image = Image.new("RGBA", (size * SCALE, size * SCALE), (0, 0, 0, 0))
    return image, ImageDraw.Draw(image), size * SCALE


def text(draw, box, value, fill, ratio=0.82):
    left, top, right, bottom = box
    height = bottom - top
    face = font(int(height * ratio))
    bounds = draw.textbbox((0, 0), value, font=face)
    x = left + (right - left - (bounds[2] - bounds[0])) / 2 - bounds[0]
    y = top + (height - (bounds[3] - bounds[1])) / 2 - bounds[1]
    draw.text((x, y), value, font=face, fill=fill)


def plus(draw, cx, cy, r, width, fill):
    draw.rectangle((cx - r, cy - width / 2, cx + r, cy + width / 2), fill=fill)
    draw.rectangle((cx - width / 2, cy - r, cx + width / 2, cy + r), fill=fill)


def insert(draw, s):
    text(draw, (0, 0, s * 0.8, s * 0.92), "\u2211", INK, 0.95)
    plus(draw, s * 0.78, s * 0.76, s * 0.17, s * 0.09, ACCENT)


def ocr(draw, s):
    w, m, c = s * 0.07, s * 0.1, s * 0.28
    for x, y, dx, dy in ((m, m, 1, 1), (s - m, m, -1, 1), (m, s - m, 1, -1), (s - m, s - m, -1, -1)):
        draw.line((x, y, x + dx * c, y), fill=INK, width=int(w))
        draw.line((x, y, x, y + dy * c), fill=INK, width=int(w))
    text(draw, (s * 0.2, s * 0.2, s * 0.8, s * 0.8), "x", ACCENT, 1.0)


def edit(draw, s):
    angle = math.radians(45)
    length, half = s * 0.78, s * 0.11
    cx, cy = s * 0.52, s * 0.48
    ux, uy = math.cos(angle), -math.sin(angle)
    px, py = -uy, ux
    tip = (cx - ux * length / 2, cy - uy * length / 2)
    end = (cx + ux * length / 2, cy + uy * length / 2)
    neck = (tip[0] + ux * s * 0.18, tip[1] + uy * s * 0.18)
    body = [
        (neck[0] + px * half, neck[1] + py * half),
        (end[0] + px * half, end[1] + py * half),
        (end[0] - px * half, end[1] - py * half),
        (neck[0] - px * half, neck[1] - py * half),
    ]
    draw.polygon(body, fill=INK)
    draw.polygon([tip, body[0], body[3]], fill=ACCENT)
    draw.line((s * 0.1, s * 0.92, s * 0.9, s * 0.92), fill=INK, width=int(s * 0.06))


def delete(draw, s):
    w = int(s * 0.07)
    draw.rectangle((s * 0.16, s * 0.2, s * 0.84, s * 0.27), fill=INK)
    draw.rectangle((s * 0.38, s * 0.1, s * 0.62, s * 0.2), outline=INK, width=w)
    draw.rounded_rectangle((s * 0.24, s * 0.3, s * 0.76, s * 0.92), radius=s * 0.06, outline=INK, width=w)
    for x in (0.39, 0.5, 0.61):
        draw.line((s * x, s * 0.42, s * x, s * 0.8), fill=DANGER, width=w)


def format_selected(draw, s):
    text(draw, (0, s * 0.02, s, s * 0.72), "Aa", INK, 0.9)
    draw.rectangle((s * 0.12, s * 0.8, s * 0.88, s * 0.9), fill=ACCENT)


def format_all(draw, s):
    text(draw, (0, 0, s, s * 0.58), "Aa", INK, 0.9)
    for index, y in enumerate((0.66, 0.78, 0.9)):
        draw.rectangle((s * 0.12, s * (y - 0.04), s * (0.88 - index * 0.12), s * (y + 0.03)), fill=ACCENT)


def pane(draw, s):
    w = int(s * 0.07)
    draw.rounded_rectangle((s * 0.08, s * 0.14, s * 0.92, s * 0.86), radius=s * 0.06, outline=INK, width=w)
    draw.rectangle((s * 0.6, s * 0.14, s * 0.92, s * 0.86), fill=ACCENT)
    for y in (0.36, 0.5, 0.64):
        draw.line((s * 0.2, s * y, s * 0.48, s * y), fill=INK, width=w)


def settings(draw, s):
    cx = cy = s / 2
    outer, inner, teeth = s * 0.42, s * 0.31, 8
    points = []
    for index in range(teeth * 4):
        angle = 2 * math.pi * index / (teeth * 4)
        radius = outer if index % 4 in (1, 2) else inner
        points.append((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    draw.polygon(points, fill=INK)
    draw.ellipse((cx - s * 0.13, cy - s * 0.13, cx + s * 0.13, cy + s * 0.13), fill=(0, 0, 0, 0))


def help_icon(draw, s):
    w = int(s * 0.08)
    draw.ellipse((s * 0.08, s * 0.08, s * 0.92, s * 0.92), outline=INK, width=w)
    text(draw, (s * 0.2, s * 0.18, s * 0.8, s * 0.82), "?", ACCENT, 1.0)


def app(draw, s):
    draw.rounded_rectangle((0, 0, s - 1, s - 1), radius=s * 0.2, fill=ACCENT)
    text(draw, (0, 0, s, s * 0.96), "\u2211", (255, 255, 255, 255), 0.85)


ICONS = {
    "insert": insert,
    "ocr": ocr,
    "edit": edit,
    "delete": delete,
    "format": format_selected,
    "format-all": format_all,
    "pane": pane,
    "settings": settings,
    "help": help_icon,
}


def render(painter, size: int) -> Image.Image:
    image, draw, scaled = canvas(size)
    painter(draw, scaled)
    return image.resize((size, size), Image.Resampling.LANCZOS)


def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    for name, painter in ICONS.items():
        for size in SIZES:
            render(painter, size).save(OUTPUT / f"{name}-{size}.png", optimize=True)
    for size in (32, 64, 80):
        render(app, size).save(OUTPUT / f"app-{size}.png", optimize=True)


if __name__ == "__main__":
    main()
