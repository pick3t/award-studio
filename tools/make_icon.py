"""生成应用图标 assets/icon.ico。"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
ICON = ROOT / "assets" / "icon.ico"


def main() -> None:
    size = 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle([6, 6, size - 6, size - 6], radius=58, fill=(79, 70, 229))

    font = None
    for candidate in ("C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/segoeuib.ttf"):
        try:
            font = ImageFont.truetype(candidate, 148)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()

    draw.text((size / 2, size / 2 + 6), "A", font=font, fill="white", anchor="mm")

    ICON.parent.mkdir(parents=True, exist_ok=True)
    image.save(ICON, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"saved {ICON}")


if __name__ == "__main__":
    main()
