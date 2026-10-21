"""把奖状版面渲染成 PNG（效果预览用，不参与软件运行）。

按 A4 横向 150dpi 复刻模板的排版参数：页面边框、标题字距、正文字号与行距。
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent

W, H = 1754, 1240
CM = W / 29.7
DPI = 150

FONT_SONG = "C:/Windows/Fonts/simsun.ttc"
FONT_KAI = "C:/Windows/Fonts/simkai.ttf"
FONT_HEI = "C:/Windows/Fonts/simhei.ttf"
FONT_TIMES = "C:/Windows/Fonts/times.ttf"
FONT_TIMES_B = "C:/Windows/Fonts/timesbd.ttf"

INK = (26, 26, 26)
RED = (156, 43, 43)
GOLD = (176, 141, 87)
GREY = (140, 140, 140)


def px(points: float) -> float:
    return points * DPI / 72


def load(path: str, size: float) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, int(round(size)))


def run_width(text: str, font: ImageFont.FreeTypeFont, spacing: float = 0) -> float:
    return sum(font.getlength(ch) + spacing for ch in text)


def draw_runs(draw: ImageDraw.ImageDraw, x: float, baseline: float, runs: list[dict],
              spacing: float = 0) -> float:
    for run in runs:
        font = load(run["font"], px(run["size"]))
        for ch in run["text"]:
            draw.text((x, baseline), ch, font=font, fill=run["color"], anchor="ls")
            width = font.getlength(ch) + spacing + run.get("spacing", 0)
            if run.get("underline"):
                draw.line(
                    [(x, baseline + px(run["size"]) * 0.14),
                     (x + width - 1, baseline + px(run["size"]) * 0.14)],
                    fill=run["color"], width=max(1, int(px(run["size"]) * 0.055)),
                )
            x += width
    return x


def render(out: Path) -> None:
    image = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(image)

    # 双线金色边框
    inset = 0.85 * CM
    for offset, width in ((0, 5), (11, 2)):
        draw.rectangle(
            [inset + offset, inset + offset, W - inset - offset, H - inset - offset],
            outline=GOLD, width=width,
        )

    left, right = 3.2 * CM, W - 3.2 * CM
    top = 2.4 * CM
    y = top

    # 证书编号
    y += px(30)
    no_font = load(FONT_KAI, px(10.5))
    text = "证书编号：AS-C024"
    draw.text((right, y + px(10.5)), text, font=no_font, fill=GREY, anchor="rs")
    y += px(10.5) + px(26)

    # 主标题
    title_size = 42
    title_font = load(FONT_SONG, px(title_size))
    title = "获 奖 证 书"
    spacing = px(8)
    width = run_width(title, title_font, spacing)
    draw_runs(draw, (W - width) / 2, y + px(title_size) * 0.9,
              [{"text": title, "font": FONT_SONG, "size": title_size, "color": RED,
                "spacing": spacing}])
    y += px(title_size) * 1.25 + px(10)

    # 英文副标题
    en_font = load(FONT_TIMES, px(11))
    en_text = "CERTIFICATE  OF  AWARD"
    en_width = run_width(en_text, en_font, px(3))
    draw_runs(draw, (W - en_width) / 2, y + px(11),
              [{"text": en_text, "font": FONT_TIMES, "size": 11, "color": GOLD,
                "spacing": px(3)}])
    y += px(11) + px(40)

    # 正文
    body = 18
    name_size = 20
    indent = px(32)
    line_gap = px(body) * 2.4
    baseline = y + px(body)

    runs = [
        {"text": "兹有　", "font": FONT_KAI, "size": body, "color": INK},
        {"text": "何清越", "font": FONT_HEI, "size": name_size, "color": INK,
         "bold": True, "underline": True},
        {"text": "　同学，在「2026 校园科技创新大赛」中表现优异，",
         "font": FONT_KAI, "size": body, "color": INK},
    ]
    draw_runs(draw, left + indent, baseline, runs)
    baseline += line_gap

    runs2 = [
        {"text": "荣获 ", "font": FONT_KAI, "size": body, "color": INK},
        {"text": "一等奖", "font": FONT_HEI, "size": name_size, "color": RED, "bold": True},
        {"text": "。特发此证，以资鼓励。", "font": FONT_KAI, "size": body, "color": INK},
    ]
    draw_runs(draw, left + indent, baseline, runs2)
    y = baseline + px(body) * 0.4

    # 装饰
    y += px(36)
    deco_font = load(FONT_SONG, px(11))
    deco = "◆　◆　◆"
    deco_width = run_width(deco, deco_font)
    draw.text(((W - deco_width) / 2, y + px(11)), deco, font=deco_font, fill=GOLD,
              anchor="ls")
    y += px(11) + px(64)

    # 落款
    sign_font = load(FONT_KAI, px(15))
    for line in ("2026 校园科技创新大赛 组织委员会", "2026 年 10 月 21 日"):
        draw.text((right - 1.2 * CM, y + px(15)), line, font=sign_font, fill=INK,
                  anchor="rs")
        y += px(15) + px(6)

    out.parent.mkdir(parents=True, exist_ok=True)
    image.save(out, quality=95)
    print(f"saved {out} ({W}x{H})")


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "docs" / "证书预览.png"
    render(out)


if __name__ == "__main__":
    main()
