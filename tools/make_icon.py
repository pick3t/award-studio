"""生成应用图标：候选方案预览 + 导出 .ico。

    python tools/make_icon.py preview     # 出对比图 docs/icon_options.png
    python tools/make_icon.py export b    # 把方案 b 写入 assets/icon.ico

图标在 4096 画布上绘制（4 倍超采样）再缩到 1024，保证小尺寸边缘干净。
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
DOCS = ROOT / "docs"

N = 1024  # 图标设计坐标（1x）
S = 4     # 超采样倍数
W = N * S

INDIGO_TOP = (109, 102, 240)
INDIGO_BOTTOM = (67, 56, 202)
INDIGO = (79, 70, 229)
INDIGO_LIGHT = (129, 140, 248)
WHITE = (255, 255, 255)
PAPER = (252, 251, 255)
LINE = (199, 203, 224)
GOLD = (217, 165, 33)
GOLD_LIGHT = (242, 193, 78)
GOLD_DARK = (168, 123, 11)

ICO_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def px(*values: float) -> tuple:
    return tuple(int(round(v * S)) for v in values)


def star(draw: ImageDraw.ImageDraw, cx: float, cy: float, r_out: float,
         r_in: float, fill, points: int = 5) -> None:
    coords = []
    for index in range(points * 2):
        radius = r_out if index % 2 == 0 else r_in
        angle = math.radians(-90 + index * 180 / points)
        coords.append((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    draw.polygon([(x * S, y * S) for x, y in coords], fill=fill)


def ribbon(cx: float, cy: float, angle: float, width: float, length: float,
           notch: float = 0.30) -> list[tuple[float, float]]:
    """从 (cx, cy) 往下拉的一条绶带，末端带 V 形缺口。angle 为正度数。"""
    a = math.radians(angle)
    dx, dy = math.sin(a), math.cos(a)
    nx, ny = math.cos(a), -math.sin(a)
    x0, y0 = cx - nx * width / 2, cy - ny * width / 2
    x1, y1 = cx + nx * width / 2, cy + ny * width / 2
    x2, y2 = x1 + dx * length, y1 + dy * length
    x3, y3 = x0 + dx * length, y0 + dy * length
    mx, my = (x2 + x3) / 2 - dx * length * notch, (y2 + y3) / 2 - dy * length * notch
    return [(x0, y0), (x1, y1), (x2, y2), (mx, my), (x3, y3)]


def make_tile() -> Image.Image:
    """圆角方形底：靛蓝渐变 + 左上高光。"""
    gradient = Image.new("RGB", (1, N))
    for y in range(N):
        t = y / (N - 1)
        gradient.putpixel((0, y), tuple(
            int(INDIGO_TOP[i] + (INDIGO_BOTTOM[i] - INDIGO_TOP[i]) * t) for i in range(3)
        ))
    tile = gradient.resize((W, W), Image.BICUBIC).convert("RGBA")

    glow = Image.new("L", (W, W), 0)
    gd = ImageDraw.Draw(glow)
    gd.ellipse(px(-260, -420, 760, 560), fill=46)
    glow = glow.filter(ImageFilter.GaussianBlur(120 * S // 4))
    tile.paste(Image.new("RGBA", (W, W), WHITE), (0, 0), glow)

    mask = Image.new("L", (W, W), 0)
    ImageDraw.Draw(mask).rounded_rectangle(px(24, 24, N - 24, N - 24), radius=px(212)[0], fill=255)
    out = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    out.paste(tile, (0, 0), mask)
    return out


def drop_shadow(shape: Image.Image, blur: int = 26, opacity: int = 90) -> Image.Image:
    """按 alpha 生成投影层（同尺寸）。"""
    alpha = shape.split()[-1].filter(ImageFilter.GaussianBlur(blur * S // 4))
    shadow = Image.new("RGBA", shape.size, (15, 12, 60, 0))
    shadow.putalpha(alpha.point(lambda v: v * opacity // 255))
    return shadow


def sheet_layer(width: float, height: float, radius: float) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    """在独立图层上画一张纸，返回 (图层, 画笔)，坐标为纸内相对坐标。"""
    pad = 40
    layer = Image.new("RGBA", (int((width + pad * 2) * S), int((height + pad * 2) * S)), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    return layer, draw, pad


def rounded(draw, box, radius, fill):
    draw.rounded_rectangle(px(*box), radius=px(radius)[0], fill=fill)


# ---------------------------------------------------------------- 方案 A
def design_a() -> Image.Image:
    """奖状（微倾斜）+ 金色印章绶带。"""
    base = make_tile()

    sw, sh, sr = 366, 476, 28
    layer, draw, pad = sheet_layer(sw, sh, sr)
    rounded(draw, (pad, pad, pad + sw, pad + sh), sr, PAPER)
    # 纸上的抬头条 + 文字条
    rounded(draw, (pad + 48, pad + 58, pad + sw - 48, pad + 80), 11, INDIGO)
    for i, w in enumerate((sw - 108, sw - 140, sw - 178)):
        top = pad + 142 + i * 52
        rounded(draw, (pad + 48, top, pad + 48 + w, top + 18), 9, LINE)
    rounded(draw, (pad + 48, pad + sh - 92, pad + 142, pad + sh - 74), 9, LINE)

    rotated = layer.rotate(7, resample=Image.BICUBIC, expand=True)
    cx, cy = 436, 408
    pos = (int(cx * S - rotated.width / 2), int(cy * S - rotated.height / 2))
    base.alpha_composite(drop_shadow(rotated, blur=30, opacity=110), (pos[0] + 10 * S, pos[1] + 24 * S))
    base.alpha_composite(rotated, pos)

    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    # 绶带（先画，压在印章下面）
    for angle, color in ((-34, GOLD_DARK), (34, GOLD)):
        points = ribbon(722, 690, angle, 96, 186)
        d.polygon([(x * S, y * S) for x, y in points], fill=color + (255,))
    # 印章
    d.ellipse(px(606, 574, 838, 806), fill=GOLD_DARK + (255,))
    d.ellipse(px(620, 588, 824, 792), fill=GOLD + (255,))
    d.ellipse(px(650, 618, 794, 762), fill=GOLD_LIGHT + (255,))
    star(d, 722, 690, 50, 21, GOLD_DARK + (255,))
    base.alpha_composite(overlay)
    return base


# ---------------------------------------------------------------- 方案 B
def design_b() -> Image.Image:
    """上升排名柱 + 金星（最能体现“按分数排名取前几”）。"""
    base = make_tile()
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    baseline = 806
    bars = [(232, 306, 168, 232), (426, 414, 168, 246), (620, 520, 168, 255)]
    for left, height, width, alpha in bars:
        rounded(d, (left, baseline - height, left + width, baseline), 34, WHITE + (alpha,))

    #  tallest 柱顶的金星徽章
    cx, cy, r = 704, 236, 86
    d.ellipse(px(cx - r, cy - r, cx + r, cy + r), fill=GOLD + (255,))
    d.ellipse(px(cx - r + 12, cy - r + 12, cx + r - 12, cy + r - 12), fill=GOLD_LIGHT + (255,))
    star(d, cx, cy + 4, 50, 21, GOLD_DARK + (255,))

    base.alpha_composite(overlay)
    return base


# ---------------------------------------------------------------- 方案 C
def design_c() -> Image.Image:
    """奖状里画出排名柱 + 金印，把两件事合成一个图案。"""
    base = make_tile()
    sw, sh, sr = 500, 600, 36
    layer, draw, pad = sheet_layer(sw, sh, sr)
    rounded(draw, (pad, pad, pad + sw, pad + sh), sr, PAPER)

    baseline = pad + sh - 150
    for i, (left, height, color) in enumerate((
        (pad + 96, 118, INDIGO_LIGHT), (pad + 196, 176, INDIGO), (pad + 296, 238, (67, 56, 202))
    )):
        rounded(draw, (left, baseline - height, left + 72, baseline), 20, color)

    rounded(draw, (pad + 96, pad + 78, pad + sw - 96, pad + 100), 11, LINE)

    rotated = layer
    cx, cy = 470, 470
    pos = (int(cx * S - rotated.width / 2), int(cy * S - rotated.height / 2))
    base.alpha_composite(drop_shadow(rotated, blur=30, opacity=110), (pos[0] + 8 * S, pos[1] + 24 * S))
    base.alpha_composite(rotated, pos)

    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    cx, cy, r = 716, 704, 92
    d.ellipse(px(cx - r, cy - r, cx + r, cy + r), fill=GOLD_DARK + (255,))
    d.ellipse(px(cx - r + 12, cy - r + 12, cx + r - 12, cy + r - 12), fill=GOLD + (255,))
    d.ellipse(px(cx - r + 36, cy - r + 36, cx + r - 36, cy + r - 36), fill=GOLD_LIGHT + (255,))
    star(d, cx, cy + 4, 42, 17, GOLD_DARK + (255,))
    base.alpha_composite(overlay)
    return base


# ---------------------------------------------------------------- 方案 D
def design_d() -> Image.Image:
    """奖章 + 绶带（最简洁、16px 也认得出）。"""
    base = make_tile()
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    #  绶带（两条斜带）
    for angle, color in ((-24, (199, 210, 254)), (24, (129, 140, 248))):
        points = ribbon(512, 470, angle, 104, 300)
        d.polygon([(x * S, y * S) for x, y in points], fill=color + (255,))

    #  奖章圆盘
    d.ellipse(px(348, 268, 676, 596), fill=GOLD_DARK + (255,))
    d.ellipse(px(364, 284, 660, 580), fill=GOLD + (255,))
    d.ellipse(px(392, 312, 632, 552), fill=GOLD_LIGHT + (255,))
    star(d, 512, 436, 92, 40, GOLD_DARK + (255,))

    base.alpha_composite(overlay)
    return base


DESIGNS = {"a": design_a, "b": design_b, "c": design_c, "d": design_d}


def render(key: str) -> Image.Image:
    big = DESIGNS[key]()
    return big.resize((N, N), Image.LANCZOS)


def preview() -> None:
    keys = ["a", "b", "c", "d"]
    icons = {k: render(k) for k in keys}

    pad = 44
    row_h = 300
    big = 256
    cells = [128, 48, 32, 16]
    width = pad * 2 + big + 60 + sum((size + 16) * 2 + 22 for size in cells) + 30
    height = pad + 70 + row_h * len(keys) + 40

    canvas = Image.new("RGB", (width, height), (243, 245, 250))
    draw = ImageDraw.Draw(canvas)
    from PIL import ImageFont
    font_path = "C:/Windows/Fonts/msyhbd.ttc"
    try:
        title_font = ImageFont.truetype(font_path, 28)
        small_font = ImageFont.truetype(font_path, 16)
        tiny_font = ImageFont.truetype(font_path, 13)
    except OSError:
        title_font = small_font = tiny_font = ImageFont.load_default()

    draw.text((pad, pad - 12), "应用图标候选：大图预览 + 各尺寸在浅色/深色背景上的实际像素",
              font=title_font, fill=(17, 24, 39))

    names = {"a": "A　奖状 + 金印", "b": "B　排名柱 + 金星",
             "c": "C　奖状内嵌排名柱", "d": "D　奖章 + 绶带"}

    for index, key in enumerate(keys):
        top = pad + 66 + index * row_h
        icon = icons[key]
        big_img = icon.resize((big, big), Image.LANCZOS)
        canvas.paste(big_img, (pad, top), big_img)
        draw.text((pad, top + big + 12), names[key], font=small_font, fill=(17, 24, 39))

        x = pad + big + 60
        for size in cells:
            cell = size + 16
            for row, bg in ((0, (255, 255, 255)), (1, (17, 24, 39))):
                box = Image.new("RGB", (cell, cell), bg)
                small = icon.resize((size, size), Image.LANCZOS)
                box.paste(small, (8, 8), small)
                canvas.paste(box, (x, top + 8 + row * (cell + 10)))
                draw.rectangle([x, top + 8 + row * (cell + 10),
                                x + cell - 1, top + 8 + row * (cell + 10) + cell - 1],
                               outline=(226, 232, 240))
            draw.text((x, top + 20 + 2 * (128 + 16) + 18), f"{size}px", font=tiny_font, fill=(120, 128, 145))
            x += cell + 22

    DOCS.mkdir(parents=True, exist_ok=True)
    out = DOCS / "icon_options.png"
    canvas.save(out, optimize=True)
    print("saved", out, canvas.size)


def _dib_bytes(frame: Image.Image) -> bytes:
    """把一帧图标写成 ICO 里的 32 位 BMP（DIB）数据，小尺寸必须用它。"""
    import struct

    width, height = frame.size
    pixels = frame.load()
    rows = []
    for y in range(height - 1, -1, -1):
        row = bytearray()
        for x in range(width):
            r, g, b, a = pixels[x, y]
            row += bytes((b, g, r, a))
        rows.append(bytes(row))
    xor = b"".join(rows)
    mask_row = ((width + 31) // 32) * 4
    mask = b"\x00" * (mask_row * height)
    header = struct.pack("<IiiHHIIiiII", 40, width, height * 2, 1, 32, 0,
                         len(xor) + len(mask), 0, 0, 0, 0)
    return header + xor + mask


def save_ico(icon: Image.Image, path: Path, sizes: list[tuple[int, int]]) -> None:
    """自己拼 .ico：≤128 用 BMP 数据，256 用 PNG。

    Pillow 存出来的每个尺寸都是 PNG，Windows 的 LoadIcon 对 PNG 小图支持不好，
    结果就是任务栏/标题栏回退成默认图标。
    """
    import io
    import struct

    frames: list[tuple[int, bytes]] = []
    for size, _ in sizes:
        frame = icon.resize((size, size), Image.LANCZOS)
        if size >= 256:
            buffer = io.BytesIO()
            frame.save(buffer, format="PNG")
            frames.append((size, buffer.getvalue()))
        else:
            frames.append((size, _dib_bytes(frame)))

    header = struct.pack("<HHH", 0, 1, len(frames))
    offset = len(header) + 16 * len(frames)
    directory = b""
    payload = b""
    for size, data in frames:
        directory += struct.pack("<BBBBHHII", size % 256, size % 256, 0, 0, 1, 32,
                                 len(data), offset)
        payload += data
        offset += len(data)
    path.write_bytes(header + directory + payload)


def export(key: str) -> None:
    icon = render(key)
    ASSETS.mkdir(parents=True, exist_ok=True)
    ico = ASSETS / "icon.ico"
    save_ico(icon, ico, ICO_SIZES)
    png = ASSETS / "icon.png"
    icon.save(png)
    print(f"saved {ico} ({len(ICO_SIZES)} sizes) and {png}")


def main() -> None:
    mode = sys.argv[1] if len(sys.argv) > 1 else "preview"
    if mode == "preview":
        preview()
    elif mode == "export":
        key = sys.argv[2] if len(sys.argv) > 2 else "b"
        if key not in DESIGNS:
            raise SystemExit(f"方案只能是 {sorted(DESIGNS)}")
        export(key)
    elif mode == "all":
        for key in DESIGNS:
            export(key)
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main()
