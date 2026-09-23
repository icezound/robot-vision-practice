# -*- coding: utf-8 -*-
"""Build comparison figures for the practical training report."""
import os
from PIL import Image, ImageDraw, ImageFont
import numpy as np

SRC1 = r"D:\matlab chuchun\w\1"
SRC2 = r"D:\matlab chuchun\w\2"
SRC3 = r"D:\matlab chuchun\w\3"
OUT = r"D:\matlab chuchun\w\_assets"
os.makedirs(OUT, exist_ok=True)

FONT_PATH = r"C:\Windows\Fonts\simhei.ttf"
LABEL_FONT = ImageFont.truetype(FONT_PATH, 30)

ORANGE = np.array([242, 141, 38])


def load(path):
    return Image.open(path).convert("RGB")


def graphics_area(im):
    """Cut away the HALCON variable window and keep only the graphics window."""
    a = np.asarray(im).astype(int)
    d = np.abs(a - ORANGE).sum(axis=2)
    frac = (d < 60).mean(axis=1)
    rows = np.where(frac > 0.9)[0]
    bottom = rows[0] - 8 if len(rows) else a.shape[0]
    gra = a[:bottom]
    # trim the window frame, keep the viewport
    return Image.fromarray(gra[4:].astype("uint8"))


def panel(img, label, box=None):
    """Return (label_height, bordered image)."""
    if box:
        img = img.crop(box)
    return img


def compose(panels, out_path, gap=26, margin=18, pad=8, label_h=52):
    """panels: list of (PIL image, label). Side by side with captions above."""
    heights = [im.height for im, _ in panels]
    widths = [im.width for im, _ in panels]
    h = max(heights)
    W = margin * 2 + sum(widths) + gap * (len(panels) - 1)
    H = margin * 2 + label_h + h
    canvas = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(canvas)
    x = margin
    for (im, label), w in zip(panels, widths):
        if label:
            tw = d.textlength(label, font=LABEL_FONT)
            d.text((x + (w - tw) / 2, margin + 6), label, fill="black", font=LABEL_FONT)
        y = margin + label_h
        canvas.paste(im, (x, y + (h - im.height) // 2))
        d.rectangle(
            [x - 1, y + (h - im.height) // 2 - 1,
             x + w, y + (h - im.height) // 2 + im.height],
            outline=(150, 150, 150), width=1,
        )
        x += w + gap
    canvas.save(out_path)
    return canvas.size


def halcon_norm(im, size=(1150, 510)):
    return graphics_area(im).resize(size, Image.LANCZOS)


NO_CROP = (320, 0, 830, 510)


def main():
    # ---------- HALCON panels ----------
    hal = {
        "orig": "Snipaste_2026-09-21_15-56-04.png",
        "amp": "Snipaste_2026-09-21_15-56-14.png",
        "binary": "Snipaste_2026-09-21_15-56-19.png",
        "conn": "Snipaste_2026-09-21_15-56-26.png",
        "valid": "Snipaste_2026-09-21_15-56-30.png",
        "vars": "Snipaste_2026-09-21_15-56-35.png",
        "code": "Snipaste_2026-09-21_15-55-48.png",
    }
    hp = {k: halcon_norm(load(os.path.join(SRC1, v))).crop(NO_CROP)
          for k, v in hal.items() if k not in ("vars", "code")}

    compose([(hp["orig"], "原图"), (hp["amp"], "均值滤波+边缘检测(幅值)")],
            os.path.join(OUT, "f4_1.png"))
    compose([(hp["orig"], "原图"), (hp["binary"], "边缘二值化区域")],
            os.path.join(OUT, "f4_2.png"))
    compose([(hp["conn"], "连通区域(Blob)"), (hp["valid"], "面积筛选后有效区域")],
            os.path.join(OUT, "f4_3.png"))

    # HALCON code
    code = load(os.path.join(SRC1, hal["code"]))
    code.save(os.path.join(OUT, "f3_1.png"))
    # HALCON variable window
    load(os.path.join(SRC1, hal["vars"])).save(os.path.join(OUT, "f4_4.png"))

    # ---------- OpenCV panels ----------
    cv = {
        "orig": "Snipaste_2026-09-21_15-58-21.png",
        "gray": "Snipaste_2026-09-21_15-58-25.png",
        "canny": "Snipaste_2026-09-21_15-58-31.png",
        "binary": "Snipaste_2026-09-21_15-58-35.png",
        "contour": "Snipaste_2026-09-21_15-58-42.png",
        "log": "Snipaste_2026-09-21_15-58-46.png",
        "code": "Snipaste_2026-09-21_15-59-00.png",
    }
    cp = {k: load(os.path.join(SRC2, v)) for k, v in cv.items()}

    compose([(cp["orig"], "原图(彩色)"), (cp["gray"], "灰度图像")],
            os.path.join(OUT, "f4_5.png"))
    compose([(cp["gray"], "灰度图像"), (cp["log"], "对数(Log)变换后图像")],
            os.path.join(OUT, "f4_6.png"))
    compose([(cp["gray"], "灰度图像"), (cp["binary"], "二值化图像")],
            os.path.join(OUT, "f4_7.png"))
    compose([(cp["gray"], "灰度图像"), (cp["canny"], "Canny 边缘检测")],
            os.path.join(OUT, "f4_8.png"))
    compose([(cp["orig"], "原图"), (cp["contour"], "轮廓检测并绘制")],
            os.path.join(OUT, "f4_9.png"))
    cp["code"].save(os.path.join(OUT, "f3_2.png"))

    # ---------- PyTorch ----------
    load(os.path.join(SRC3, "Snipaste_2026-09-21_15-59-31.png")).save(
        os.path.join(OUT, "f3_3.png"))
    term = load(os.path.join(SRC3, "Snipaste_2026-09-21_15-59-43.png"))
    a = np.asarray(term).astype(int)
    bg = a[5, -5]
    mask = np.abs(a - bg).sum(axis=2) > 30
    ys, xs = np.nonzero(mask)
    term.crop((0, 0, min(term.width, xs.max() + 12),
               min(term.height, ys.max() + 12))).save(
        os.path.join(OUT, "f4_10.png"))

    for f in sorted(os.listdir(OUT)):
        im = Image.open(os.path.join(OUT, f))
        print(f, im.size)


if __name__ == "__main__":
    main()
