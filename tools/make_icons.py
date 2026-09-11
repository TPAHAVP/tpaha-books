"""Generate the PWA icons: a TPAHA-green rounded square with a white ledger mark. Run: python tools/make_icons.py"""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "site", "icons")
os.makedirs(OUT, exist_ok=True)
GREEN = (16, 124, 65)      # #107C41
DARK = (57, 88, 76)        # #39584C
WHITE = (255, 255, 255)


def draw_icon(size, maskable=False):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = 0 if maskable else int(size * 0.06)
    radius = int(size * (0.0 if maskable else 0.22))
    d.rounded_rectangle([pad, pad, size - pad, size - pad], radius=radius, fill=GREEN)
    # A simple "ledger page" mark: white sheet with three lines and a dark header band.
    x0, y0, x1, y1 = size * 0.28, size * 0.22, size * 0.72, size * 0.78
    d.rounded_rectangle([x0, y0, x1, y1], radius=int(size * 0.04), fill=WHITE)
    d.rectangle([x0, y0, x1, y0 + size * 0.10], fill=DARK)
    lw = max(2, int(size * 0.03))
    for i, frac in enumerate((0.42, 0.53, 0.64)):
        y = size * frac
        d.line([(x0 + size * 0.06, y), (x1 - size * 0.06 - (size * 0.10 if i == 2 else 0), y)], fill=DARK, width=lw)
    return img


for name, size, maskable in [("icon-192.png", 192, False), ("icon-512.png", 512, False), ("icon-maskable-512.png", 512, True), ("apple-touch-icon.png", 180, True)]:
    img = draw_icon(size, maskable)
    if name == "apple-touch-icon.png":
        bg = Image.new("RGBA", img.size, GREEN + (255,))
        bg.alpha_composite(img)
        img = bg.convert("RGB")
    img.save(os.path.join(OUT, name))
    print("wrote", os.path.join("site", "icons", name), img.size)
