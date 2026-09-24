#!/usr/bin/env python3
"""Generate every icon/image asset Affice needs from the two source logos.

Sources (kept in build/logo-source/):
  affice-icon.png  - white cube on the Affice blue square (the app icon)
  affice-logo.png  - blue cube on white (the brand mark)

Outputs:
  build/icon.png, build/icon.ico            - app icon (electron-builder)
  build/icons/<size>x<size>.png             - Linux icon theme sizes
  build/installerSidebar.bmp, build/uninstallerSidebar.bmp, build/installerHeader.bmp
  src/assets/logo-mark.png                  - transparent brand mark for the UI
  src/assets/app-icon.png                   - rounded app icon for the UI
  public/favicon.png

Requires Pillow:  pip install Pillow
"""
from __future__ import annotations

import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "build", "logo-source")
BUILD = os.path.join(ROOT, "build")
ASSETS = os.path.join(ROOT, "src", "assets")
PUBLIC = os.path.join(ROOT, "public")

BRAND = (0, 79, 255)


def rounded_mask(size: int, radius_ratio: float = 0.2) -> Image.Image:
    """Anti-aliased rounded-square mask (drawn at 4x then downsampled)."""
    big = size * 4
    m = Image.new("L", (big, big), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle((0, 0, big - 1, big - 1), radius=int(big * radius_ratio), fill=255)
    return m.resize((size, size), Image.LANCZOS)


def app_icon(size: int) -> Image.Image:
    """The provided icon with softly rounded corners (transparent outside)."""
    src = Image.open(os.path.join(SRC, "affice-icon.png")).convert("RGBA")
    img = src.resize((size, size), Image.LANCZOS)
    img.putalpha(rounded_mask(size))
    return img


def transparent_mark() -> Image.Image:
    """Remove the white background from the blue logo, keeping soft edges.

    Each edge pixel is a blend  P = a*C + (1-a)*White  where the logo colour C
    lies on the gradient (0,79,255)->(32,134,255), for which R ~= 0.582*(G-79).
    Solving the red and green channels together gives the coverage a directly.
    """
    src = Image.open(os.path.join(SRC, "affice-logo.png")).convert("RGB")
    w, h = src.size
    out = Image.new("RGBA", (w, h))
    sp = src.load()
    op = out.load()
    for y in range(h):
        for x in range(w):
            r, g, b = sp[x, y]
            u = 255 - r
            v = 255 - g
            a = (u - 0.582 * v) / 152.57
            if a <= 0.004:
                op[x, y] = (0, 0, 0, 0)
                continue
            a = min(1.0, a)
            cr = (r - (1 - a) * 255) / a
            cg = (g - (1 - a) * 255) / a
            cb = (b - (1 - a) * 255) / a
            op[x, y] = (
                max(0, min(255, round(cr))),
                max(0, min(255, round(cg))),
                max(0, min(255, round(cb))),
                round(a * 255),
            )
    bbox = out.getbbox()
    out = out.crop(bbox)
    # pad to a square with a little breathing room
    side = int(max(out.size) * 1.04)
    sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    sq.paste(out, ((side - out.size[0]) // 2, (side - out.size[1]) // 2))
    return sq


def load_font(size: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    candidates = [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
    ]
    for c in candidates:
        if os.path.exists(c):
            return ImageFont.truetype(c, size)
    return ImageFont.load_default()


def vertical_gradient(w: int, h: int, top, bottom) -> Image.Image:
    img = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / max(1, h - 1)
        d.line([(0, y), (w, y)], fill=tuple(round(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return img


def installer_sidebar(mark_white: Image.Image) -> Image.Image:
    # NSIS welcome/finish page sidebar: 164x314 (drawn 2x, downsampled)
    W, H = 164 * 2, 314 * 2
    img = vertical_gradient(W, H, (0, 79, 255), (0, 40, 170)).convert("RGBA")
    # soft glow
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse((-W * 0.4, -H * 0.1, W * 1.3, H * 0.6), fill=(80, 150, 255, 90))
    glow = glow.filter(ImageFilter.GaussianBlur(60))
    img = Image.alpha_composite(img, glow)
    m = mark_white.resize((int(W * 0.62), int(W * 0.62)), Image.LANCZOS)
    img.alpha_composite(m, ((W - m.size[0]) // 2, int(H * 0.2)))
    d = ImageDraw.Draw(img)
    f = load_font(58)
    tw = d.textlength("Affice", font=f)
    d.text(((W - tw) / 2, H * 0.6), "Affice", font=f, fill=(255, 255, 255))
    f2 = load_font(22, bold=False)
    for i, line in enumerate(["Documents · Sheets · Slides", "Free & open source"]):
        tw = d.textlength(line, font=f2)
        d.text(((W - tw) / 2, H * 0.72 + i * 34), line, font=f2, fill=(210, 225, 255))
    return img.resize((164, 314), Image.LANCZOS).convert("RGB")


def installer_header(mark: Image.Image) -> Image.Image:
    # NSIS header image: 150x57
    W, H = 150 * 2, 57 * 2
    img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    m = mark.resize((int(H * 0.8), int(H * 0.8)), Image.LANCZOS)
    img.alpha_composite(m, (W - m.size[0] - 14, (H - m.size[1]) // 2))
    return img.resize((150, 57), Image.LANCZOS).convert("RGB")


KIND_COLORS = {
    "doc": ((47, 109, 255), (18, 70, 230)),
    "sheet": ((28, 184, 96), (10, 132, 62)),
    "slides": ((255, 120, 50), (226, 76, 18)),
}


def file_icon(kind: str, mark_white: Image.Image, size: int = 256) -> Image.Image:
    """A document page with a folded corner, type-specific content and an Affice badge."""
    S = size * 4
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    x0, y0, x1, y1 = int(S * 0.18), int(S * 0.05), int(S * 0.86), int(S * 0.95)
    fold = int(S * 0.2)
    r = int(S * 0.045)
    # shadow
    sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle((x0 + 6, y0 + 14, x1 + 6, y1 + 14), radius=r, fill=(0, 20, 60, 70))
    img = Image.alpha_composite(img, sh.filter(ImageFilter.GaussianBlur(S * 0.02)))
    d = ImageDraw.Draw(img)
    page = [(x0 + r, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1 - r), (x1 - r, y1), (x0 + r, y1), (x0, y1 - r), (x0, y0 + r)]
    d.polygon(page, fill=(255, 255, 255, 255), outline=(196, 205, 222, 255))
    d.rounded_rectangle((x0, y0, x0 + 2 * r, y0 + 2 * r), radius=r, fill=(255, 255, 255, 255))
    d.rounded_rectangle((x0, y1 - 2 * r, x0 + 2 * r, y1), radius=r, fill=(255, 255, 255, 255))
    d.rounded_rectangle((x1 - 2 * r, y1 - 2 * r, x1, y1), radius=r, fill=(255, 255, 255, 255))
    d.line([(x0 + r, y0), (x1 - fold, y0)], fill=(196, 205, 222, 255), width=6)
    d.line([(x0, y0 + r), (x0, y1 - r)], fill=(196, 205, 222, 255), width=6)
    d.line([(x0 + r, y1), (x1 - r, y1)], fill=(196, 205, 222, 255), width=6)
    d.line([(x1, y0 + fold), (x1, y1 - r)], fill=(196, 205, 222, 255), width=6)
    d.polygon([(x1 - fold, y0), (x1 - fold, y0 + fold), (x1, y0 + fold)], fill=(226, 232, 244, 255), outline=(196, 205, 222, 255))
    c1, c2 = KIND_COLORS[kind]
    light = tuple(min(255, int(c + (255 - c) * 0.72)) for c in c1)
    cx0, cx1 = int(S * 0.3), int(S * 0.76)
    if kind == "doc":
        for i in range(5):
            y = int(S * 0.3) + i * int(S * 0.075)
            w = cx1 if i % 3 != 2 else int(S * 0.62)
            d.rounded_rectangle((cx0, y, w, y + int(S * 0.028)), radius=int(S * 0.014), fill=light + (255,))
    elif kind == "sheet":
        gx0, gy0, gx1, gy1 = cx0, int(S * 0.28), cx1, int(S * 0.62)
        d.rectangle((gx0, gy0, gx1, gy0 + int(S * 0.07)), fill=light + (255,))
        for i in range(5):
            y = gy0 + i * (gy1 - gy0) // 4
            d.line([(gx0, y), (gx1, y)], fill=light + (255,), width=8)
        for j in range(4):
            x = gx0 + j * (gx1 - gx0) // 3
            d.line([(x, gy0), (x, gy1)], fill=light + (255,), width=8)
    else:
        sx0, sy0, sx1, sy1 = cx0, int(S * 0.28), cx1, int(S * 0.58)
        d.rounded_rectangle((sx0, sy0, sx1, sy1), radius=int(S * 0.02), outline=light + (255,), width=10)
        bw = (sx1 - sx0) // 7
        for i, hgt in enumerate([0.35, 0.6, 0.85]):
            bx = sx0 + bw * (1 + i * 2)
            d.rectangle((bx, sy1 - int((sy1 - sy0) * hgt * 0.8) - 14, bx + bw, sy1 - 14), fill=light + (255,))
    # badge
    bx0, by0, bx1, by1 = int(S * 0.06), int(S * 0.5), int(S * 0.56), int(S * 0.9)
    badge = Image.new("RGBA", (bx1 - bx0, by1 - by0), (0, 0, 0, 0))
    grad = vertical_gradient(bx1 - bx0, by1 - by0, c1, c2).convert("RGBA")
    mask = Image.new("L", badge.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, badge.size[0] - 1, badge.size[1] - 1), radius=int(S * 0.07), fill=255)
    badge.paste(grad, (0, 0), mask)
    m = mark_white.resize((int(badge.size[1] * 0.78), int(badge.size[1] * 0.78)), Image.LANCZOS)
    badge.alpha_composite(m, ((badge.size[0] - m.size[0]) // 2, (badge.size[1] - m.size[1]) // 2))
    bsh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(bsh).rounded_rectangle((bx0, by0 + 12, bx1, by1 + 12), radius=int(S * 0.07), fill=(0, 0, 0, 60))
    img = Image.alpha_composite(img, bsh.filter(ImageFilter.GaussianBlur(S * 0.015)))
    img.alpha_composite(badge, (bx0, by0))
    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    os.makedirs(os.path.join(BUILD, "icons"), exist_ok=True)
    os.makedirs(ASSETS, exist_ok=True)
    os.makedirs(PUBLIC, exist_ok=True)

    icon1024 = app_icon(1024)
    icon1024.save(os.path.join(BUILD, "icon.png"))
    for s in (16, 24, 32, 48, 64, 128, 256, 512, 1024):
        app_icon(s).save(os.path.join(BUILD, "icons", f"{s}x{s}.png"))
    ico_sizes = [(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (128, 128), (256, 256)]
    app_icon(256).save(os.path.join(BUILD, "icon.ico"), sizes=ico_sizes)

    app_icon(256).save(os.path.join(ASSETS, "app-icon.png"))
    app_icon(64).save(os.path.join(PUBLIC, "favicon.png"))

    mark = transparent_mark()
    mark.resize((512, 512), Image.LANCZOS).save(os.path.join(ASSETS, "logo-mark.png"), optimize=True)

    # white version of the mark for dark backgrounds / installer
    white = mark.copy()
    px = white.load()
    for y in range(white.size[1]):
        for x in range(white.size[0]):
            r, g, b, a = px[x, y]
            if a:
                # map the blue gradient to a white->ice gradient (like the app icon)
                t = min(1.0, max(0.0, (g - 79) / 55))
                px[x, y] = (round(255 - 70 * t), round(255 - 50 * t), 255, a)
    white.resize((512, 512), Image.LANCZOS).save(os.path.join(ASSETS, "logo-mark-white.png"), optimize=True)

    os.makedirs(os.path.join(BUILD, "fileicons"), exist_ok=True)
    for kind in ("doc", "sheet", "slides"):
        fi = file_icon(kind, white, 256)
        fi.save(os.path.join(BUILD, "fileicons", f"{kind}.png"))
        fi.save(os.path.join(BUILD, "fileicons", f"{kind}.ico"), sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        fi.resize((128, 128), Image.LANCZOS).save(os.path.join(ASSETS, f"file-{kind}.png"))

    sidebar = installer_sidebar(white)
    sidebar.save(os.path.join(BUILD, "installerSidebar.bmp"))
    sidebar.save(os.path.join(BUILD, "uninstallerSidebar.bmp"))
    installer_header(mark).save(os.path.join(BUILD, "installerHeader.bmp"))
    print("icons generated")


if __name__ == "__main__":
    main()
