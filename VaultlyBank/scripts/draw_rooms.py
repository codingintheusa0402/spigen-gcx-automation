#!/usr/bin/env python3
"""미니룸 — Vaultly's original pixel-art rooms (Cyworld-style), one per wealth tier, plus furniture that appears as
net worth grows toward the next tier and "poor" clutter when the account is empty or in debt.

    ~/.venvs/pixel/bin/python scripts/draw_rooms.py

Writes Vaultly/Rooms/room_<tier>.png (200×150 background), Vaultly/Rooms/rp_<prop>.png and Vaultly/Rooms/rooms.json
(placement, unlock progress 0…1 toward the next tier, "when": always | poor). Every pixel is drawn by this file.
"""
import json, os
from PIL import Image, ImageDraw
import numpy as np

OUT = os.path.join(os.path.dirname(__file__), "..", "Vaultly", "Rooms")
W, H = 200, 150
WALL_BOTTOM = 84
OUTLINE = (30, 24, 34, 255)

def tones(rgb, light=1.18, dark=0.72):
    r, g, b = rgb
    f = lambda x, k: max(0, min(255, int(x * k)))
    return [(f(r, light), f(g, light), f(b, light), 255), (r, g, b, 255), (f(r, dark), f(g, dark), f(b, dark), 255)]

def new(w, h):
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0)); return im, ImageDraw.Draw(im)

def finish(layer, tone, outline=True):
    """3-tone shading (light top-left, dark right/bottom) + 1px outline by dilation."""
    a = np.array(layer); h, w = a.shape[:2]
    m = a[:, :, 3] > 0
    base = np.all(a[:, :, :3] == np.array(tone[1][:3]), axis=2) & m
    def sh(arr, dy, dx):
        out = np.zeros_like(arr)
        ys, ye = max(0, dy), h + min(0, dy); xs, xe = max(0, dx), w + min(0, dx)
        out[ys:ye, xs:xe] = arr[ys - dy:ye - dy, xs - dx:xe - dx]
        return out
    dark = base & (~sh(m, 0, -1) | ~sh(m, -1, 0))
    light = base & ~dark & (~sh(m, 0, 1) | ~sh(m, 1, 0))
    a[dark] = tone[2]; a[light] = tone[0]
    if outline:
        near = sh(m, 0, 1) | sh(m, 0, -1) | sh(m, 1, 0) | sh(m, -1, 0)
        a[near & ~m] = OUTLINE
    return Image.fromarray(a, "RGBA")

def part(w, h, rgb, fn, outline=True):
    """One shaded piece: fn(draw, base_color) draws flat shapes in the base tone."""
    t = tones(rgb); im, d = new(w, h); fn(d, t[1]); return finish(im, t, outline)

def compose(w, h, *layers):
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    for layer, (x, y) in layers:
        im.alpha_composite(layer, (x, y))
    return im

# ── room backgrounds ───────────────────────────────────────────────────────────────────────
def room_bg(wall, floor, floor_style="wood", wall_style="plain", side=(0, 0, 0)):
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    wt, ft = tones(wall), tones(floor)
    d.rectangle((0, 0, W, WALL_BOTTOM), fill=wt[1])
    if wall_style == "stripes":
        for x in range(0, W, 12): d.rectangle((x, 0, x + 5, WALL_BOTTOM), fill=wt[0])
    elif wall_style == "dots":
        for y in range(6, WALL_BOTTOM, 12):
            for x in range((y // 12 % 2) * 6, W, 12): d.point((x, y), fill=wt[2])
    elif wall_style == "panel":
        for x in range(0, W, 40): d.line((x, 0, x, WALL_BOTTOM), fill=wt[2]); d.line((x + 1, 0, x + 1, WALL_BOTTOM), fill=wt[0])
        d.rectangle((0, 40, W, 43), fill=wt[2])
    elif wall_style == "stain":
        for (x, y, r) in [(150, 20, 9), (40, 50, 6), (120, 62, 5)]: d.ellipse((x - r, y - r, x + r, y + r), fill=wt[2])
        d.line((60, 0, 64, 30), fill=wt[2]); d.line((64, 30, 61, 48), fill=wt[2])
    d.rectangle((0, WALL_BOTTOM - 3, W, WALL_BOTTOM), fill=wt[2])          # baseboard
    d.rectangle((0, WALL_BOTTOM, W, H), fill=ft[1])
    if floor_style == "wood":
        for y in range(WALL_BOTTOM, H, 9):
            d.line((0, y, W, y), fill=ft[2])
            off = (y // 9 % 2) * 30
            for x in range(off, W, 60): d.line((x, y, x, y + 9), fill=ft[2])
    elif floor_style == "lino":
        for y in range(WALL_BOTTOM, H, 11): d.line((0, y, W, y), fill=ft[2])
        for x in range(0, W, 11): d.line((x, WALL_BOTTOM, x, H), fill=ft[2])
        for (x, y) in [(30, 120), (140, 100), (90, 140)]: d.rectangle((x, y, x + 10, y + 10), fill=ft[2])
    elif floor_style == "tile":
        for y in range(WALL_BOTTOM, H, 16):
            for x in range(-8, W, 16):
                d.polygon([(x + 8, y), (x + 16, y + 8), (x + 8, y + 16), (x, y + 8)], fill=ft[0] if ((x // 16 + y // 16) % 2) else ft[1], outline=ft[2])
    elif floor_style == "marble":
        for (x0, y0, x1, y1) in [(10, 90, 60, 140), (120, 86, 170, 120), (80, 110, 190, 148), (0, 130, 40, 150)]:
            d.line((x0, y0, x1, y1), fill=ft[0]); d.line((x0 + 2, y0 + 1, x1 + 2, y1 + 1), fill=ft[2])
        for x in range(0, W, 50): d.line((x, WALL_BOTTOM, x, H), fill=ft[2])
        d.line((0, 117, W, 117), fill=ft[2])
    elif floor_style == "carpet":
        for y in range(WALL_BOTTOM + 2, H, 4):
            for x in range((y // 4 % 2) * 2, W, 4): d.point((x, y), fill=ft[2])
    # left side wall (2.5D corner) — darker, slanted bottom edge
    st = tones(side if side != (0, 0, 0) else tuple(int(c * 0.82) for c in wall))
    d.polygon([(0, 0), (34, 0), (34, WALL_BOTTOM), (0, WALL_BOTTOM + 26)], fill=st[1])
    d.line((34, 0, 34, WALL_BOTTOM), fill=st[2]); d.line((0, WALL_BOTTOM + 26, 34, WALL_BOTTOM), fill=st[2])
    d.rectangle((0, 0, W - 1, H - 1), outline=OUTLINE)
    return im

ROOM_BG = {
    "broke":   dict(wall=(196, 184, 150), floor=(150, 146, 136), floor_style="lino", wall_style="stain"),
    "working": dict(wall=(206, 226, 214), floor=(178, 140, 96), floor_style="wood", wall_style="plain"),
    "middle":  dict(wall=(238, 236, 228), floor=(206, 170, 120), floor_style="wood", wall_style="dots"),
    "rich":    dict(wall=(58, 62, 86), floor=(92, 76, 70), floor_style="carpet", wall_style="panel", side=(44, 46, 68)),
    "tycoon":  dict(wall=(236, 226, 206), floor=(224, 220, 212), floor_style="marble", wall_style="panel", side=(206, 190, 164)),
}

# ── props (each: name → (w, h, builder)) ─────────────────────────────────────────────────
def rect(d, c, box): d.rectangle(box, fill=c)
def props():
    P = {}
    def add(name, w, h, fn): P[name] = (w, h, fn)

    # 거지 — 고시원
    add("mattress", 54, 22, lambda: compose(54, 22,
        (part(54, 14, (120, 128, 150), lambda d, c: rect(d, c, (0, 6, 53, 13))), (0, 8)),
        (part(16, 8, (236, 236, 228), lambda d, c: rect(d, c, (0, 0, 15, 7))), (3, 4))))
    add("cupramen", 20, 26, lambda: compose(20, 26, *[(part(16, 10, (228, 60, 48) if i % 2 else (240, 150, 40),
        lambda d, c: d.polygon([(2, 0), (13, 0), (15, 9), (0, 9)], fill=c)), (2, 16 - i * 8)) for i in range(3)]))
    add("fan", 22, 34, lambda: compose(22, 34,
        (part(6, 14, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 5, 13))), (8, 18)),
        (part(14, 4, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 13, 3))), (4, 30)),
        (part(22, 22, (210, 214, 220), lambda d, c: d.ellipse((0, 0, 21, 21), fill=c)), (0, 0)),
        (part(8, 8, (90, 92, 100), lambda d, c: d.ellipse((0, 0, 7, 7), fill=c)), (7, 7))))
    add("bulb", 10, 30, lambda: compose(10, 30,
        (part(2, 20, (80, 80, 84), lambda d, c: rect(d, c, (0, 0, 1, 19)), outline=False), (4, 0)),
        (part(10, 12, (250, 236, 150), lambda d, c: d.ellipse((0, 0, 9, 11), fill=c)), (0, 18))))
    add("boxes", 30, 28, lambda: compose(30, 28,
        (part(26, 14, (190, 150, 100), lambda d, c: rect(d, c, (0, 0, 25, 13))), (2, 14)),
        (part(20, 12, (200, 160, 110), lambda d, c: rect(d, c, (0, 0, 19, 11))), (5, 3))))
    add("bills", 22, 26, lambda: compose(22, 26, *[(part(14, 18, (240, 240, 232), lambda d, c: rect(d, c, (0, 0, 13, 17))), (i * 4, i * 3)) for i in range(3)],
        (part(8, 2, (220, 60, 60), lambda d, c: rect(d, c, (0, 0, 7, 1)), outline=False), (11, 12))))
    add("soju", 24, 16, lambda: compose(24, 16, *[(part(6, 16, (90, 160, 110), lambda d, c: (rect(d, c, (1, 0, 4, 4)), rect(d, c, (0, 4, 5, 15)))), (i * 7, 0)) for i in range(3)]))
    add("window_small", 30, 26, lambda: compose(30, 26,
        (part(30, 26, (160, 150, 130), lambda d, c: rect(d, c, (0, 0, 29, 25))), (0, 0)),
        (part(24, 20, (150, 190, 230), lambda d, c: rect(d, c, (0, 0, 23, 19))), (3, 3)),
        (part(2, 20, (160, 150, 130), lambda d, c: rect(d, c, (0, 0, 1, 19)), outline=False), (14, 3))))
    add("books", 18, 12, lambda: compose(18, 12, *[(part(18, 4, col, lambda d, c: rect(d, c, (0, 0, 17, 3))), (0, 8 - i * 4)) for i, col in enumerate([(70, 110, 180), (200, 80, 70), (230, 190, 90)])]))
    add("minifridge", 22, 30, lambda: compose(22, 30,
        (part(22, 30, (230, 232, 236), lambda d, c: rect(d, c, (0, 0, 21, 29))), (0, 0)),
        (part(2, 8, (120, 120, 130), lambda d, c: rect(d, c, (0, 0, 1, 7)), outline=False), (17, 10))))
    add("laptop", 22, 14, lambda: compose(22, 14,
        (part(18, 12, (90, 92, 100), lambda d, c: rect(d, c, (0, 0, 17, 11))), (2, 0)),
        (part(14, 8, (120, 180, 230), lambda d, c: rect(d, c, (0, 0, 13, 7)), outline=False), (4, 2)),
        (part(22, 3, (120, 122, 130), lambda d, c: rect(d, c, (0, 0, 21, 2))), (0, 11))))

    # 서민 — 원룸
    add("bed", 60, 32, lambda: compose(60, 32,
        (part(60, 10, (150, 110, 70), lambda d, c: rect(d, c, (0, 0, 59, 9))), (0, 0)),
        (part(58, 20, (110, 150, 200), lambda d, c: rect(d, c, (0, 0, 57, 19))), (1, 10)),
        (part(18, 9, (240, 240, 236), lambda d, c: rect(d, c, (0, 0, 17, 8))), (4, 12))))
    add("desk_laptop", 44, 36, lambda: compose(44, 36,
        (part(44, 8, (160, 120, 80), lambda d, c: rect(d, c, (0, 0, 43, 7))), (0, 14)),
        (part(4, 16, (120, 90, 60), lambda d, c: rect(d, c, (0, 0, 3, 15))), (3, 20)),
        (part(4, 16, (120, 90, 60), lambda d, c: rect(d, c, (0, 0, 3, 15))), (37, 20)),
        (part(18, 12, (90, 92, 100), lambda d, c: rect(d, c, (0, 0, 17, 11))), (13, 2)),
        (part(14, 8, (120, 180, 230), lambda d, c: rect(d, c, (0, 0, 13, 7)), outline=False), (15, 4))))
    add("hanger", 26, 44, lambda: compose(26, 44,
        (part(26, 3, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 25, 2))), (0, 0)),
        (part(3, 44, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 2, 43))), (0, 0)),
        (part(3, 44, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 2, 43))), (23, 0)),
        (part(8, 20, (64, 112, 196), lambda d, c: rect(d, c, (0, 0, 7, 19))), (4, 4)),
        (part(8, 20, (228, 160, 180), lambda d, c: rect(d, c, (0, 0, 7, 19))), (13, 4))))
    add("microwave", 26, 16, lambda: compose(26, 16,
        (part(26, 16, (210, 212, 216), lambda d, c: rect(d, c, (0, 0, 25, 15))), (0, 0)),
        (part(14, 10, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 13, 9)), outline=False), (3, 3))))
    add("plant", 20, 30, lambda: compose(20, 30,
        (part(14, 12, (200, 120, 80), lambda d, c: d.polygon([(0, 0), (13, 0), (11, 11), (2, 11)], fill=c)), (3, 18)),
        (part(20, 20, (70, 160, 90), lambda d, c: (d.ellipse((0, 4, 11, 15), fill=c), d.ellipse((8, 0, 19, 11), fill=c), d.ellipse((5, 8, 16, 19), fill=c))), (0, 0))))
    add("tv_small", 30, 24, lambda: compose(30, 24,
        (part(30, 20, (50, 52, 60), lambda d, c: rect(d, c, (0, 0, 29, 19))), (0, 0)),
        (part(26, 16, (90, 150, 210), lambda d, c: rect(d, c, (0, 0, 25, 15)), outline=False), (2, 2)),
        (part(10, 4, (50, 52, 60), lambda d, c: rect(d, c, (0, 0, 9, 3))), (10, 20))))
    add("bookshelf", 34, 50, lambda: compose(34, 50,
        (part(34, 50, (150, 110, 70), lambda d, c: rect(d, c, (0, 0, 33, 49))), (0, 0)),
        *[(part(28, 10, (236, 220, 190), lambda d, c: rect(d, c, (0, 0, 27, 9)), outline=False), (3, 3 + i * 15)) for i in range(3)],
        *[(part(6, 9, col, lambda d, c: rect(d, c, (0, 0, 5, 8))), (4 + (i % 4) * 7, 4 + (i // 4) * 15)) for i, col in enumerate([(70, 110, 180), (200, 80, 70), (230, 190, 90), (80, 160, 110), (120, 80, 160), (220, 120, 60), (60, 60, 70), (160, 200, 230), (200, 100, 150), (90, 130, 90), (240, 160, 60), (100, 100, 180)])]))
    add("purifier", 16, 30, lambda: compose(16, 30,
        (part(16, 30, (240, 240, 244), lambda d, c: d.rounded_rectangle((0, 0, 15, 29), 4, fill=c)), (0, 0)),
        (part(6, 2, (80, 200, 240), lambda d, c: rect(d, c, (0, 0, 5, 1)), outline=False), (5, 4))))
    add("bike", 44, 28, lambda: compose(44, 28,
        (part(16, 16, (60, 60, 66), lambda d, c: d.ellipse((0, 0, 15, 15), fill=c)), (0, 12)),
        (part(16, 16, (60, 60, 66), lambda d, c: d.ellipse((0, 0, 15, 15), fill=c)), (28, 12)),
        (part(30, 14, (200, 60, 60), lambda d, c: (d.line((2, 12, 12, 0), fill=c, width=3), d.line((12, 0, 24, 0), fill=c, width=3), d.line((24, 0, 28, 12), fill=c, width=3), d.line((12, 0, 16, 12), fill=c, width=3))), (7, 6))))

    # 중산층 — 아파트
    add("sofa", 64, 30, lambda: compose(64, 30,
        (part(64, 22, (96, 110, 150), lambda d, c: rect(d, c, (0, 0, 63, 21))), (0, 8)),
        (part(56, 12, (120, 136, 176), lambda d, c: rect(d, c, (0, 0, 55, 11))), (4, 0)),
        (part(20, 10, (240, 200, 90), lambda d, c: rect(d, c, (0, 0, 19, 9))), (6, 12)),
        (part(20, 10, (240, 200, 90), lambda d, c: rect(d, c, (0, 0, 19, 9))), (38, 12))))
    add("tv_big", 56, 40, lambda: compose(56, 40,
        (part(56, 30, (40, 42, 50), lambda d, c: rect(d, c, (0, 0, 55, 29))), (0, 0)),
        (part(52, 26, (80, 140, 200), lambda d, c: rect(d, c, (0, 0, 51, 25)), outline=False), (2, 2)),
        (part(56, 8, (150, 110, 70), lambda d, c: rect(d, c, (0, 0, 55, 7))), (0, 32))))
    add("rug", 70, 22, lambda: part(70, 22, (200, 90, 90), lambda d, c: (d.ellipse((0, 0, 69, 21), fill=c))))
    add("dining", 50, 34, lambda: compose(50, 34,
        (part(50, 8, (190, 150, 100), lambda d, c: rect(d, c, (0, 0, 49, 7))), (0, 10)),
        (part(4, 16, (150, 110, 70), lambda d, c: rect(d, c, (0, 0, 3, 15))), (4, 18)),
        (part(4, 16, (150, 110, 70), lambda d, c: rect(d, c, (0, 0, 3, 15))), (42, 18)),
        (part(10, 6, (240, 240, 236), lambda d, c: d.ellipse((0, 0, 9, 5), fill=c)), (12, 6)),
        (part(10, 6, (240, 240, 236), lambda d, c: d.ellipse((0, 0, 9, 5), fill=c)), (28, 6))))
    add("coffee", 16, 20, lambda: compose(16, 20,
        (part(16, 20, (70, 70, 76), lambda d, c: rect(d, c, (0, 0, 15, 19))), (0, 0)),
        (part(6, 5, (240, 240, 236), lambda d, c: rect(d, c, (0, 0, 5, 4))), (5, 13))))
    add("aircon", 40, 12, lambda: part(40, 12, (240, 240, 244), lambda d, c: d.rounded_rectangle((0, 0, 39, 11), 3, fill=c)))
    add("frame", 26, 20, lambda: compose(26, 20,
        (part(26, 20, (170, 130, 70), lambda d, c: rect(d, c, (0, 0, 25, 19))), (0, 0)),
        (part(20, 14, (120, 170, 120), lambda d, c: rect(d, c, (0, 0, 19, 13)), outline=False), (3, 3)),
        (part(8, 6, (240, 220, 120), lambda d, c: d.ellipse((0, 0, 7, 5), fill=c), outline=False), (12, 5))))
    add("robot_vacuum", 18, 10, lambda: part(18, 10, (60, 62, 70), lambda d, c: d.ellipse((0, 0, 17, 9), fill=c)))
    add("aquarium", 40, 28, lambda: compose(40, 28,
        (part(40, 28, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 39, 27))), (0, 0)),
        (part(36, 22, (90, 170, 220), lambda d, c: rect(d, c, (0, 0, 35, 21)), outline=False), (2, 3)),
        (part(8, 5, (240, 140, 60), lambda d, c: d.ellipse((0, 0, 7, 4), fill=c), outline=False), (8, 10)),
        (part(8, 5, (240, 220, 80), lambda d, c: d.ellipse((0, 0, 7, 4), fill=c), outline=False), (22, 15))))

    # 부자 — 펜트하우스
    def city(d, c):
        rect(d, c, (0, 0, 79, 49))
        for i, (x, h) in enumerate([(4, 24), (14, 34), (26, 18), (36, 40), (50, 28), (62, 36), (72, 22)]):
            d.rectangle((x, 48 - h, x + 8, 48), fill=(34, 34, 56, 255))
            for yy in range(50 - h, 46, 5):
                for xx in range(x + 1, x + 8, 3): d.point((xx, yy), fill=(250, 230, 140, 255))
    add("window_city", 84, 54, lambda: compose(84, 54,
        (part(84, 54, (120, 120, 130), lambda d, c: rect(d, c, (0, 0, 83, 53))), (0, 0)),
        (part(80, 50, (24, 26, 48), city, outline=False), (2, 2))))
    add("sofa_leather", 70, 32, lambda: compose(70, 32,
        (part(70, 24, (70, 48, 40), lambda d, c: rect(d, c, (0, 0, 69, 23))), (0, 8)),
        (part(62, 12, (96, 66, 54), lambda d, c: rect(d, c, (0, 0, 61, 11))), (4, 0)),
        *[(part(18, 10, (120, 84, 68), lambda d, c: rect(d, c, (0, 0, 17, 9))), (6 + i * 20, 12)) for i in range(3)]))
    add("winebar", 36, 52, lambda: compose(36, 52,
        (part(36, 52, (60, 44, 40), lambda d, c: rect(d, c, (0, 0, 35, 51))), (0, 0)),
        *[(part(5, 14, (90, 30, 50), lambda d, c: (rect(d, c, (1, 0, 3, 3)), rect(d, c, (0, 3, 4, 13)))), (4 + (i % 5) * 6, 5 + (i // 5) * 16)) for i in range(15)]))
    add("painting", 40, 30, lambda: compose(40, 30,
        (part(40, 30, (200, 170, 80), lambda d, c: rect(d, c, (0, 0, 39, 29))), (0, 0)),
        (part(32, 22, (90, 60, 120), lambda d, c: rect(d, c, (0, 0, 31, 21)), outline=False), (4, 4)),
        (part(12, 12, (240, 120, 80), lambda d, c: d.ellipse((0, 0, 11, 11), fill=c), outline=False), (14, 8))))
    add("rug_rich", 80, 26, lambda: compose(80, 26,
        (part(80, 26, (120, 40, 50), lambda d, c: rect(d, c, (0, 0, 79, 25))), (0, 0)),
        (part(66, 14, (200, 160, 80), lambda d, c: rect(d, c, (0, 0, 65, 13)), outline=False), (7, 6))))
    add("piano", 60, 44, lambda: compose(60, 44,
        (part(60, 26, (30, 30, 36), lambda d, c: d.polygon([(0, 8), (40, 0), (59, 6), (59, 25), (0, 25)], fill=c)), (0, 0)),
        (part(44, 6, (240, 240, 236), lambda d, c: rect(d, c, (0, 0, 43, 5))), (2, 26)),
        *[(part(3, 3, (30, 30, 36), lambda d, c: rect(d, c, (0, 0, 2, 2)), outline=False), (5 + i * 6, 26)) for i in range(7)],
        (part(4, 14, (30, 30, 36), lambda d, c: rect(d, c, (0, 0, 3, 13))), (4, 30)),
        (part(4, 14, (30, 30, 36), lambda d, c: rect(d, c, (0, 0, 3, 13))), (52, 30))))
    add("theater", 70, 44, lambda: compose(70, 44,
        (part(70, 38, (30, 32, 40), lambda d, c: rect(d, c, (0, 0, 69, 37))), (0, 0)),
        (part(66, 34, (160, 90, 140), lambda d, c: rect(d, c, (0, 0, 65, 33)), outline=False), (2, 2)),
        (part(26, 10, (240, 200, 120), lambda d, c: d.ellipse((0, 0, 25, 9), fill=c), outline=False), (22, 12)),
        (part(70, 6, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 69, 5))), (0, 38))))
    add("sculpture", 20, 44, lambda: compose(20, 44,
        (part(20, 10, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 19, 9))), (0, 34)),
        (part(12, 36, (230, 230, 226), lambda d, c: (d.ellipse((2, 0, 9, 7), fill=c), d.polygon([(0, 8), (11, 8), (9, 35), (2, 35)], fill=c))), (4, 0))))
    add("golf", 18, 34, lambda: compose(18, 34,
        (part(12, 26, (40, 40, 46), lambda d, c: d.polygon([(0, 4), (11, 0), (11, 25), (0, 25)], fill=c)), (3, 8)),
        *[(part(3, 12, (200, 200, 206), lambda d, c: rect(d, c, (0, 0, 2, 11))), (5 + i * 3, 0)) for i in range(3)]))
    add("watch_case", 24, 16, lambda: compose(24, 16,
        (part(24, 16, (40, 40, 46), lambda d, c: rect(d, c, (0, 0, 23, 15))), (0, 0)),
        *[(part(5, 7, (230, 200, 90), lambda d, c: d.ellipse((0, 0, 4, 6), fill=c), outline=False), (3 + i * 7, 4)) for i in range(3)]))
    add("treadmill", 44, 30, lambda: compose(44, 30,
        (part(44, 8, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 43, 7))), (0, 22)),
        (part(4, 24, (120, 120, 126), lambda d, c: rect(d, c, (0, 0, 3, 23))), (38, 0)),
        (part(16, 6, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 15, 5))), (26, 0))))

    # 재벌 — 저택
    add("chandelier", 44, 36, lambda: compose(44, 36,
        (part(2, 10, (200, 170, 80), lambda d, c: rect(d, c, (0, 0, 1, 9)), outline=False), (21, 0)),
        (part(44, 10, (200, 170, 80), lambda d, c: d.polygon([(0, 9), (22, 0), (43, 9)], fill=c)), (0, 8)),
        *[(part(6, 10, (250, 240, 170), lambda d, c: d.polygon([(3, 0), (5, 9), (0, 9)], fill=c)), (1 + i * 9, 18)) for i in range(5)]))
    add("stairs", 70, 60, lambda: compose(70, 60, *[(part(70 - i * 10, 10, (210, 200, 180), lambda d, c, w=70 - i * 10: rect(d, c, (0, 0, w - 1, 9))), (i * 10, 50 - i * 10)) for i in range(5)],
        (part(4, 50, (200, 170, 80), lambda d, c: rect(d, c, (0, 0, 3, 49))), (66, 0))))
    add("fireplace", 56, 50, lambda: compose(56, 50,
        (part(56, 50, (120, 110, 100), lambda d, c: rect(d, c, (0, 0, 55, 49))), (0, 0)),
        (part(30, 28, (40, 36, 36), lambda d, c: rect(d, c, (0, 0, 29, 27)), outline=False), (13, 18)),
        (part(20, 16, (240, 120, 40), lambda d, c: d.polygon([(10, 0), (19, 15), (0, 15)], fill=c)), (18, 28)),
        (part(56, 6, (200, 190, 170), lambda d, c: rect(d, c, (0, 0, 55, 5))), (0, 10))))
    add("portrait", 36, 46, lambda: compose(36, 46,
        (part(36, 46, (200, 170, 80), lambda d, c: rect(d, c, (0, 0, 35, 45))), (0, 0)),
        (part(28, 38, (60, 70, 110), lambda d, c: rect(d, c, (0, 0, 27, 37)), outline=False), (4, 4)),
        (part(14, 14, (236, 196, 160), lambda d, c: d.ellipse((0, 0, 13, 13), fill=c)), (11, 10)),
        (part(20, 14, (40, 40, 50), lambda d, c: rect(d, c, (0, 0, 19, 13)), outline=False), (8, 26))))
    add("statue", 26, 60, lambda: compose(26, 60,
        (part(26, 12, (200, 196, 190), lambda d, c: rect(d, c, (0, 0, 25, 11))), (0, 48)),
        (part(16, 48, (236, 236, 232), lambda d, c: (d.ellipse((4, 0, 11, 8), fill=c), d.polygon([(2, 9), (13, 9), (15, 47), (0, 47)], fill=c))), (5, 0))))
    add("vault", 40, 44, lambda: compose(40, 44,
        (part(40, 44, (90, 94, 110), lambda d, c: rect(d, c, (0, 0, 39, 43))), (0, 0)),
        (part(18, 18, (140, 146, 160), lambda d, c: d.ellipse((0, 0, 17, 17), fill=c)), (11, 13)),
        (part(8, 8, (200, 170, 80), lambda d, c: d.ellipse((0, 0, 7, 7), fill=c), outline=False), (16, 18))))
    add("car_model", 40, 18, lambda: compose(40, 18,
        (part(40, 12, (220, 40, 40), lambda d, c: d.polygon([(0, 11), (6, 4), (14, 0), (28, 0), (36, 4), (39, 11)], fill=c)), (0, 0)),
        (part(10, 10, (30, 30, 36), lambda d, c: d.ellipse((0, 0, 9, 9), fill=c)), (6, 8)),
        (part(10, 10, (30, 30, 36), lambda d, c: d.ellipse((0, 0, 9, 9), fill=c)), (24, 8))))
    add("fountain", 44, 40, lambda: compose(44, 40,
        (part(44, 14, (200, 196, 190), lambda d, c: d.ellipse((0, 0, 43, 13), fill=c)), (0, 26)),
        (part(36, 10, (90, 170, 220), lambda d, c: d.ellipse((0, 0, 35, 9), fill=c), outline=False), (4, 27)),
        (part(6, 24, (200, 196, 190), lambda d, c: rect(d, c, (0, 0, 5, 23))), (19, 6)),
        (part(16, 8, (160, 210, 240), lambda d, c: d.ellipse((0, 0, 15, 7), fill=c), outline=False), (14, 0))))
    add("gold", 30, 18, lambda: compose(30, 18, *[(part(14, 6, (240, 200, 60), lambda d, c: d.polygon([(2, 0), (11, 0), (13, 5), (0, 5)], fill=c)), (x, y)) for (x, y) in [(0, 12), (14, 12), (7, 6), (7, 0)]]))
    add("heli", 44, 24, lambda: compose(44, 24,
        (part(44, 3, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 43, 2))), (0, 0)),
        (part(26, 14, (30, 30, 40), lambda d, c: d.rounded_rectangle((0, 0, 25, 13), 5, fill=c)), (4, 5)),
        (part(12, 6, (120, 180, 230), lambda d, c: d.ellipse((0, 0, 11, 5), fill=c), outline=False), (7, 8)),
        (part(14, 4, (30, 30, 40), lambda d, c: rect(d, c, (0, 0, 13, 3))), (28, 10)),
        (part(26, 3, (60, 62, 70), lambda d, c: rect(d, c, (0, 0, 25, 2))), (4, 21))))
    return P

# ── placement ─────────────────────────────────────────────────────────────────────────────
A, P = "always", "poor"
ROOMS = {
    "broke": dict(name="고시원 2평", anchor=[108, 118], scale=0.62, props=[
        ("window_small", 120, 20, 0, A), ("bulb", 96, 0, 0, A), ("mattress", 42, 98, 0, A), ("boxes", 160, 100, 0, A),
        ("cupramen", 150, 60, 0, A), ("fan", 60, 54, 0.25, A), ("books", 150, 120, 0.45, A), ("laptop", 56, 126, 0.65, A),
        ("minifridge", 170, 58, 0.85, A), ("bills", 48, 24, 0, P), ("soju", 128, 134, 0, P)]),
    "working": dict(name="원룸 7평", anchor=[104, 120], scale=0.66, props=[
        ("window_small", 150, 14, 0, A), ("bed", 40, 86, 0, A), ("desk_laptop", 120, 60, 0, A), ("hanger", 170, 44, 0, A),
        ("minifridge", 44, 56, 0.15, A), ("microwave", 44, 40, 0.3, A), ("plant", 176, 104, 0.4, A), ("tv_small", 96, 44, 0.5, A),
        ("bookshelf", 60, 36, 0.65, A), ("purifier", 150, 112, 0.8, A), ("bike", 130, 124, 0.92, A), ("bills", 48, 14, 0, P), ("soju", 112, 130, 0, P)]),
    "middle": dict(name="아파트 32평 거실", anchor=[100, 124], scale=0.7, props=[
        ("window_small", 60, 12, 0, A), ("aircon", 140, 8, 0, A), ("rug", 60, 112, 0, A), ("sofa", 64, 88, 0, A), ("tv_big", 20, 40, 0, A),
        ("plant", 176, 60, 0, A), ("dining", 136, 94, 0.2, A), ("coffee", 150, 70, 0.35, A), ("bookshelf", 150, 30, 0.5, A),
        ("frame", 100, 24, 0.6, A), ("robot_vacuum", 40, 134, 0.75, A), ("aquarium", 44, 48, 0.9, A), ("bills", 110, 40, 0, P), ("soju", 20, 120, 0, P)]),
    "rich": dict(name="한강뷰 펜트하우스", anchor=[96, 126], scale=0.72, props=[
        ("window_city", 56, 10, 0, A), ("rug_rich", 56, 118, 0, A), ("sofa_leather", 60, 86, 0, A), ("winebar", 160, 30, 0, A),
        ("painting", 10, 20, 0, A), ("piano", 136, 90, 0.25, A), ("theater", 0, 44, 0.45, A), ("sculpture", 176, 86, 0.6, A),
        ("golf", 150, 110, 0.7, A), ("watch_case", 170, 110, 0.8, A), ("treadmill", 10, 118, 0.9, A), ("bills", 150, 6, 0, P), ("soju", 20, 130, 0, P)]),
    "tycoon": dict(name="평창동 대저택 홀", anchor=[100, 128], scale=0.74, props=[
        ("chandelier", 78, 0, 0, A), ("stairs", 128, 30, 0, A), ("fireplace", 40, 36, 0, A), ("portrait", 4, 28, 0, A), ("rug_rich", 56, 120, 0, A),
        ("statue", 170, 86, 0.2, A), ("vault", 8, 96, 0.35, A), ("car_model", 136, 128, 0.5, A), ("fountain", 50, 92, 0.65, A),
        ("gold", 20, 134, 0.8, A), ("heli", 150, 4, 0.92, A), ("bills", 110, 70, 0, P), ("soju", 112, 130, 0, P)]),
}

def main():
    os.makedirs(OUT, exist_ok=True)
    for tier, cfg in ROOM_BG.items(): room_bg(**cfg).save(os.path.join(OUT, f"room_{tier}.png"))
    built = props()
    sizes = {}
    for name, (w, h, fn) in built.items():
        im = fn(); sizes[name] = im.size; im.save(os.path.join(OUT, f"rp_{name}.png"))
    manifest = {"canvas": [W, H], "wallBottom": WALL_BOTTOM, "tiers": {}}
    for tier, cfg in ROOMS.items():
        manifest["tiers"][tier] = {"bg": f"room_{tier}", "name": cfg["name"], "anchor": cfg["anchor"], "scale": cfg["scale"],
            "props": [{"name": n, "x": x, "y": y, "w": sizes[n][0], "h": sizes[n][1], "unlock": u, "when": w} for (n, x, y, u, w) in cfg["props"]]}
    json.dump(manifest, open(os.path.join(OUT, "rooms.json"), "w"), ensure_ascii=False, indent=1)
    # preview sheet for a quick look
    sheet = Image.new("RGBA", (W * 5, H), (0, 0, 0, 255))
    for i, tier in enumerate(ROOMS):
        bg = Image.open(os.path.join(OUT, f"room_{tier}.png")).convert("RGBA")
        for p in sorted(manifest["tiers"][tier]["props"], key=lambda p: p["y"] + p["h"]):
            if p["when"] == "always": bg.alpha_composite(Image.open(os.path.join(OUT, f"rp_{p['name']}.png")).convert("RGBA"), (p["x"], p["y"]))
        sheet.alpha_composite(bg, (i * W, 0))
    sheet.resize((W * 5 * 3, H * 3), Image.NEAREST).save("/tmp/rooms_preview.png")
    print(f"rooms: {len(ROOMS)}, props: {len(built)} → {OUT}; preview /tmp/rooms_preview.png")

if __name__ == "__main__":
    main()
