#!/usr/bin/env python3
"""Vaultly's ORIGINAL pixel-art characters — drawn from scratch by code (no LPC pixels traced, copied or edited).

Same frame format as LPC (64×64 cells; rows = up, left, down, right; walk 9 frames with frame 0 standing) so the
characters fit the same renderer, but every pixel comes from the shapes below: a posed skeleton → body parts →
clothing layers, each with a dark outline and 3-tone shading. Vaultly owns all of it.

    python3 scripts/draw_original_characters.py [--examples DIR]

Writes Vaultly/Characters/co_<item>_<body>_<color>_<anim>.png, co_custom_<piece>_<body>_<anim>.png and
characters_original.json (same schema as characters.json, "prefix": "co"). Props (ch_prop_*) are shared — they were
already original Vaultly art.
"""
import argparse, json, math, os, random
from PIL import Image, ImageDraw
import numpy as np

OUT = os.path.join(os.path.dirname(__file__), "..", "Vaultly", "Characters")
S = 64
ANIMS = {"idle": (2, 4), "walk": (9, 4), "sit": (3, 4), "eat": (4, 4), "work": (4, 4), "hurt": (6, 1)}
OUTLINE = (30, 24, 34, 255)

# ── palettes (3 tones: light, base, dark) ───────────────────────────────────────────────
def tones(rgb, light=1.18, dark=0.72):
    r, g, b = rgb
    f = lambda x, k: max(0, min(255, int(x * k)))
    return [(f(r, light), f(g, light), f(b, light), 255), (r, g, b, 255), (f(r, dark), f(g, dark), f(b, dark), 255)]

SKINS = {"light": (246, 212, 182), "fair": (236, 196, 160), "tan": (212, 160, 118), "olive": (196, 150, 104),
         "brown": (160, 108, 72), "deep": (112, 72, 48)}
HAIR = {"black": (36, 32, 38), "dark_brown": (74, 50, 36), "brown": (112, 76, 48), "gray": (150, 150, 156),
        "white": (220, 220, 224), "dyed_brown": (150, 96, 60), "ash": (110, 104, 100)}
CLOTH = {"white": (238, 238, 232), "gray": (140, 140, 146), "charcoal": (70, 72, 80), "black": (40, 40, 46),
         "navy": (44, 58, 104), "blue": (64, 112, 196), "sky": (150, 196, 232), "green": (60, 150, 90), "forest": (44, 96, 64),
         "maroon": (128, 40, 52), "red": (196, 52, 52), "tan": (196, 168, 120), "brown": (112, 78, 52), "beige": (214, 198, 166),
         "pink": (228, 160, 180), "lavender": (176, 160, 214), "olive": (120, 120, 70), "denim": (70, 98, 150), "orange": (232, 120, 40)}

# ── skeleton ─────────────────────────────────────────────────────────────────────────────
def pose(direction, anim, f, body):
    """Return part geometry for one frame. direction: 0 up(back) 1 left 2 down(front) 3 right.
    Side views are drawn facing right and mirrored for left."""
    side = direction in (1, 3)
    w = 0 if body == "male" else -1                    # female: 1px narrower shoulders each side
    bob = 0
    p = {"side": side, "back": direction == 0, "front": direction == 2, "sit": anim in ("sit", "eat", "work")}
    if anim == "idle":
        bob = f % 2
    if anim == "walk" and f > 0:
        bob = 1 if f in (2, 3, 6, 7) else 0
    top = 10 + bob + (4 if p["sit"] else 0)          # sitting: whole upper body 4px lower
    p["head"] = (24, top, 39, top + 16) if not side else (25, top, 39, top + 16)
    p["neck"] = (30, top + 16, 33, top + 18)
    ty0, ty1 = top + 18, top + 31
    p["torso"] = (25 - w, ty0, 38 + w, ty1) if not side else (28, ty0, 36, ty1)
    p["hip_y"] = ty1
    # arms: list of (shoulder, hand) points
    ph = (f - 1) / 8 * 2 * math.pi if anim == "walk" and f > 0 else 0
    swing = math.sin(ph)
    if not side:
        sl, sr = (23 - w, ty0 + 2), (40 + w, ty0 + 2)
        hl, hr = (22 - w, ty1 + int(round(swing * 1.5))), (41 + w, ty1 - int(round(swing * 1.5)))
        if anim == "eat":
            hr = (36, p["head"][3] - 1 - (f % 2) * 2)            # spoon to mouth
        if anim == "work":
            hl, hr = (27, ty1 - 3 + (f % 2)), (37, ty1 - 3 + ((f + 1) % 2))   # typing
        p["arms"] = [(sl, hl), (sr, hr)]
    else:
        sh = (32, ty0 + 2)
        hand = (32 + int(round(swing * 6)), ty1 - 1)
        if anim == "eat":
            hand = (37, p["head"][3] - 1 - (f % 2) * 2)
        if anim == "work":
            hand = (40, ty1 - 3 + (f % 2))
        if p["sit"] and anim == "sit":
            hand = (38, ty1 - 1)
        p["arms"] = [(sh, hand)]
    # legs: (hip, knee, foot) — feet rest on y 60
    if not side:
        lx, rx = 29, 34
        if p["sit"]:   # thighs toward the viewer: short legs, feet below the knees
            p["legs"] = [((lx, ty1), (lx, ty1 + 5), (lx, 60)), ((rx, ty1), (rx, ty1 + 5), (rx, 60))]
        else:
            lift_l = max(0, swing) * 2.5; lift_r = max(0, -swing) * 2.5
            p["legs"] = [((lx, ty1), (lx, ty1 + 8), (lx, 60 - int(round(lift_l)))), ((rx, ty1), (rx, ty1 + 8), (rx, 60 - int(round(lift_r))))]
    else:
        hip = (32, ty1)
        if p["sit"]:
            p["legs"] = [(hip, (41, ty1 + 1), (41, 60)), (hip, (40, ty1 + 1), (40, 60))]
        else:
            legs = []
            for k in (1, -1):
                a = math.radians(26) * swing * k
                L = 60 - hip[1]
                knee = (hip[0] + int(round(math.sin(a) * L * 0.5)), hip[1] + int(round(math.cos(a) * L * 0.5)))
                foot = (hip[0] + int(round(math.sin(a) * L)), min(60, hip[1] + int(round(math.cos(a) * L))))
                legs.append((hip, knee, foot))
            p["legs"] = legs
    return p

# ── drawing helpers ──────────────────────────────────────────────────────────────────────
def thick(d, a, b, w, col):
    d.line([a, b], fill=col, width=w)
    r = w // 2
    for (x, y) in (a, b):
        d.ellipse((x - r, y - r, x + r - (1 - w % 2), y + r - (1 - w % 2)), fill=col)

def finish(layer, tone):
    """Shade (light on top-left edges, dark on right/bottom edges) and outline by dilation — vectorised."""
    a = np.array(layer)                                   # H×W×4
    m = a[:, :, 3] > 0
    base = np.all(a[:, :, :3] == np.array(tone[1][:3]), axis=2) & m
    def sh(arr, dy, dx):
        out = np.zeros_like(arr)
        ys, ye = max(0, dy), S + min(0, dy); xs, xe = max(0, dx), S + min(0, dx)
        out[ys:ye, xs:xe] = arr[ys - dy:ye - dy, xs - dx:xe - dx]
        return out
    right_empty = ~sh(m, 0, -1); below_empty = ~sh(m, -1, 0)
    left_empty = ~sh(m, 0, 1); above_empty = ~sh(m, 1, 0)
    dark = base & (right_empty | below_empty)
    light = base & ~dark & (left_empty | above_empty)
    a[dark] = tone[2]; a[light] = tone[0]
    near = sh(m, 0, 1) | sh(m, 0, -1) | sh(m, 1, 0) | sh(m, -1, 0)
    a[near & ~m] = OUTLINE
    return Image.fromarray(a, "RGBA").copy()

def torso_poly(p, grow=0, extra_bottom=0):
    x0, y0, x1, y1 = p["torso"]
    if p["side"]:
        return [(x0 - grow, y0 + 1), (x0 + 1, y0 - grow), (x1 - 1, y0 - grow), (x1 + grow, y0 + 1), (x1 + grow - 1, y1 + extra_bottom + grow), (x0 - grow + 1, y1 + extra_bottom + grow)]
    return [(x0 - grow, y0 + 2), (x0 + 2, y0 - grow), (x1 - 2, y0 - grow), (x1 + grow, y0 + 2), (x1 - 1 + grow, y1 + extra_bottom + grow), (x0 + 1 - grow, y1 + extra_bottom + grow)]

def blank():
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); return im, ImageDraw.Draw(im)

# ── body parts & clothes (each returns a finished 64×64 layer for one pose) ─────────────────
def draw_skin(p, skin):
    t = tones(SKINS[skin]); im, d = blank(); c = t[1]
    for leg in p["legs"]:
        thick(d, leg[0], leg[1], 4, c); thick(d, leg[1], leg[2], 4, c)
    d.polygon(torso_poly(p), fill=c)
    d.rectangle(p["neck"], fill=c)
    for sh, hand in p["arms"]:
        thick(d, sh, hand, 3, c)
    hx0, hy0, hx1, hy1 = p["head"]
    d.ellipse((hx0, hy0, hx1, hy1), fill=c)
    if p["side"]:
        d.rectangle((hx1, hy0 + 9, hx1 + 1, hy0 + 10), fill=c)        # nose
    im = finish(im, t)
    px = im.load(); dk = (40, 30, 34, 255)
    if p["front"]:
        cy = hy0 + 8
        for ex in (hx0 + 4, hx1 - 5):
            for dx in (0, 1):
                px[ex + dx, cy] = dk; px[ex + dx, cy + 1] = dk
            px[ex, cy] = (250, 250, 250, 255)
        for mx in (hx0 + 7, hx0 + 8): px[mx, hy0 + 12] = t[2]
        px[hx0 + 3, hy0 + 11] = tones(SKINS[skin], 1.0, 0.9)[2]; px[hx1 - 3, hy0 + 11] = tones(SKINS[skin], 1.0, 0.9)[2]
    elif p["side"]:
        px[hx1 - 4, hy0 + 8] = dk; px[hx1 - 4, hy0 + 9] = dk; px[hx1 - 3, hy0 + 8] = dk
        px[hx1 - 2, hy0 + 12] = t[2]
        px[hx0 + 5, hy0 + 8] = t[2]; px[hx0 + 5, hy0 + 9] = t[2]   # ear
    return im

def draw_hair(p, style, color):
    t = tones(HAIR[color], 1.25, 0.7); c = t[1]; im, d = blank()
    x0, y0, x1, y1 = p["head"]
    if style == "bald":
        return im
    if style == "balding":
        d.rectangle((x0, y0 + 5, x0 + 2, y0 + 9), fill=c); d.rectangle((x1 - 2, y0 + 5, x1, y0 + 9), fill=c)
        if p["back"]: d.rectangle((x0, y0 + 6, x1, y0 + 10), fill=c)
        return finish(im, t)
    if style == "buzz":
        d.chord((x0, y0 - 1, x1, y1 - 3), 180, 360, fill=c)
        if p["back"]: d.ellipse((x0, y0, x1, y1 - 3), fill=c)
        return finish(im, t)
    # cap of hair
    d.chord((x0 - 1, y0 - 2, x1 + 1, y1 - 2), 180, 360, fill=c)
    d.rectangle((x0 - 1, y0 + 4, x0 + 1, y0 + 9), fill=c)
    if not p["side"]:
        d.rectangle((x1 - 1, y0 + 4, x1 + 1, y0 + 9), fill=c)
    if p["back"]:
        d.ellipse((x0 - 1, y0 - 1, x1 + 1, y1 - 1), fill=c)
    if p["side"]:
        d.rectangle((x0 - 1, y0 + 2, x0 + 5, y0 + 10), fill=c)     # back of the head
    if style == "side_part" and p["front"]:
        d.polygon([(x0, y0 + 2), (x1 + 1, y0 + 1), (x1 + 1, y0 + 5), (x0 + 4, y0 + 5)], fill=c)
    if style == "fringe" and (p["front"] or p["side"]):
        d.rectangle((x0, y0 + 2, x1, y0 + 5), fill=c)
    if style == "messy":
        for k in range(x0, x1 + 1, 3):
            d.polygon([(k, y0 + 1), (k + 2, y0 - 3), (k + 3, y0 + 1)], fill=c)
    if style in ("long", "bob", "ponytail"):
        bottom = {"long": y1 + 6, "bob": y1 + 1, "ponytail": y0 + 9}[style]
        if p["back"]:
            d.rectangle((x0 - 1, y0 + 4, x1 + 1, bottom), fill=c)
        elif p["side"]:
            d.rectangle((x0 - 1, y0 + 3, x0 + 4, bottom), fill=c)
        else:
            d.rectangle((x0 - 2, y0 + 4, x0, bottom), fill=c); d.rectangle((x1, y0 + 4, x1 + 2, bottom), fill=c)
        if style == "ponytail" and not p["front"]:
            tx = x0 - 2 if p["side"] else (x0 + x1) // 2 - 1
            d.rectangle((tx, y0 + 5, tx + 2, y1 + 6), fill=c)
    return finish(im, t)

def draw_top(p, kind, color):
    """kind: tshirt | vneck | polo | shirt | cardigan | longsleeve"""
    t = tones(CLOTH[color]); c = t[1]; im, d = blank()
    x0, y0, x1, y1 = p["torso"]
    d.polygon(torso_poly(p, grow=1), fill=c)
    long = kind in ("shirt", "cardigan", "longsleeve")
    for sh, hand in p["arms"]:
        end = hand if long else (sh[0] + (hand[0] - sh[0]) * 0.45, sh[1] + (hand[1] - sh[1]) * 0.45)
        if long:   # stop the sleeve 2px before the hand so the hand shows
            L = max(1, math.dist(sh, hand)); end = (hand[0] - (hand[0] - sh[0]) / L * 2, hand[1] - (hand[1] - sh[1]) / L * 2)
        thick(d, sh, (int(end[0]), int(end[1])), 4, c)
    im = finish(im, t)
    px = im.load(); cx = (x0 + x1) // 2
    if p["front"]:
        if kind == "vneck":
            for k in range(4): px[cx - 2 + k // 2, y0 + k] = tones(CLOTH[color])[2]; px[cx + 1 - k // 2, y0 + k] = tones(CLOTH[color])[2]
        if kind in ("polo", "shirt"):
            wc = (250, 250, 248, 255) if color != "white" else (210, 214, 222, 255)
            for k in range(3): px[cx - 3 + k, y0 + k // 2] = wc; px[cx + 2 - k, y0 + k // 2] = wc
            for yy in range(y0 + 3, y1, 3): px[cx, yy] = t[2]
        if kind == "cardigan":
            for yy in range(y0 + 1, y1 + 1): px[cx, yy] = OUTLINE
            for yy in range(y0 + 3, y1, 3): px[cx - 2, yy] = (220, 210, 190, 255)
    return im

def draw_jacket(p, kind, color):
    """kind: suit | coat (doctor, long) | vest (store) | apron (cafe)"""
    t = tones(CLOTH[color], 1.12, 0.72); c = t[1]; im, d = blank()
    x0, y0, x1, y1 = p["torso"]
    if kind == "apron":
        d.rectangle((x0 + 2, y0 + 3, x1 - 2, min(60, y1 + (4 if p["sit"] else 10))), fill=c)
        if p["back"] or p["side"]:
            im2, d2 = blank(); d2.rectangle((x0, y1 - 4, x1, y1 - 3), fill=c); return finish(im2, t)
        im = finish(im, t); px = im.load()
        for xx in range((x0 + x1) // 2 - 3, (x0 + x1) // 2 + 3): px[xx, y1 - 1] = t[2]   # pocket
        return im
    long_ = kind == "coat"
    bottom = y1 + (2 if p["sit"] else (12 if long_ else 3))
    d.polygon(torso_poly(p, grow=1, extra_bottom=bottom - y1), fill=c)
    if kind != "vest":
        for sh, hand in p["arms"]:
            L = max(1, math.dist(sh, hand)); end = (hand[0] - (hand[0] - sh[0]) / L * 2, hand[1] - (hand[1] - sh[1]) / L * 2)
            thick(d, sh, (int(end[0]), int(end[1])), 5, c)
    im = finish(im, t)
    px = im.load(); cx = (x0 + x1) // 2
    if p["front"]:
        if kind in ("suit", "coat", "vest"):   # open front showing the shirt + lapels
            for k in range(0, 9):
                for xx in range(cx - 2 + k // 4, cx + 2 - k // 4):
                    if kind != "coat" or k < 5:
                        px[xx, y0 + k] = (0, 0, 0, 0)
            for k in range(7):
                px[cx - 3 + k // 3, y0 + k] = t[2]; px[cx + 2 - k // 3, y0 + k] = t[2]
        if kind == "coat":
            for yy in range(y0 + 8, bottom - 2, 4): px[cx + 2, yy] = (170, 175, 185, 255)
            for xx in range(x0 + 2, x0 + 6): px[xx, y0 + 10] = t[2]   # breast pocket
        if kind == "vest":
            for xx in range(x0 + 1, x1): px[xx, y0 + 9] = (245, 245, 240, 255)   # contrast stripe
    return im

def draw_bottom(p, kind, color):
    """kind: pants | shorts | skirt | slacks"""
    t = tones(CLOTH[color]); c = t[1]; im, d = blank()
    hip_y = p["hip_y"]
    if kind == "skirt" and not p["sit"]:
        x0, _, x1, _ = p["torso"]
        d.polygon([(x0 + 1, hip_y - 1), (x1 - 1, hip_y - 1), (x1 + 2, hip_y + 11), (x0 - 2, hip_y + 11)], fill=c)
        return finish(im, t)
    x0, _, x1, _ = p["torso"]
    d.rectangle((x0 + (0 if not p["side"] else 1), hip_y - 2, x1 - (0 if not p["side"] else 1), hip_y + 2), fill=c)
    for hip, knee, foot in p["legs"]:
        thick(d, hip, knee, 5, c)
        if kind in ("pants", "slacks"):
            end = (foot[0], foot[1] - 3)
            thick(d, knee, end, 5, c)
        elif kind == "skirt":
            pass
    im = finish(im, t)
    if kind == "slacks" and p["front"]:
        px = im.load()
        for hip, knee, foot in p["legs"]:
            for yy in range(hip[1] + 3, foot[1] - 4): px[hip[0], yy] = t[2]   # crease
    return im

def draw_shoes(p, kind, color):
    t = tones(CLOTH[color], 1.35, 0.6); c = t[1]; im, d = blank()
    for hip, knee, foot in p["legs"]:
        fx, fy = foot
        if p["side"]:
            d.rectangle((fx - 2, fy - 2, fx + 4, fy), fill=c)
        else:
            d.rectangle((fx - 3, fy - 2, fx + 2, fy), fill=c)
    im = finish(im, t)
    if kind == "sneakers":
        px = im.load()
        for hip, knee, foot in p["legs"]:
            for xx in range(foot[0] - 3, foot[0] + 4):
                if 0 <= xx < S and px[xx, foot[1]][3]: px[xx, foot[1]] = (250, 250, 250, 255)
    return im

def draw_accessory(p, kind, color):
    im, d = blank(); x0, y0, x1, y1 = p["head"]; px = im.load()
    if kind in ("glasses", "sunglasses") and not p["back"]:
        frame = (30, 30, 36, 255); lens = (30, 30, 40, 255) if kind == "sunglasses" else (200, 225, 245, 120)
        eyes = [(x0 + 2, y0 + 6), (x1 - 5, y0 + 6)] if p["front"] else [(x1 - 4, y0 + 6)]
        for ex, ey in eyes:
            d.rectangle((ex, ey, ex + 3, ey + 2), fill=lens, outline=frame)
        if p["front"]:
            d.line((x0 + 6, y0 + 7, x1 - 5, y0 + 7), fill=frame)
    if kind in ("tie", "bowtie") and p["front"]:
        cx = (p["torso"][0] + p["torso"][2]) // 2; ty = p["torso"][1]
        col = tones(CLOTH[color])
        if kind == "tie":
            d.polygon([(cx - 1, ty + 1), (cx + 1, ty + 1), (cx + 2, ty + 10), (cx, ty + 12), (cx - 2, ty + 10)], fill=col[1], outline=OUTLINE)
        else:
            d.polygon([(cx - 3, ty), (cx, ty + 1), (cx + 3, ty), (cx + 3, ty + 3), (cx, ty + 2), (cx - 3, ty + 3)], fill=col[1], outline=OUTLINE)
    return im

def draw_custom(p, piece):
    im, d = blank(); x0, y0, x1, y1 = p["head"]; cx = (x0 + x1) // 2
    if piece == "helmet":
        shell = tones((232, 112, 36))
        d.ellipse((x0 - 1, y0 - 2, x1 + 1, y1 - 2), fill=shell[1])
        d.rectangle((x0 - 1, y0 + 6, x1 + 1, y0 + 8), fill=shell[1])
        im = finish(im, shell); d = ImageDraw.Draw(im); px = im.load()
        for yy in range(y0 - 1, y0 + 6):
            if px[cx, yy][3]: px[cx, yy] = (245, 245, 240, 255)
        if p["front"]:
            d.rectangle((x0 + 1, y0 + 7, x1 - 1, y0 + 9), fill=(40, 50, 72, 255))
        elif p["side"]:
            d.rectangle((x1 - 4, y0 + 6, x1, y0 + 9), fill=(40, 50, 72, 255))
    if piece == "stethoscope" and p["front"]:
        ty = p["torso"][1]; tube = (60, 64, 80, 255)
        d.arc((cx - 5, ty - 3, cx + 5, ty + 6), 0, 180, fill=tube, width=1)
        d.line((cx + 4, ty + 3, cx + 4, ty + 9), fill=tube)
        d.ellipse((cx + 3, ty + 9, cx + 6, ty + 12), fill=(196, 200, 210, 255), outline=tube)
    if piece == "badge" and p["front"]:
        ty = p["torso"][1]
        d.rectangle((cx + 3, ty + 4, cx + 7, ty + 6), fill=(250, 250, 245, 255), outline=OUTLINE)
    return im

# ── wardrobe & tiers ────────────────────────────────────────────────────────────────────
WARDROBE = {   # key: (slot, drawer, kind, colours)
    "skin":        ("body", None, None, list(SKINS)),
    **{f"hair_{s}": ("hair", draw_hair, s, list(HAIR)) for s in ["short", "side_part", "fringe", "messy", "buzz", "long", "bob", "ponytail", "balding"]},
    "tshirt":      ("torso", draw_top, "tshirt", ["gray", "white", "navy", "charcoal", "olive", "maroon", "sky", "black"]),
    "vneck":       ("torso", draw_top, "vneck", ["white", "gray", "black", "pink"]),
    "polo":        ("torso", draw_top, "polo", ["navy", "white", "forest", "sky", "maroon"]),
    "shirt":       ("torso", draw_top, "shirt", ["white", "sky", "lavender", "pink", "gray"]),
    "cardigan":    ("torso", draw_top, "cardigan", ["beige", "gray", "navy", "forest", "maroon"]),
    "suit":        ("jacket", draw_jacket, "suit", ["black", "charcoal", "navy", "gray"]),
    "coat":        ("jacket", draw_jacket, "coat", ["white"]),
    "vest":        ("jacket", draw_jacket, "vest", ["green", "blue"]),
    "apron":       ("jacket", draw_jacket, "apron", ["brown", "black", "green"]),
    "pants":       ("legs", draw_bottom, "pants", ["gray", "denim", "charcoal", "navy", "brown", "beige", "black"]),
    "slacks":      ("legs", draw_bottom, "slacks", ["black", "charcoal", "navy", "gray"]),
    "shorts":      ("legs", draw_bottom, "shorts", ["gray", "denim", "navy"]),
    "skirt":       ("legs", draw_bottom, "skirt", ["black", "navy", "gray", "beige"]),
    "slippers":    ("feet", draw_shoes, "slippers", ["blue", "gray"]),
    "sneakers":    ("feet", draw_shoes, "sneakers", ["white", "black", "gray"]),
    "dress_shoes": ("feet", draw_shoes, "dress", ["black", "brown"]),
    "glasses":     ("face", draw_accessory, "glasses", ["black"]),
    "sunglasses":  ("face", draw_accessory, "sunglasses", ["black"]),
    "tie":         ("neck", draw_accessory, "tie", ["red", "navy", "maroon", "blue"]),
    "bowtie":      ("neck", draw_accessory, "bowtie", ["black", "maroon"]),
}
Z = {"body": 10, "feet": 15, "legs": 20, "torso": 35, "neck": 45, "jacket": 55, "hair": 120, "face": 125}
FEMALE_HAIR = ["hair_long", "hair_bob", "hair_ponytail", "hair_fringe", "hair_short"]
MALE_HAIR = ["hair_short", "hair_side_part", "hair_messy", "hair_buzz", "hair_fringe"]

TIERS = {
    "broke":   {"head": ["head_m", "head_m", "head_f"], "hair": ["hair_messy", "hair_buzz", "hair_long", "hair_short", "hair_balding"],
                "torso": ["tshirt", "vneck"], "torso_colors": ["gray", "charcoal", "olive", "black"], "legs": ["pants", "shorts"],
                "legs_colors": ["gray", "brown", "charcoal"], "feet": ["slippers"], "jacket": [None], "neck": [None], "face": [None]},
    "working": {"head": ["head_m", "head_f"], "hair": ["hair_short", "hair_fringe", "hair_messy", "hair_bob", "hair_ponytail"],
                "torso": ["tshirt", "polo", "cardigan"], "legs": ["pants", "skirt"], "legs_colors": ["denim", "navy", "charcoal", "beige"],
                "feet": ["sneakers"], "jacket": [None], "neck": [None], "face": [None, None, "glasses"]},
    "middle":  {"head": ["head_m", "head_f"], "hair": ["hair_side_part", "hair_short", "hair_bob", "hair_long", "hair_ponytail"],
                "torso": ["shirt", "cardigan"], "legs": ["slacks", "skirt"], "feet": ["dress_shoes", "sneakers"],
                "jacket": [None], "neck": [None, "tie"], "face": [None, "glasses"]},
    "rich":    {"head": ["head_m", "head_f"], "hair": ["hair_side_part", "hair_long", "hair_bob"], "torso": ["shirt"],
                "torso_colors": ["white", "sky", "lavender"], "legs": ["slacks"], "feet": ["dress_shoes"], "jacket": ["suit"],
                "jacket_colors": ["navy", "charcoal", "gray"], "neck": ["tie", "bowtie"], "face": [None, "glasses", "sunglasses"]},
    "tycoon":  {"head": ["head_m", "head_f"], "hair": ["hair_side_part", "hair_balding", "hair_long"], "torso": ["shirt"], "torso_colors": ["white"],
                "legs": ["slacks"], "legs_colors": ["black", "charcoal"], "feet": ["dress_shoes"], "feet_colors": ["black"],
                "jacket": ["suit"], "jacket_colors": ["black", "charcoal"], "neck": ["tie", "bowtie"], "face": [None, "sunglasses"]},
}
ACTIVITIES = {
    "sleep":     {"anim": "hurt", "frame": 5, "props": ["bed", "zzz"]},
    "commute":   {"anim": "walk", "row": 3},
    "rest":      {"anim": "idle", "row": 2},
    "office":    {"anim": "work", "row": 2, "props": ["desk"]},
    "store":     {"anim": "idle", "row": 2, "props": ["counter"], "outfit": {"torso": ["tshirt", "white"], "jacket": ["vest", "green"]}, "badge": True},
    "cafe":      {"anim": "idle", "row": 2, "props": ["counter"], "outfit": {"jacket": ["apron", "brown"]}},
    "hospital":  {"anim": "idle", "row": 2, "outfit": {"torso": ["shirt", "sky"], "jacket": ["coat", "white"]}, "stethoscope": True},
    "lab":       {"anim": "work", "row": 2, "props": ["desk"], "outfit": {"jacket": ["coat", "white"]}},
    "manual":    {"anim": "walk", "row": 3, "props": ["box"]},
    "delivery":  {"anim": "sit", "row": 3, "props": ["scooter"], "helmet": True},
    "eat":       {"anim": "eat", "row": 2, "props": ["bowl"]},
    "eat_store": {"anim": "eat", "row": 2, "props": ["cupramen"]},
}

def sheet_for(drawer, kind, color, body, anim):
    frames, rows = ANIMS[anim]
    sh = Image.new("RGBA", (frames * S, rows * S), (0, 0, 0, 0))
    for r in range(rows):
        for c in range(frames):
            if anim == "hurt":     # lying down: the standing front frame turned on its side, settling over 6 frames
                p = pose(2, "idle", 0, body)
                cell = drawer(p, kind, color) if drawer else draw_skin(p, color)
                angle = min(90, c * 18)
                cell = cell.rotate(angle, resample=Image.NEAREST, center=(32, 40), translate=(0, int(angle / 90 * 14)))
            else:
                direction = r
                mirror = direction == 1
                p = pose(3 if mirror else direction, anim, c, body)
                cell = drawer(p, kind, color) if drawer else draw_skin(p, color)
                if mirror:
                    cell = cell.transpose(Image.FLIP_LEFT_RIGHT)
            sh.paste(cell, (c * S, r * S))
    return sh

def build(out):
    os.makedirs(out, exist_ok=True)
    layers, n = {}, 0
    for body in ("male", "female"):
        heads = {"head_m": "male", "head_f": "female"}
        for key, (slot, drawer, kind, colors) in WARDROBE.items():
            name = "body" if key == "skin" else key
            entry = layers.setdefault(name, {"z": Z[slot], "slot": slot, "bodies": {}})
            entry["bodies"][body] = {}
            for col in colors:
                for anim in ANIMS:
                    sheet_for(drawer, kind, col, body, anim).save(os.path.join(out, f"co_{name}_{body}_{col}_{anim}.png"), optimize=True); n += 1
                entry["bodies"][body][col] = list(ANIMS)
        # heads are part of the skin layer; register head keys that just pick the body type
    layers["head_m"] = {"z": 11, "slot": "head", "bodies": {"male": {c: [] for c in SKINS}}}
    layers["head_f"] = {"z": 11, "slot": "head", "bodies": {"female": {c: [] for c in SKINS}}}
    for body in ("male", "female"):
        for piece in ("helmet", "stethoscope", "badge"):
            for anim in ANIMS:
                frames, rows = ANIMS[anim]
                sh = Image.new("RGBA", (frames * S, rows * S), (0, 0, 0, 0))
                for r in range(rows):
                    for c in range(frames):
                        mirror = r == 1
                        p = pose(2 if anim == "hurt" else (3 if mirror else r), "idle" if anim == "hurt" else anim, c, body)
                        cell = draw_custom(p, piece)
                        if anim == "hurt":
                            angle = min(90, c * 18); cell = cell.rotate(angle, resample=Image.NEAREST, center=(32, 40), translate=(0, int(angle / 90 * 14)))
                        if mirror: cell = cell.transpose(Image.FLIP_LEFT_RIGHT)
                        sh.paste(cell, (c * S, r * S))
                sh.save(os.path.join(out, f"co_custom_{piece}_{body}_{anim}.png"), optimize=True); n += 1
    manifest = {"prefix": "co", "anims": {a: {"frames": f, "rows": r, "size": S} for a, (f, r) in ANIMS.items()},
                "layers": layers, "tiers": TIERS, "activities": ACTIVITIES, "skins": list(SKINS), "hairColors": list(HAIR),
                "custom": {"helmet": {"z": 130}, "stethoscope": {"z": 60}, "badge": {"z": 61}}}
    json.dump(manifest, open(os.path.join(out, "characters_original.json"), "w"), ensure_ascii=False, indent=1)
    return n

# ── examples ────────────────────────────────────────────────────────────────────────────
def look(m, tier, rng, activity=None):
    pool = m["tiers"][tier]; L = m["layers"]
    head = rng.choice(pool["head"]); body = "female" if head == "head_f" else "male"
    skin = rng.choice(m["skins"]); out = [("body", skin)]
    for slot in ("legs", "feet", "torso", "neck", "jacket", "face"):
        opts = [o for o in pool.get(slot, [None]) if o is None or body in L.get(o, {}).get("bodies", {})]
        it = rng.choice(opts) if opts else None
        if not it: continue
        cols = [c for c in pool.get(f"{slot}_colors", list(L[it]["bodies"][body])) if c in L[it]["bodies"][body]] or list(L[it]["bodies"][body])
        out.append((it, rng.choice(cols)))
    hairs = [h for h in pool["hair"] if body in L.get(h, {}).get("bodies", {})]
    if hairs: out.append((rng.choice(hairs), rng.choice(m["hairColors"])))
    a = m["activities"].get(activity or "", {})
    for slot, (it, col) in a.get("outfit", {}).items():
        if body in L.get(it, {}).get("bodies", {}):
            out = [(k, c) for k, c in out if L.get(k, {}).get("slot") != slot] + [(it, col)]
    if a.get("helmet"): out = [(k, c) for k, c in out if L.get(k, {}).get("slot") != "hair"] + [("custom_helmet", "")]
    if a.get("stethoscope"): out.append(("custom_stethoscope", ""))
    if a.get("badge"): out.append(("custom_badge", ""))
    z = lambda k: L[k]["z"] if k in L else m["custom"][k[7:]]["z"]
    return body, sorted(out, key=lambda kc: z(kc[0]))

def frame(out, body, lk, anim, row, col):
    cv = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    for k, c in lk:
        f = os.path.join(out, f"co_{k}_{body}_{anim}.png" if k.startswith("custom_") else f"co_{k}_{body}_{c}_{anim}.png")
        if os.path.exists(f):
            cv.alpha_composite(Image.open(f).crop((col * S, row * S, col * S + S, row * S + S)))
    return cv

def examples(out, dest, seed=11):
    m = json.load(open(os.path.join(out, "characters_original.json"))); rng = random.Random(seed)
    os.makedirs(dest, exist_ok=True); bg = (236, 232, 222, 255)
    tiers = ["broke", "working", "middle", "rich", "tycoon"]
    # 1) line-up: 3 looks per tier (front, idle)
    k = 3; sc = 3
    lu = Image.new("RGBA", (len(tiers) * k * S * sc // 2, S * sc // 2 + 0), bg)
    x = 0
    for t in tiers:
        for _ in range(k):
            b, lk = look(m, t, rng); lu.alpha_composite(frame(out, b, lk, "idle", 2, 0).resize((S * sc // 2, S * sc // 2), Image.NEAREST), (x, 0)); x += S * sc // 2
    lu.save(os.path.join(dest, "original_lineup_by_tier.png"))
    # 2) four directions + animations for one character
    b, lk = look(m, "middle", random.Random(4))
    rows = [("walk", r, list(range(1, 9))) for r in (2, 3, 0, 1)] + [("eat", 2, [0, 1, 2, 3]), ("work", 2, [0, 1, 2, 3]), ("sit", 3, [0])]
    grid = Image.new("RGBA", (8 * S * 2, len(rows) * S * 2), bg)
    for ri, (a, r, cols) in enumerate(rows):
        for ci, c in enumerate(cols):
            grid.alpha_composite(frame(out, b, lk, a, r, c).resize((S * 2, S * 2), Image.NEAREST), (ci * S * 2, ri * S * 2))
    grid.save(os.path.join(dest, "original_animation_sheet.png"))
    # 3) daily-life scenes with props (props are Vaultly's own ch_prop_*)
    scenes = [("broke", "eat_store"), ("broke", "delivery"), ("working", "store"), ("working", "office"), ("middle", "hospital"),
              ("rich", "commute"), ("tycoon", "rest"), ("working", "sleep")]
    W = 96; img = Image.new("RGBA", (len(scenes) * W * 3, W * 3), bg)
    for i, (t, act) in enumerate(scenes):
        b, lk = look(m, t, rng, act); a = m["activities"][act]
        tile = Image.new("RGBA", (W, W), (0, 0, 0, 0)); props = a.get("props", [])
        for pr in [p for p in props if p in ("bed", "scooter")]:
            tile.alpha_composite(Image.open(os.path.join(out, f"ch_prop_{pr}.png")), (16, 22 if pr == "bed" else 20))
        tile.alpha_composite(frame(out, b, lk, a["anim"], 0 if a["anim"] == "hurt" else a.get("row", 2), a.get("frame", 0)), (16, 4))
        for pr in [p for p in props if p not in ("bed", "scooter")]:
            tile.alpha_composite(Image.open(os.path.join(out, f"ch_prop_{pr}.png")), (16, 14 if pr in ("desk", "counter") else 10))
        img.alpha_composite(tile.resize((W * 3, W * 3), Image.NEAREST), (i * W * 3, 0))
    img.save(os.path.join(dest, "original_daily_life.png"))
    # 4) walking GIF, one per tier
    lks = [look(m, t, random.Random(20 + i)) for i, t in enumerate(tiers)]
    frames = []
    for f in range(1, 9):
        s = Image.new("RGBA", (len(tiers) * S * 3, S * 3), bg)
        for i, (b, lk) in enumerate(lks):
            s.alpha_composite(frame(out, b, lk, "walk", 2, f).resize((S * 3, S * 3), Image.NEAREST), (i * S * 3, 0))
        frames.append(s.convert("RGB"))
    frames[0].save(os.path.join(dest, "original_walk_by_tier.gif"), save_all=True, append_images=frames[1:], duration=110, loop=0)
    return dest

if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default=OUT); ap.add_argument("--examples"); a = ap.parse_args()
    print(f"wrote {build(a.out)} original sprite sheets")
    if a.examples: print(examples(a.out, a.examples))
