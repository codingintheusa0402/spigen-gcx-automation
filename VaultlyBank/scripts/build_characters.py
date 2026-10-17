#!/usr/bin/env python3
"""Build Vaultly's pixel-art character assets from the Universal LPC Spritesheet Character Generator.

    git clone --depth 1 --filter=blob:none --sparse https://github.com/LiberatedPixelCup/Universal-LPC-Spritesheet-Character-Generator lpc
    python3 scripts/build_characters.py --lpc ../lpc [--examples out_dir]

What it does
  · reads the curated wardrobe below (WARDROBE) and each item's sheet definition in the LPC repo
  · keeps only art usable in a closed-source App Store app: every item must offer CC-BY-SA / OGA-BY / CC-BY / CC0
    (GPL-only art is refused — GPL conflicts with App Store terms). We use the art under its CC-BY-SA/OGA-BY license.
  · recolors base-palette sprites with LPC's own palette files (or picks the pre-coloured file), for the colours the
    tier outfits use, and writes flat bundle files  Vaultly/Characters/ch_<item>_<body>_<color>_<anim>.png
  · draws Vaultly's own pieces from scratch (delivery helmet, stethoscope, name tag, props: scooter, desk, bowl,
    box, bed) — original art, not derived from LPC pixels; only head positions are measured to place the helmet
  · writes Characters/characters.json (layers, z-order, colours, animations, tier outfit pools, activity map) and
    Characters/credits.json (every LPC file used, its authors, licences and URLs) for the in-app credits screen
"""
import argparse, json, glob, os, random, sys
from PIL import Image, ImageDraw

ANIMS = {"idle": 2, "walk": 9, "sit": 3, "hurt": 6}       # frames per row; hurt is a single row (falls down → lying)
OK_LICENSES = {"CC-BY-SA 3.0", "CC-BY-SA 4.0", "OGA-BY 3.0", "CC-BY 3.0", "CC-BY 4.0", "CC0"}
BODIES = ("male", "female")
PALETTE_BASE = {"cloth": "white", "body": "light", "hair": "orange", "eye": "blue", "metal": "steel", "wood": "maple"}

SKINS = ["light", "amber", "olive", "taupe", "bronze", "brown"]
HAIR_COLORS = ["black", "raven", "dark_brown", "chestnut", "dark_gray", "gray", "light_brown"]

# item key → (LPC definition name, colours to build, slot). Slot decides z-order fallbacks and what the outfit fills.
WARDROBE = {
    # heads (skin colours applied)
    "head_m": ("heads_human_male", SKINS, "head"), "head_f": ("heads_human_female", SKINS, "head"),
    "head_m_gaunt": ("heads_human_male_gaunt", SKINS, "head"), "head_m_plump": ("heads_human_male_plump", SKINS, "head"),
    # hair
    **{f"hair_{h}": (f"hair_{h}", HAIR_COLORS, "hair") for h in
       ["plain", "parted", "messy1", "buzzcut", "unkempt", "bedhead", "pixie", "bangs", "long_straight", "swoop", "balding", "page"]},
    # tops
    "tshirt": ("torso_clothes_tshirt", ["gray", "charcoal", "brown", "tan", "white", "navy", "blue", "forest", "maroon", "sky"], "torso"),
    "tshirt_vneck": ("torso_clothes_tshirt_vneck", ["gray", "white", "black", "slate", "rose"], "torso"),
    "polo": ("torso_clothes_shortsleeve_polo", ["navy", "white", "forest", "sky", "maroon"], "torso"),
    "cardigan": ("torso_clothes_longsleeve2_cardigan", ["gray", "navy", "tan", "maroon", "forest"], "torso"),
    "shirt": ("torso_clothes_longsleeve2_buttoned", ["white", "sky", "slate", "lavender", "pink"], "torso"),
    "suit_jacket": ("torso_jacket_collared", ["black", "charcoal", "navy", "gray"], "jacket"),
    "trench": ("torso_jacket_trench", ["white", "gray", "dark_gray"], "jacket"),     # white (lightened gray) = doctor's coat
    "vest": ("torso_clothes_vest", ["green", "blue", "black"], "jacket"),              # green = convenience-store vest
    "apron": ("torso_aprons_apron", ["brown", "black", "green"], "jacket"),            # cafe / restaurant
    "necktie": ("neck_necktie", ["red", "maroon", "navy", "black", "blue"], "neck"),
    "bowtie": ("neck_bowtie", ["black", "maroon", "navy"], "neck"),
    # bottoms & shoes
    "pants": ("legs_pants", ["gray", "brown", "slate", "navy", "charcoal", "blue", "tan"], "legs"),
    "formal_pants": ("legs_formal", ["black", "charcoal", "navy", "gray"], "legs"),
    "cuffed": ("legs_cuffed", ["blue", "navy", "tan"], "legs"),
    "shorts": ("legs_shorts", ["gray", "brown", "navy"], "legs"),
    "skirt": ("legs_skirt_straight", ["black", "navy", "gray", "maroon"], "legs"),
    "shoes": ("feet_shoes_basic", ["black", "brown"], "feet"),
    "shoes2": ("feet_shoes_revised", ["black", "brown", "white"], "feet"),
    "boots": ("feet_boots_basic", ["black", "brown"], "feet"),
    "slippers": ("feet_slippers", ["gray", "blue"], "feet"),
    "sandals": ("feet_sandals", ["brown", "black"], "feet"),
    "glasses": ("facial_glasses", ["black", "brown"], "face"),
    "round_glasses": ("facial_glasses_round", ["black", "gold"], "face"),
    "sunglasses": ("facial_glasses_sunglasses", ["black"], "face"),
}

# Outfit pools by tier: one choice per slot (None = leave empty). Colours are picked from the item's list.
TIER_OUTFITS = {
    "broke":   {"head": ["head_m_gaunt", "head_m", "head_f"], "hair": ["hair_unkempt", "hair_bedhead", "hair_messy1", "hair_buzzcut", "hair_long_straight"],
                "torso": ["tshirt", "tshirt_vneck"], "torso_colors": ["gray", "charcoal", "slate", "navy", "white"],
                "legs": ["pants", "shorts"], "legs_colors": ["gray", "brown", "slate"], "feet": ["slippers", "sandals"], "jacket": [None], "neck": [None], "face": [None]},
    "working": {"head": ["head_m", "head_f", "head_m_plump"], "hair": ["hair_plain", "hair_messy1", "hair_buzzcut", "hair_bangs", "hair_page", "hair_pixie"],
                "torso": ["tshirt", "polo", "cardigan", "shirt"], "legs": ["pants", "cuffed"], "feet": ["shoes", "boots", "shoes2"],
                "jacket": [None], "neck": [None], "face": [None, None, "glasses"]},
    "middle":  {"head": ["head_m", "head_f"], "hair": ["hair_parted", "hair_plain", "hair_swoop", "hair_page", "hair_pixie", "hair_long_straight"],
                "torso": ["shirt", "cardigan", "polo"], "legs": ["formal_pants", "cuffed"], "feet": ["shoes2", "shoes"],
                "jacket": [None, None], "neck": [None, "necktie"], "face": [None, "glasses", "round_glasses"]},
    "rich":    {"head": ["head_m", "head_f", "head_m_plump"], "hair": ["hair_parted", "hair_swoop", "hair_long_straight", "hair_page"],
                "torso": ["shirt"], "torso_colors": ["white", "sky", "pink", "lavender"], "legs": ["formal_pants"], "legs_colors": ["black", "charcoal", "navy"],
                "feet": ["shoes2", "shoes"], "jacket": ["suit_jacket"], "neck": ["necktie", "bowtie"], "face": [None, "round_glasses", "sunglasses"]},
    "tycoon":  {"head": ["head_m", "head_f"], "hair": ["hair_parted", "hair_swoop", "hair_balding", "hair_long_straight"],
                "torso": ["shirt"], "torso_colors": ["white"], "legs": ["formal_pants"], "legs_colors": ["black", "charcoal"],
                "feet": ["shoes2"], "feet_colors": ["black"], "jacket": ["suit_jacket"], "jacket_colors": ["black", "charcoal"],
                "neck": ["necktie", "bowtie"], "face": [None, "sunglasses"]},
}

# Schedule activity → animation, facing row, outfit overrides (work uniforms) and props drawn around the character.
ACTIVITIES = {
    "sleep":     {"anim": "hurt", "frame": 5, "props": ["bed", "zzz"]},
    "commute":   {"anim": "walk", "row": 3},
    "rest":      {"anim": "idle", "row": 2},
    "office":    {"anim": "sit", "row": 2, "props": ["desk"], "outfit": {}},
    "store":     {"anim": "idle", "row": 2, "props": ["counter"], "outfit": {"torso": ["tshirt", "white"], "jacket": [["vest", "green"], ["apron", "green"]]}, "badge": True},
    "cafe":      {"anim": "idle", "row": 2, "props": ["counter"], "outfit": {"jacket": [["apron", "brown"], ["apron", "black"]]}},
    "hospital":  {"anim": "idle", "row": 2, "outfit": {"torso": ["shirt", "white"], "jacket": ["trench", "white"]}, "stethoscope": True},
    "lab":       {"anim": "sit", "row": 2, "props": ["desk"], "outfit": {"torso": ["shirt", "white"], "jacket": ["trench", "white"]}},
    "manual":    {"anim": "walk", "row": 3, "props": ["box"]},
    "delivery":  {"anim": "sit", "row": 3, "props": ["scooter"], "helmet": True},
    "eat":       {"anim": "sit", "row": 2, "props": ["bowl"]},
    "eat_store": {"anim": "idle", "row": 2, "props": ["cupramen"]},
}


def load_defs(lpc):
    return {os.path.basename(f)[:-5]: f for f in glob.glob(os.path.join(lpc, "sheet_definitions", "**", "*.json"), recursive=True)}


def palettes(lpc, material):
    f = os.path.join(lpc, "palette_definitions", material, f"{material}_ulpc.json")
    return json.load(open(f)) if os.path.exists(f) else {}


def hexrgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def recolor(img, src, dst):
    """Swap palette colours (±1 per channel, like the LPC generator)."""
    img = img.convert("RGBA")
    pairs = [(hexrgb(s), hexrgb(d)) for s, d in zip(src, dst)]
    px = img.load()
    for y in range(img.height):
        for x in range(img.width):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            for (sr, sg, sb), (dr, dg, db) in pairs:
                if abs(r - sr) <= 1 and abs(g - sg) <= 1 and abs(b - sb) <= 1:
                    px[x, y] = (dr, dg, db, a)
                    break
    return img


def item_sheet(lpc, layer_dir, anim, color, material, pal):
    """Return an RGBA sheet for one anim+colour, or None if this item doesn't have that animation."""
    base = os.path.join(lpc, "spritesheets", layer_dir.rstrip("/"))
    pre = os.path.join(base, anim, f"{color}.png")
    if os.path.exists(pre):
        return Image.open(pre).convert("RGBA")
    gray = os.path.join(base, anim, "gray.png")
    if color == "white" and os.path.exists(gray):      # e.g. the trench only ships in grays → lab/doctor's coat
        img = Image.open(gray).convert("RGBA"); px = img.load()
        for y in range(img.height):
            for x in range(img.width):
                r, g, b, a = px[x, y]
                if a:
                    lum = (r + g + b) / 3
                    v = int(min(255, 150 + lum * 0.75))
                    px[x, y] = (v, v, min(255, v + 4), a) if lum > 40 else (r, g, b, a)
        return img
    flat = os.path.join(base, f"{anim}.png")
    if not os.path.exists(flat):
        return None
    img = Image.open(flat)
    if material and pal and color in pal:
        src = pal.get(PALETTE_BASE.get(material, "white"))
        if src:
            return recolor(img, src, pal[color])
    return img.convert("RGBA")


def build(lpc, out):
    defs = load_defs(lpc)
    os.makedirs(out, exist_ok=True)
    manifest = {"anims": {a: {"frames": n, "rows": 1 if a == "hurt" else 4, "size": 64} for a, n in ANIMS.items()},
                "layers": {}, "tiers": TIER_OUTFITS, "activities": ACTIVITIES, "skins": SKINS, "hairColors": HAIR_COLORS}
    credits, refused, written = {}, [], 0
    body_def = json.load(open(defs["body"]))
    # Body (skin colours)
    bodypal = palettes(lpc, "body")
    manifest["layers"]["body"] = {"z": 10, "slot": "body", "bodies": {}}
    for b in BODIES:
        layer_dir = body_def["layer_1"][b]
        manifest["layers"]["body"]["bodies"][b] = {}
        for skin in SKINS:
            anims = []
            for a in ANIMS:
                im = item_sheet(lpc, layer_dir, a, skin, "body", bodypal)
                if im is None:
                    continue
                im.save(os.path.join(out, f"ch_body_{b}_{skin}_{a}.png"), optimize=True); written += 1
                anims.append(a)
            manifest["layers"]["body"]["bodies"][b][skin] = anims
    for c in body_def.get("credits", []):
        credits[c["file"]] = c

    for key, (defname, colors, slot) in WARDROBE.items():
        if defname not in defs:
            refused.append((key, "definition not found")); continue
        d = json.load(open(defs[defname]))
        lic = {l for c in d.get("credits", []) for l in c["licenses"]}
        if not lic & OK_LICENSES:
            refused.append((key, f"GPL-only ({sorted(lic)})")); continue
        layer = d["layer_1"]
        rec = d.get("recolors") or {}
        material = rec.get("material") or (rec.get("color_1") or {}).get("material")
        pal = palettes(lpc, material) if material else {}
        entry = {"z": layer.get("zPos", 50), "slot": slot, "bodies": {}}
        for b in BODIES:
            if not layer.get(b):
                continue
            entry["bodies"][b] = {}
            for col in colors:
                anims = []
                for a in ANIMS:
                    im = item_sheet(lpc, layer[b], a, col, material, pal)
                    if im is None:
                        continue
                    if slot == "head":   # heads also carry eye colour (second palette): dark brown eyes
                        eyepal = palettes(lpc, "eye")
                        if eyepal.get("brown"):
                            im = recolor(im, eyepal["blue"], eyepal["brown"])
                    im.save(os.path.join(out, f"ch_{key}_{b}_{col}_{a}.png"), optimize=True); written += 1
                    anims.append(a)
                if "idle" not in anims and "walk" in anims:   # older LPC items: stand still = walk frame 0
                    w = Image.open(os.path.join(out, f"ch_{key}_{b}_{col}_walk.png"))
                    idle = Image.new("RGBA", (128, 256), (0, 0, 0, 0))
                    for r in range(4):
                        cell = w.crop((0, r * 64, 64, r * 64 + 64))
                        idle.paste(cell, (0, r * 64)); idle.paste(cell, (64, r * 64))
                    idle.save(os.path.join(out, f"ch_{key}_{b}_{col}_idle.png"), optimize=True); written += 1
                    anims.append("idle")
                if anims:
                    entry["bodies"][b][col] = anims
        if entry["bodies"]:
            manifest["layers"][key] = entry
            for c in d.get("credits", []):
                credits[c["file"]] = c
    # Vaultly's own pieces (drawn from scratch below)
    written += draw_custom(lpc, out, manifest)
    manifest["custom"] = {"helmet": {"z": 120}, "stethoscope": {"z": 101}, "badge": {"z": 96},
                          "props": ["bed", "zzz", "desk", "counter", "box", "scooter", "bowl", "cupramen"]}
    json.dump(manifest, open(os.path.join(out, "characters.json"), "w"), ensure_ascii=False, indent=1)
    used = [{"file": k, "authors": v.get("authors", []), "licenses": [l for l in v.get("licenses", []) if l in OK_LICENSES] or v.get("licenses", []),
             "urls": v.get("urls", []), "notes": v.get("notes", "")} for k, v in sorted(credits.items())]
    json.dump({"source": "Universal LPC Spritesheet Character Generator — https://github.com/LiberatedPixelCup/Universal-LPC-Spritesheet-Character-Generator",
               "used_under": "CC-BY-SA 3.0 / OGA-BY 3.0 (and CC-BY / CC0 where offered)", "files": used,
               "vaultly_original": ["delivery helmet", "stethoscope", "name badge", "scooter", "desk & laptop", "store counter", "box", "bowl", "cup ramen", "bed"]},
              open(os.path.join(out, "credits.json"), "w"), ensure_ascii=False, indent=1)
    return written, refused, len(used)


# ── Original Vaultly pieces ──────────────────────────────────────────────────────────────
OUT = (24, 20, 28, 255)   # LPC-style dark outline


def head_boxes(lpc, body):
    """Per-frame bounding box of the head (from the LPC head layer's alpha) so the helmet sits where the head is."""
    boxes = {}
    head_dir = "head/heads/human/male/" if body == "male" else "head/heads/human/female/"
    for a, n in ANIMS.items():
        f = os.path.join(lpc, "spritesheets", head_dir, f"{a}.png")
        if not os.path.exists(f):
            continue
        im = Image.open(f).convert("RGBA")
        rows = 1 if a == "hurt" else 4
        for r in range(rows):
            for c in range(n):
                cell = im.crop((c * 64, r * 64, c * 64 + 64, r * 64 + 64))
                boxes[(a, r, c)] = cell.getbbox()
    return boxes


def draw_helmet(dr, box, row):
    """Delivery helmet: rounded orange shell, white stripe, dark visor, chin strap. Original art."""
    if not box:
        return
    x0, y0, x1, y1 = box
    cx = (x0 + x1) // 2
    top = y0
    shell, light, stripe, visor = (232, 112, 36, 255), (250, 160, 80, 255), (245, 245, 240, 255), (40, 50, 70, 255)
    dr.ellipse((cx - 8, top, cx + 8, top + 12), fill=OUT)
    dr.ellipse((cx - 7, top + 1, cx + 7, top + 11), fill=shell)
    dr.rectangle((cx - 7, top + 7, cx + 7, top + 9), fill=shell)
    dr.line((cx - 5, top + 3, cx - 1, top + 2), fill=light)
    if row in (0,):            # facing away: stripe down the middle
        dr.rectangle((cx - 1, top + 1, cx + 1, top + 10), fill=stripe)
    elif row == 2:             # facing us: visor across the eyes
        dr.rectangle((cx - 6, top + 8, cx + 6, top + 10), fill=visor)
        dr.line((cx - 5, top + 8, cx + 5, top + 8), fill=(120, 140, 170, 255))
        dr.rectangle((cx - 1, top + 1, cx + 1, top + 6), fill=stripe)
    else:                      # side: visor on the facing side
        sx = cx + (3 if row == 3 else -6)
        dr.rectangle((sx, top + 7, sx + 3, top + 10), fill=visor)
        dr.rectangle((cx - 6, top + 4, cx + 6, top + 5), fill=stripe)


def draw_custom(lpc, out, manifest):
    n = 0
    for body in BODIES:
        boxes = head_boxes(lpc, body)
        for a, frames in ANIMS.items():
            rows = 1 if a == "hurt" else 4
            sheet = Image.new("RGBA", (frames * 64, rows * 64), (0, 0, 0, 0))
            steth = Image.new("RGBA", sheet.size, (0, 0, 0, 0))
            badge = Image.new("RGBA", sheet.size, (0, 0, 0, 0))
            for r in range(rows):
                for c in range(frames):
                    box = boxes.get((a, r, c))
                    cell = Image.new("RGBA", (64, 64), (0, 0, 0, 0)); draw_helmet(ImageDraw.Draw(cell), box, r)
                    sheet.paste(cell, (c * 64, r * 64))
                    if box and r == 2:   # front view only: stethoscope round the neck, badge on the chest
                        x0, y0, x1, y1 = box; cx = (x0 + x1) // 2
                        sd = ImageDraw.Draw(steth)
                        ox, oy = c * 64, r * 64
                        tube = (55, 60, 75, 255)
                        sd.arc((ox + cx - 6, oy + y1 - 3, ox + cx + 6, oy + y1 + 7), 20, 160, fill=tube, width=2)
                        sd.line((ox + cx + 4, oy + y1 + 5, ox + cx + 4, oy + y1 + 11), fill=tube, width=1)
                        sd.ellipse((ox + cx + 2, oy + y1 + 10, ox + cx + 6, oy + y1 + 14), fill=(190, 195, 205, 255), outline=tube)
                        bd = ImageDraw.Draw(badge)
                        bd.rectangle((ox + cx + 2, oy + y1 + 6, ox + cx + 6, oy + y1 + 8), fill=(250, 250, 245, 255), outline=OUT)
            for name, im in (("helmet", sheet), ("stethoscope", steth), ("badge", badge)):
                im.save(os.path.join(out, f"ch_custom_{name}_{body}_{a}.png"), optimize=True); n += 1
    n += draw_props(out)
    return n


def draw_props(out):
    """64×64 props, drawn pixel by pixel in an LPC-like palette. Original art."""
    def canvas(): im = Image.new("RGBA", (64, 64), (0, 0, 0, 0)); return im, ImageDraw.Draw(im)
    n = 0
    # Scooter (side view, facing right): body, seat, wheels, delivery box
    im, d = canvas()
    for wx in (14, 48):
        d.ellipse((wx - 7, 47, wx + 7, 61), fill=OUT); d.ellipse((wx - 5, 49, wx + 5, 59), fill=(60, 60, 66, 255)); d.ellipse((wx - 2, 52, wx + 2, 56), fill=(180, 180, 190, 255))
    d.polygon([(10, 50), (22, 40), (46, 40), (54, 50), (44, 52), (18, 52)], fill=OUT)
    d.polygon([(12, 49), (23, 41), (45, 41), (52, 49), (43, 51), (19, 51)], fill=(200, 40, 40, 255))
    d.line((46, 40, 52, 26), fill=OUT, width=3); d.line((50, 26, 56, 26), fill=OUT, width=2)
    d.rectangle((2, 26, 18, 40), fill=OUT); d.rectangle((3, 27, 17, 39), fill=(240, 240, 235, 255)); d.rectangle((3, 31, 17, 33), fill=(232, 112, 36, 255))
    im.save(os.path.join(out, "ch_prop_scooter.png")); n += 1
    # Desk with a laptop (front)
    im, d = canvas()
    d.rectangle((4, 40, 60, 46), fill=OUT); d.rectangle((5, 41, 59, 45), fill=(150, 105, 70, 255))
    d.rectangle((8, 46, 11, 63), fill=(110, 75, 50, 255)); d.rectangle((53, 46, 56, 63), fill=(110, 75, 50, 255))
    d.rectangle((20, 28, 44, 40), fill=OUT); d.rectangle((21, 29, 43, 38), fill=(120, 180, 230, 255)); d.rectangle((18, 39, 46, 41), fill=(170, 170, 180, 255))
    im.save(os.path.join(out, "ch_prop_desk.png")); n += 1
    # Store counter with a register
    im, d = canvas()
    d.rectangle((0, 42, 63, 63), fill=OUT); d.rectangle((1, 43, 62, 62), fill=(225, 225, 220, 255)); d.rectangle((1, 43, 62, 46), fill=(60, 150, 90, 255))
    d.rectangle((40, 32, 56, 43), fill=OUT); d.rectangle((41, 33, 55, 42), fill=(70, 70, 80, 255)); d.rectangle((43, 35, 53, 38), fill=(130, 230, 150, 255))
    im.save(os.path.join(out, "ch_prop_counter.png")); n += 1
    # Cardboard box (carried)
    im, d = canvas()
    d.rectangle((22, 30, 42, 46), fill=OUT); d.rectangle((23, 31, 41, 45), fill=(196, 150, 96, 255)); d.line((23, 37, 41, 37), fill=(160, 115, 70, 255)); d.rectangle((30, 31, 33, 45), fill=(220, 190, 140, 255))
    im.save(os.path.join(out, "ch_prop_box.png")); n += 1
    # Bowl with chopsticks (on a small table)
    im, d = canvas()
    d.rectangle((14, 46, 50, 50), fill=OUT); d.rectangle((15, 47, 49, 49), fill=(150, 105, 70, 255))
    d.chord((22, 36, 42, 50), 0, 180, fill=OUT); d.chord((23, 37, 41, 48), 0, 180, fill=(240, 240, 235, 255)); d.rectangle((24, 41, 40, 43), fill=(230, 160, 60, 255))
    d.line((36, 30, 44, 42), fill=(180, 140, 90, 255), width=1); d.line((38, 30, 45, 41), fill=(180, 140, 90, 255), width=1)
    im.save(os.path.join(out, "ch_prop_bowl.png")); n += 1
    # Cup ramen (convenience-store lunch)
    im, d = canvas()
    d.polygon([(26, 34), (38, 34), (36, 48), (28, 48)], fill=OUT); d.polygon([(27, 35), (37, 35), (35, 47), (29, 47)], fill=(250, 245, 235, 255))
    d.rectangle((27, 38, 37, 41), fill=(210, 50, 40, 255)); d.line((30, 26, 30, 34), fill=(200, 200, 200, 160)); d.line((34, 25, 34, 33), fill=(200, 200, 200, 160))
    im.save(os.path.join(out, "ch_prop_cupramen.png")); n += 1
    # Bed / blanket (lying character is drawn on top)
    im, d = canvas()
    d.rectangle((2, 30, 62, 50), fill=OUT); d.rectangle((3, 31, 61, 49), fill=(120, 150, 210, 255)); d.rectangle((3, 31, 16, 49), fill=(245, 245, 240, 255))
    d.rectangle((4, 50, 7, 58), fill=(110, 75, 50, 255)); d.rectangle((57, 50, 60, 58), fill=(110, 75, 50, 255))
    im.save(os.path.join(out, "ch_prop_bed.png")); n += 1
    # Zzz
    im, d = canvas()
    for i, (x, y, s) in enumerate([(40, 18, 5), (48, 9, 7)]):
        d.line((x, y, x + s, y), fill=(255, 255, 255, 255), width=2); d.line((x + s, y, x, y + s), fill=(255, 255, 255, 255), width=2); d.line((x, y + s, x + s, y + s), fill=(255, 255, 255, 255), width=2)
    im.save(os.path.join(out, "ch_prop_zzz.png")); n += 1
    return n


# ── Example renders (for review) ─────────────────────────────────────────────────────────

def compose(out, layers, anim, row, col, body):
    canvas = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    for key, color in sorted(layers, key=lambda kc: kc[2])[::1] if False else layers:
        f = os.path.join(out, f"ch_{key}_{body}_{color}_{anim}.png") if not key.startswith("custom_") else os.path.join(out, f"ch_{key}_{body}_{anim}.png")
        if not os.path.exists(f):
            continue
        sheet = Image.open(f)
        r = 0 if anim == "hurt" else row
        canvas.alpha_composite(sheet.crop((col * 64, r * 64, col * 64 + 64, r * 64 + 64)))
    return canvas


def random_look(manifest, tier, rng, activity=None):
    pool = manifest["tiers"][tier]
    L = manifest["layers"]
    head = rng.choice(pool["head"]); body = "female" if head.endswith("_f") else "male"
    skin = rng.choice(manifest["skins"])
    look = [("body", skin), (head, skin)]
    def pick(slot):
        opts = [o for o in pool.get(slot, [None]) if o is None or (o in L and body in L[o]["bodies"])]
        item = rng.choice(opts) if opts else None
        if not item:
            return
        cols = [c for c in pool.get(f"{slot}_colors", list(L[item]["bodies"][body])) if c in L[item]["bodies"][body]] or list(L[item]["bodies"][body])
        look.append((item, rng.choice(cols)))
    for s in ("legs", "feet", "torso", "neck", "jacket", "face"):
        pick(s)
    hair = [h for h in pool["hair"] if h in L and body in L[h]["bodies"]]
    if hair:
        look.append((rng.choice(hair), rng.choice(manifest["hairColors"])))
    act = manifest["activities"].get(activity or "", {})
    for slot, choice in act.get("outfit", {}).items():
        alts = choice if isinstance(choice[0], list) else [choice]
        for item, col in alts:
            if item in L and body in L[item]["bodies"] and col in L[item]["bodies"][body]:
                look = [(k, c) for k, c in look if L.get(k, {}).get("slot") != slot] + [(item, col)]
                break
    if act.get("helmet"):
        look = [(k, c) for k, c in look if L.get(k, {}).get("slot") != "hair"] + [("custom_helmet", "")]
    if act.get("stethoscope"):
        look.append(("custom_stethoscope", ""))
    if act.get("badge"):
        look.append(("custom_badge", ""))
    z = lambda k: L[k]["z"] if k in L else manifest["custom"].get(k.replace("custom_", ""), {}).get("z", 200)
    return body, sorted(look, key=lambda kc: z(kc[0]))


def examples(out, dest, seed=7):
    manifest = json.load(open(os.path.join(out, "characters.json")))
    rng = random.Random(seed)
    os.makedirs(dest, exist_ok=True)
    SCALE = 4
    titles = {"broke": "거지", "working": "서민", "middle": "중산층", "rich": "부자", "tycoon": "재벌"}
    # 1) A line-up: 3 random looks per tier, walking toward us, plus a back view
    tiers = list(titles)
    sheet = Image.new("RGBA", (len(tiers) * 3 * 64 * SCALE // 2, 64 * SCALE // 2 + 0), (236, 232, 222, 255))
    x = 0
    for t in tiers:
        for i in range(3):
            body, look = random_look(manifest, t, rng)
            fr = compose(out, look, "walk", 2, 0, body).resize((64 * SCALE // 2, 64 * SCALE // 2), Image.NEAREST)
            sheet.alpha_composite(fr, (x, 0)); x += 64 * SCALE // 2
    sheet.save(os.path.join(dest, "lineup_by_tier.png"))
    # 2) One character per tier doing a daily activity, with props
    scenes = [("broke", "eat_store"), ("working", "delivery"), ("working", "store"), ("middle", "office"), ("middle", "hospital"), ("rich", "commute"), ("tycoon", "rest"), ("broke", "sleep")]
    W = 96
    sc = Image.new("RGBA", (len(scenes) * W * 3, W * 3), (236, 232, 222, 255))
    for i, (t, act) in enumerate(scenes):
        body, look = random_look(manifest, t, rng, act)
        a = manifest["activities"][act]
        tile = Image.new("RGBA", (W, W), (0, 0, 0, 0))
        props = a.get("props", [])
        under = [p for p in props if p in ("bed", "scooter")]
        over = [p for p in props if p not in under]
        for p in under:
            tile.alpha_composite(Image.open(os.path.join(out, f"ch_prop_{p}.png")), (16, 22 if p == "bed" else 20))
        ch = compose(out, look, a["anim"], a.get("row", 2), a.get("frame", 0), body)
        tile.alpha_composite(ch, (16, 4 if a["anim"] != "hurt" else 4))
        for p in over:
            tile.alpha_composite(Image.open(os.path.join(out, f"ch_prop_{p}.png")), (16, 14 if p in ("desk", "counter") else 10))
        sc.alpha_composite(tile.resize((W * 3, W * 3), Image.NEAREST), (i * W * 3, 0))
    sc.save(os.path.join(dest, "daily_life_scenes.png"))
    # 3) Animated walk GIF for each tier
    frames = []
    looks = [random_look(manifest, t, rng) for t in tiers]
    for f in range(1, 9):
        strip = Image.new("RGBA", (len(tiers) * 64 * 3, 64 * 3), (236, 232, 222, 255))
        for i, (body, look) in enumerate(looks):
            strip.alpha_composite(compose(out, look, "walk", 2, f, body).resize((192, 192), Image.NEAREST), (i * 192, 0))
        frames.append(strip.convert("RGB"))
    frames[0].save(os.path.join(dest, "walk_by_tier.gif"), save_all=True, append_images=frames[1:], duration=110, loop=0)
    return [os.path.join(dest, f) for f in ("lineup_by_tier.png", "daily_life_scenes.png", "walk_by_tier.gif")]


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--lpc", required=True)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "..", "Vaultly", "Characters"))
    ap.add_argument("--examples")
    a = ap.parse_args()
    written, refused, ncred = build(a.lpc, a.out)
    print(f"wrote {written} sprite sheets · {ncred} credited LPC files · refused: {refused or 'none'}")
    if a.examples:
        print("\n".join(examples(a.out, a.examples)))
