"""Line-up normalisation shared by find_siren.py.

A SIREN "case" = one product line-up (device generation + product model, all sizes/colors
merged — e.g. iPhone 18 Pro + 18 Pro Max Tough Armor Pro) × one defect type.
SKU -> 기종명/모델명 comes from the product master 'Data' tab (1fx9K4r2T9...), whose 기종명
already lists compatible devices together ('iPhone 18 Pro Max/17 Pro Max'); the FIRST
device there decides the generation, so 17 Pro Max claims and 18 Pro Max SKUs merge."""
import re

SERIES_PATTERNS = [
    r"iPhone \d+e", r"iPhone Air", r"iPhone \d+",
    r"Galaxy Z Fold ?\d+", r"Galaxy Z Flip ?\d+", r"Galaxy S\d+ FE", r"Galaxy S\d+", r"Galaxy A\d+",
    r"Galaxy Tab S\d+", r"Pixel \d+ Pro Fold", r"Pixel \d+a", r"Pixel \d+",
    r"Apple Watch Ultra", r"Apple Watch Series \d+", r"Galaxy Watch ?\d+", r"Pixel Watch ?\d*",
    r"AirPods Pro ?\d*", r"AirPods ?\d*", r"Galaxy Buds ?\w*",
    r"Model Y L", r"Model Y", r"Model 3", r"Model [SX]", r"Cybertruck",
]
NO_DEVICE = {"", "common", "알 수 없음_추가정보 필요", "해당없음", "-"}


def norm_model(s):
    s = (s or "").lower().replace("mag fit", "magfit").replace("glas.tr", "glastr").replace("glas tr", "glastr")
    return re.sub(r"[^0-9a-z가-힣]", "", s)


def series_of(device):
    d = (device or "").strip().lstrip("#")
    first = d.split("/")[0].strip()
    if first.lower() in NO_DEVICE:
        return ""
    for p in SERIES_PATTERNS:
        m = re.search(p, first, re.I)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).replace("Fold ", "Fold").replace("Flip ", "Flip")
    return first


class ProductMaster:
    def __init__(self, rows):
        h = rows[0]
        ix = {k: i for i, k in enumerate(h)}
        g = lambda r, k: (r[ix[k]] if ix[k] < len(r) else "").strip()
        self.by_sku, self.by_asin, self.device_alias = {}, {}, {}
        for r in rows[1:]:
            sku, asin = g(r, "SKU")[:8].upper(), g(r, "ASIN")
            rec = {"sku": sku, "asin": asin, "device": g(r, "기종명"), "model": g(r, "모델명"),
                   "maker": g(r, "생산업체"), "cat": g(r, "대분류")}
            if sku and sku not in self.by_sku:
                self.by_sku[sku] = rec
            if asin.startswith("B0") and asin not in self.by_asin:
                self.by_asin[asin] = rec
            # device alias: every compatible device in '기종명' -> series of the first one
            parts = [p.strip() for p in rec["device"].lstrip("#").split("/") if p.strip()]
            if parts:
                fam = re.match(r"^(iPhone|Galaxy S|Galaxy Z|Galaxy|Pixel|Apple Watch)", parts[0])
                ser = series_of(parts[0])
                for p in parts:
                    full = p if (not fam or p.startswith(fam.group(1).split()[0])) else f"{fam.group(1)} {p}"
                    self.device_alias.setdefault(full.lower(), ser)

    def lineup_from_sku(self, sku):
        rec = self.by_sku.get((sku or "")[:8].upper())
        if not rec or not rec["model"]:
            return None
        return (series_of(rec["device"]), norm_model(rec["model"]))

    def lineup_from_fields(self, device, product):
        if not product or product.lower() in NO_DEVICE:
            return None
        ser = self.device_alias.get((device or "").strip().lower(), series_of(device))
        return (ser, norm_model(product))
