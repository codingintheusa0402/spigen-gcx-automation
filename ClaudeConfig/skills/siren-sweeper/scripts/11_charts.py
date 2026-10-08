"""Step 11 — Apple-style charts (same look as Bi-Weekly tools/apple_theme/chart.py): m<deck>.png monthly trend per case
(gray rounded-top bars, latest month blue, dashed trend) and overview.png (one rounded shape per bar, identical
corner radius, claims clipped inside; x labels = full product name + defect row, NO #N — user rule 2026-10-02).
Run with /usr/bin/python3 (matplotlib). Usage: /usr/bin/python3 11_charts.py"""
import common  # noqa
import json, math
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import PathPatch
from matplotlib.path import Path
for f in ['/Library/Fonts/SF-Pro-Text-Regular.otf', '/Library/Fonts/SF-Pro-Text-Semibold.otf', '/System/Library/Fonts/AppleSDGothicNeo.ttc']:
    try: fm.fontManager.addfont(f)
    except Exception: pass
plt.rcParams['font.family'] = ['SF Pro Text', 'Apple SD Gothic Neo']
INK, GRAY, GRAY2, BAR, BLUE, LBLUE, GRID = '#1D1D1F', '#6E6E73', '#86868B', '#D2D2D7', '#0071E3', '#64D2FF', '#EDEDF0'

def rounded(ax, fig, x0, w, y0, h, n, ymax, xmin, xmax, r_frac=0.30):
    bbox = ax.get_position(); fw, fh = fig.get_size_inches() * fig.dpi
    px_x = bbox.width * fw / (xmax - xmin); px_y = bbox.height * fh / ymax
    rx = w * r_frac; ry = rx * px_x / px_y; ry = min(ry, h); rx = min(rx, ry * px_y / px_x)
    top = y0 + h
    pts = [(x0, y0), (x0, top - ry)]
    for k in range(13):
        a = math.pi - k * (math.pi / 2) / 12; pts.append((x0 + rx + rx * math.cos(a), top - ry + ry * math.sin(a)))
    for k in range(13):
        a = math.pi / 2 - k * (math.pi / 2) / 12; pts.append((x0 + w - rx + rx * math.cos(a), top - ry + ry * math.sin(a)))
    pts += [(x0 + w, y0), (x0, y0)]
    return Path(pts, [Path.MOVETO] + [Path.LINETO] * (len(pts) - 2) + [Path.CLOSEPOLY])

def monthly(c, path, W=300, H=118):
    import datetime
    Y = datetime.date.today().year
    months = [f'{Y}-{m:02d}' for m in range(1, 13)]
    vals = [c['months'].get(m, 0) for m in months]
    while len(vals) > 1 and vals[-1] == 0 and months[-1] > c['last'][:7]: vals.pop(); months.pop()
    n = len(vals)
    fig = plt.figure(figsize=(W/72, H/72), dpi=300)
    ax = fig.add_axes([0.03, 0.2, 0.94, 0.66])
    ymax = max(max(vals) * 1.25, 1); xmin, xmax = -0.6, n - 0.4
    ax.set_xlim(xmin, xmax); ax.set_ylim(0, ymax); bw = 0.6
    for i, v in enumerate(vals):
        col = BLUE if i == n - 1 else BAR
        if v: ax.add_patch(PathPatch(rounded(ax, fig, i - bw/2, bw, 0, v, n, ymax, xmin, xmax), fc=col, ec='none', zorder=2))
        ax.text(i, v + ymax * 0.03, str(v), ha='center', va='bottom', fontsize=7, color=BLUE if i == n - 1 else GRAY,
                fontweight='semibold' if i == n - 1 else 'normal')
        ax.text(i, -ymax * 0.07, f'{int(months[i][5:])}월', ha='center', va='top', fontsize=6.5, color=BLUE if i == n - 1 else GRAY)
    if n > 2:
        xm = (n-1)/2; ym = sum(vals)/n; b = sum((i-xm)*(v-ym) for i, v in enumerate(vals)) / sum((i-xm)**2 for i in range(n)); a = ym - b*xm
        ax.plot([-0.4, n-0.6], [a + b*(-0.4), a + b*(n-0.6)], color=BLUE, lw=0.7, alpha=0.35, dashes=(3, 2.2))
    ax.axhline(0, color=BAR, lw=0.6)
    for s in ax.spines.values(): s.set_visible(False)
    ax.set_xticks([]); ax.set_yticks([])
    fig.savefig(path, facecolor='white'); plt.close(fig)

def overview(cases, path, W=620, H=215):
    """Every bar = ONE rounded-top shape of the same corner radius (full total, light blue); the claims part
    is drawn inside it and clipped to that shape, so thin review segments never change the corner."""
    n = len(cases)
    fig = plt.figure(figsize=(W/72, H/72), dpi=300)
    ax = fig.add_axes([0.02, 0.33, 0.96, 0.61])
    ymax = max(c['total'] for c in cases) * 1.22; xmin, xmax = -0.6, n - 0.4
    ax.set_xlim(xmin, xmax); ax.set_ylim(0, ymax); bw = 0.56
    bbox = ax.get_position(); fw, fh = fig.get_size_inches() * fig.dpi
    px_x = bbox.width * fw / (xmax - xmin); px_y = bbox.height * fh / ymax
    rx = bw * 0.30; ry = rx * px_x / px_y            # identical corner for every bar
    def shape(x0, h):
        r_y = min(ry, h); r_x = rx * (r_y / ry)
        pts = [(x0, 0), (x0, h - r_y)]
        for k in range(13):
            a = math.pi - k * (math.pi / 2) / 12; pts.append((x0 + r_x + r_x * math.cos(a), h - r_y + r_y * math.sin(a)))
        for k in range(13):
            a = math.pi / 2 - k * (math.pi / 2) / 12; pts.append((x0 + bw - r_x + r_x * math.cos(a), h - r_y + r_y * math.sin(a)))
        pts += [(x0 + bw, 0), (x0, 0)]
        return Path(pts, [Path.MOVETO] + [Path.LINETO] * (len(pts) - 2) + [Path.CLOSEPOLY])
    for i, c in enumerate(cases):
        outer = PathPatch(shape(i - bw/2, c['total']), fc=LBLUE, ec='none', zorder=2)
        ax.add_patch(outer)
        inner = plt.Rectangle((i - bw/2, 0), bw, c['claims'], fc=BLUE, ec='none', zorder=3)
        ax.add_patch(inner); inner.set_clip_path(outer)
        ax.text(i, c['total'] + ymax * 0.025, str(c['total']), ha='center', va='bottom', fontsize=7.5, color=INK, fontweight='semibold')
        import textwrap
        name = c['name'].replace(' 시리즈용 ', ' ')
        lines = textwrap.wrap(name, 11, break_long_words=False)
        ax.annotate('\n'.join(lines), xy=(i, 0), xytext=(0, -5), textcoords='offset points', ha='center', va='top',
                    fontsize=6.6, color=INK, linespacing=1.12, annotation_clip=False)
        ax.annotate(c['defect'], xy=(i, 0), xytext=(0, -5 - 4 * 7.8 - 3), textcoords='offset points',
                    ha='center', va='top', fontsize=6.6, color=GRAY, annotation_clip=False)
    ax.axhline(0, color=BAR, lw=0.6)
    for s_ in ax.spines.values(): s_.set_visible(False)
    ax.set_xticks([]); ax.set_yticks([])
    fig.savefig(path, facecolor='white'); plt.close(fig)

cases = json.load(open('cases.json'))
for c in cases: monthly(c, f"m{c['n']:02d}.png")
overview(cases, 'overview.png')
print('ok')
