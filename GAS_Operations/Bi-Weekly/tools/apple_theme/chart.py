import json, re
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import PathPatch
from matplotlib.path import Path
import math
S = '/private/tmp/claude-501/-Users-kevinkim/8ab0abda-73a1-40e3-934d-c473628fe948/scratchpad/'
for f in ['/Library/Fonts/SF-Pro-Text-Regular.otf', '/Library/Fonts/SF-Pro-Text-Semibold.otf', '/Library/Fonts/SF-Pro-Display-Regular.otf', '/System/Library/Fonts/AppleSDGothicNeo.ttc']:
    try: fm.fontManager.addfont(f)
    except Exception as ex: print('font', f, ex)
plt.rcParams['font.family'] = ['SF Pro Text', 'Apple SD Gothic Neo']
INK, GRAY, GRAY2, BAR, BLUE, GRID = '#1D1D1F', '#6E6E73', '#86868B', '#D2D2D7', '#0071E3', '#EDEDF0'

weeks = json.load(open(S + 'weekly.json'))['weeks']
def key(lbl):
    m = re.match(r'(\d+)월 (\d+)주 (\d+)년', lbl); return (int(m.group(3)), int(m.group(1)), int(m.group(2)))
data = sorted([(key(l), l, n) for l, n, _ in weeks])
data = [d for d in data if d[0] >= (2026, 3, 5)]          # same window start as the deck's chart
labels = [d[1] for d in data]; vals = [d[2] for d in data]; n = len(vals)

W, H = 545, 257                                             # slide box in pt
fig = plt.figure(figsize=(W/72, H/72), dpi=300)
ax = fig.add_axes([0.035, 0.17, 0.93, 0.62])
ymax = max(vals) * 1.18
ax.set_xlim(-0.6, n - 0.4); ax.set_ylim(0, ymax)
bw = 0.62
# pixels per data unit -> circular (not elliptical) corners on screen
bbox = ax.get_position(); fw, fh = fig.get_size_inches() * fig.dpi
px_x = bbox.width * fw / (n - 0.4 + 0.6); px_y = bbox.height * fh / ymax
def top_rounded_bar(x0, w, h, r_frac=0.30):
    rx = w * r_frac; ry = rx * px_x / px_y
    ry = min(ry, h); rx = min(rx, ry * px_y / px_x)
    pts = [(x0, 0), (x0, h - ry)]
    for k in range(0, 13):                     # top-left quarter circle
        a = math.pi - k * (math.pi / 2) / 12
        pts.append((x0 + rx + rx * math.cos(a), h - ry + ry * math.sin(a)))
    for k in range(0, 13):                     # top-right quarter circle
        a = math.pi / 2 - k * (math.pi / 2) / 12
        pts.append((x0 + w - rx + rx * math.cos(a), h - ry + ry * math.sin(a)))
    pts += [(x0 + w, 0), (x0, 0)]
    return Path(pts, [Path.MOVETO] + [Path.LINETO] * (len(pts) - 2) + [Path.CLOSEPOLY])
for i, v in enumerate(vals):
    c = BLUE if i == n - 1 else BAR
    ax.add_patch(PathPatch(top_rounded_bar(i - bw/2, bw, v), fc=c, ec='none', zorder=2))
    ax.text(i, v + ymax*0.018, f'{v:,}', ha='center', va='bottom', fontsize=4.6, color=BLUE if i == n - 1 else GRAY2,
            fontweight='semibold' if i == n - 1 else 'normal')
# trend (least squares)
xm = (n-1)/2; ym = sum(vals)/n
b = sum((i-xm)*(v-ym) for i, v in enumerate(vals)) / sum((i-xm)**2 for i in range(n)); a = ym - b*xm
ax.plot([-0.4, n-0.6], [a + b*(-0.4), a + b*(n-0.6)], color=BLUE, lw=0.7, alpha=0.35, dash_capstyle='round', dashes=(3, 2.2))
for y in (250, 500, 750, 1000):
    if y < ymax: ax.axhline(y, color=GRID, lw=0.5, zorder=0); ax.text(n - 0.35, y, f'{y:,}', fontsize=4, color=GRAY2, va='center', ha='left')
ax.axhline(0, color=BAR, lw=0.6)
for s in ax.spines.values(): s.set_visible(False)
ax.set_xticks([]); ax.set_yticks([])
prev_m = None
for i, l in enumerate(labels):
    m = re.match(r'(\d+)월 (\d+)주', l)
    ax.text(i, -ymax*0.05, f'{m.group(2)}주', ha='center', va='top', fontsize=4.4, color=BLUE if i == n-1 else GRAY)
    if m.group(1) != prev_m:
        ax.text(i - bw/2, -ymax*0.15, f'{m.group(1)}월', ha='left', va='top', fontsize=5, color=INK, fontweight='semibold')
        prev_m = m.group(1)
fig.savefig(S + 'weekly_chart.png', transparent=False, facecolor='white')
print(n, labels[0], labels[-1], vals[-3:])
