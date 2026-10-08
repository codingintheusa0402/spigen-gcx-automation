import json
"""Shared settings for siren-sweeper scripts. Every script `import common` first: it chdirs into the
run's work folder (all intermediate files live there, flat) so the steps can be re-run independently."""
import os, sys, datetime
SKILL = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get('SWEEP_DIR') or os.path.expanduser(f"~/.config/siren_sweeper/{datetime.date.today():%y%m%d}")
os.makedirs(WORK, exist_ok=True)
os.chdir(WORK)
sys.path.insert(0, '/Users/kevinkim/.claude/skills/siren-finder')
sys.path.insert(0, '/Users/kevinkim/.claude/skills/siren-report')
sys.path.insert(0, '/Users/kevinkim/Desktop/GCX/GAS_Operations/Bi-Weekly/tools')
sys.path.insert(0, '/Users/kevinkim/Desktop/GCX/GAS_Operations/Bi-Weekly/tools/apple_theme')

# user's PRIVATE room (spaces/AAQAc9NQmJQ) — the only room this skill posts to unless told otherwise
PRIVATE_WEBHOOK = json.load(open(os.path.expanduser("~/.config/gcx_webhooks.json")))["siren_private_room"]  # secret — kept out of git
CANDIDATES_DIR = os.path.expanduser('~/.config/siren_finder/runs')

# SIREN registries checked for 기등록 (besides siren-finder's CQ Emergency Net check)
GCX_KPI_26 = ('15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI', '26년 SIREN', 19)            # header row 18
GCX_KPI_25 = '1o49Ji8AMmN96uGcK6h2kdEtvoVzhUXDSDbr3V88VL8Q'                                # tabs 상반기/하반기 KPI 실적 검증
SALES_SIREN = ('1OsJ5H7yOrK16VfWVaxXnty8uK8tWzJRnDBrc8U6INFM', '영업', 4)                   # [🚨 SIREN 🚨] 영업 tab, header row 3
SKU_MASTER = '1JijzoYw9aDW-9Jx_OJOBByMVCX8-0OtNxb-0Y82EZJI'                                # SKU_Master(먼데이보드) Data: 생산업체/원산지
PRODUCT_MASTER = '1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo'                            # Data: 생산업체/원산지정보 (fallback)
GOLDEN_APPLE_DECK = '1quCr9Xj-pSsVXKrYuaEOq0LPILZMN2LPUwBkY1f_GFI'                          # approved 261002 Apple Bi-weekly deck
EU_CHANNELS = ['Amazon.de', 'Amazon.fr', 'Amazon.it', 'Amazon.es', 'Amazon.nl', 'Amazon.se', 'Amazon.com.be', 'Amazon.ie', 'Amazon.pl', 'Amazon.co.uk']
EU_COUNTRIES = {'DE', 'FR', 'IT', 'ES', 'NL', 'SE', 'BE', 'IE', 'PL', 'UK', 'GB'}

def half_label(d=None):
    d = d or datetime.date.today()
    return '상반기' if d.month <= 6 else '하반기'

def latest_candidates():
    fs = sorted(f for f in os.listdir(CANDIDATES_DIR) if f.startswith('candidates_'))
    return os.path.join(CANDIDATES_DIR, fs[-1])
