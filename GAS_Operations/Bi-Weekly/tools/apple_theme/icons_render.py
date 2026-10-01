from PIL import Image, ImageDraw, ImageFont
import json, base64, io
cp = {}
for line in open('MSR.codepoints'):
    n, c = line.split(); cp[n] = int(c, 16)
ICONS = {  # name -> (glyph, color)
  'globe': ('public', '#0071E3'), 'report': ('report', '#0071E3'), 'code': ('qr_code', '#0071E3'),
  'box': ('inventory_2', '#0071E3'), 'star': ('star', '#0071E3'), 'bag': ('shopping_bag', '#86868B'),
  'agent': ('support_agent', '#86868B'), 'insights': ('insights', '#0071E3'), 'fold': ('devices_fold', '#0071E3'),
  'phone': ('smartphone', '#0071E3'), 'iphone': ('phone_iphone', '#0071E3'), 'siren': ('emergency_home', '#FF3B30'),
  'category': ('category', '#0071E3'), 'donut': ('donut_large', '#0071E3'), 'photos': ('photo_library', '#0071E3'),
  'sheet': ('table_chart', '#0071E3'), 'db': ('database', '#0071E3'), 'doc': ('description', '#0071E3'),
}
font = ImageFont.truetype('MSR.ttf', 176)
out = {}
for key, (g, col) in ICONS.items():
    im = Image.new('RGBA', (192, 192), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    rgb = tuple(int(col[i:i+2], 16) for i in (1, 3, 5))
    d.text((96, 96), chr(cp[g]), font=font, fill=rgb + (255,), anchor='mm')
    bb = im.getbbox(); print(key, bb)
    buf = io.BytesIO(); im.save(buf, 'PNG', optimize=True); out[key] = base64.b64encode(buf.getvalue()).decode()
    im.save(f'{key}.png')
json.dump(out, open('icons_b64.json', 'w'))
print({k: len(v) for k, v in out.items()})
