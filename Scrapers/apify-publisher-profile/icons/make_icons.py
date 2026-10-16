import math
from PIL import Image, ImageDraw, ImageFilter
N=1024
NAVY=(16,32,74,255); GOLD=(255,190,30,255); W=(255,255,255,245)
def tile():
    img=Image.new('RGBA',(N,N),(0,0,0,0)); bg=Image.new('RGBA',(N,N)); px=bg.load()
    top=(16,32,74); bot=(52,64,180)
    for y in range(N):
        t=y/(N-1); c=tuple(int(top[i]+(bot[i]-top[i])*t) for i in range(3))+(255,)
        for x in range(N): px[x,y]=c
    m=Image.new('L',(N,N),0); ImageDraw.Draw(m).rounded_rectangle([0,0,N-1,N-1],radius=230,fill=255); img.paste(bg,(0,0),m); return img,m
def star(cx,cy,R,r=None):
    r=r or R*0.45
    return [(cx+(R if k%2==0 else r)*math.cos(-math.pi/2+k*math.pi/5),cy+(R if k%2==0 else r)*math.sin(-math.pi/2+k*math.pi/5)) for k in range(10)]
def glow(img,m,shape_fn,color,blur=36):
    g=Image.new('RGBA',(N,N),(0,0,0,0)); shape_fn(ImageDraw.Draw(g),color); g=g.filter(ImageFilter.GaussianBlur(blur))
    g.putalpha(Image.composite(g.getchannel('A'),Image.new('L',(N,N),0),m)); return Image.alpha_composite(img,g)

def brand(img):
    d=ImageDraw.Draw(img); cx,cy,R=165,165,100
    d.ellipse([cx-R,cy-R,cx+R,cy+R],fill=(255,255,255,255))
    d.rounded_rectangle([cx-50,cy-58,cx-22,cy+58],radius=12,fill=NAVY)                 # K stem
    def arm(x0,y0,x1,y1,col,w=28):
        d.line([(x0,y0),(x1,y1)],fill=col,width=w)
        for x,y in ((x0,y0),(x1,y1)): d.ellipse([x-w//2,y-w//2,x+w//2,y+w//2],fill=col)
    arm(cx-10,cy+2,cx+40,cy-46,NAVY)                                                  # upper arm
    arm(cx-4,cy+10,cx+42,cy+48,GOLD)                                                  # lower arm (gold = data)
    return img

# 1) ASIN rating: product tag with barcode (ASIN) + big gold star + count bars
img,m=tile(); d=ImageDraw.Draw(img)
# price tag (rotated rectangle w/ notch) drawn on separate layer then rotated
tag=Image.new('RGBA',(N,N),(0,0,0,0)); td=ImageDraw.Draw(tag)
td.rounded_rectangle([230,330,760,760],radius=60,fill=W)
td.polygon([(230,420),(130,545),(230,670)],fill=W)          # pointed end
td.ellipse([170,515,230,575],fill=(36,50,130,255))           # hole
for i,(x,w) in enumerate([(300,22),(338,10),(364,30),(410,12),(438,22),(476,10),(502,34),(552,14),(582,22),(620,10),(646,30),(692,14)]):
    td.rectangle([x,420,x+w,600],fill=NAVY)                  # barcode = ASIN
td.rounded_rectangle([300,640,700,690],radius=20,fill=(52,64,180,170))
tag=tag.rotate(-14,resample=Image.BICUBIC,center=(500,560)); img=Image.alpha_composite(img,tag)
img=glow(img,m,lambda g,c: g.polygon(star(735,300,215),fill=c),(255,196,40,150))
d=ImageDraw.Draw(img); d.polygon(star(735,300,200),fill=GOLD); d.polygon(star(728,288,122),fill=(255,214,90,255))
img=brand(img)
o1=img.resize((512,512),Image.LANCZOS); o1.save('/Users/kevinkim/Desktop/GCX/Scrapers/apify-asin-rating-scraper/icon.png')

# 2) Seller Central reviews: dashboard window + review cards + gold key
img,m=tile(); d=ImageDraw.Draw(img)
X0,Y0,X1,Y1=290,230,910,790
d.rounded_rectangle([X0,Y0,X1,Y1],radius=56,fill=W)
d.rounded_rectangle([X0,Y0,X1,Y0+100],radius=56,fill=(220,226,245,255)); d.rectangle([X0,Y0+56,X1,Y0+100],fill=(220,226,245,255))
for i,c in enumerate([(255,95,87,255),(255,189,46,255),(39,201,63,255)]): d.ellipse([X0+50+i*56,Y0+32,X0+86+i*56,Y0+68],fill=c)
d.rounded_rectangle([X0+250,Y0+33,X1-40,Y0+67],radius=17,fill=(255,255,255,255))
for row,y in enumerate([Y0+140,Y0+340]):
    d.rounded_rectangle([X0+45,y,X1-45,y+170],radius=32,fill=(237,240,250,255))
    for k in range(5):
        d.polygon(star(X0+100+k*64,y+52,28),fill=GOLD if k < (4 if row==0 else 2) else (200,206,225,255))
    d.rounded_rectangle([X0+75,y+100,X1-150,y+124],radius=12,fill=(52,64,180,150))
    d.rounded_rectangle([X0+75,y+136,X0+380,y+154],radius=9,fill=(52,64,180,90))
cx,cy=820,820
img=glow(img,m,lambda g,c: g.ellipse([cx-160,cy-160,cx+160,cy+160],fill=c),(255,196,40,140))
d=ImageDraw.Draw(img); d.ellipse([cx-135,cy-135,cx+135,cy+135],fill=GOLD,outline=NAVY,width=14)
d.ellipse([cx-82,cy-88,cx+8,cy+2],outline=NAVY,width=30); d.line([(cx-8,cy-14),(cx+78,cy+72)],fill=NAVY,width=30)
d.line([(cx+40,cy+34),(cx+68,cy+6)],fill=NAVY,width=24); d.line([(cx+68,cy+62),(cx+96,cy+34)],fill=NAVY,width=24)
img=brand(img)
o2=img.resize((512,512),Image.LANCZOS); o2.save('/Users/kevinkim/Desktop/GCX/Scrapers/apify-sc-review-scraper/icon.png')

prev=Image.new('RGBA',(760,300),(245,245,247,255))
for i,o in enumerate((o1,o2)):
    a=o.resize((220,220),Image.LANCZOS); prev.paste(a,(20+i*370,20),a)
    s=o.resize((56,56),Image.LANCZOS); prev.paste(s,(255+i*370,100),s)
    s2=o.resize((32,32),Image.LANCZOS); prev.paste(s2,(268+i*370,180),s2)
prev.save('/private/tmp/claude-501/-Users-kevinkim/a5ed7702-2ab1-4ed4-a7ba-c78b8d752db1/scratchpad/icons3_prev.png'); print('ok')
