# Make watermark as transparent PNG using PIL (raster = immune to viewer font bugs)
from PIL import Image, ImageDraw, ImageFont

TEXT = "版权所有 © 2026 改变从整理开始"
SCALE = 20            # render at 20x for crispness
PT = 7 * SCALE        # 7pt font at 20x
GRAY = 140            # 55% gray, same as reference

candidates = [
    '/tmp/PingFangSC-Regular.ttf',
    '/System/Library/Fonts/PingFang.ttc',
    '/System/Library/Fonts/Supplemental/Songti.ttc',
]
font = None
for path in candidates:
    try:
        font = ImageFont.truetype(path, PT)
        print('using font:', path)
        break
    except Exception as e:
        print('fail', path, e)
assert font

# measure
tmp = Image.new('RGBA', (10, 10))
dd = ImageDraw.Draw(tmp)
bbox = dd.textbbox((0, 0), TEXT, font=font)
w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
print('ink size at 1x:', w / SCALE, h / SCALE, 'bbox', bbox)

img = Image.new('RGBA', (w + 40, h + 40), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.text((20 - bbox[0], 20 - bbox[1]), TEXT, font=font, fill=(GRAY, GRAY, GRAY, 255))
img.save('/tmp/watermark.png')
print('saved /tmp/watermark.png', img.size)
