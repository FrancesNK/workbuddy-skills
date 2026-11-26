import os, sys, json, time, urllib.request, glob, re
import pymupdf

os.environ['TDOC_API_BASE_URL'] = 'https://docs.qq.com'
_paths = glob.glob('/Users/yufengs/.workbuddy/plugins/cache/workbuddy-builtin/tencent-docs-plugin/*/skills/tencent-docs')
def _key(p):
    m = re.search(r'/(\d+)\.(\d+)\.(\d+)-', p)
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)
sys.path.insert(0, max(_paths, key=_key))
import tencentdocs as td

FID = 'DYm9Cb0tPZWV4ZW5C'
PAGES = list(range(16))
N = len(PAGES)

# ---- 尺寸（pt）----
MM = 72 / 25.4
PW, PH = 176 * MM, 250 * MM          # B5 成品页 498.90 x 708.66
SW, SH = 420 * MM, 297 * MM          # A3 横向印张 1190.55 x 841.89
PAGE_X = (SW - 2 * PW) / 2           # 96.38
PAGE_Y = (SH - PH) / 2               # 66.61
FOLD = SW / 2
K = PW / 421.0                       # A5(421x595pt) -> B5 的缩放比 ≈1.185

LONG_EDGE = int(round(PH / 72 * 300))   # 300dpi

# 1. 导出页图
urls = {}
for start in range(0, N, 4):
    batch = PAGES[start:start + 4]
    for attempt in range(4):
        res, err = td.call_tool('slide-mcp', 'slide_export_pages_to_image_urls',
                                {"file_id": FID, "long_edge": LONG_EDGE, "page_indices": batch})
        if err:
            print('err', batch, err); time.sleep(5); continue
        try:
            obj = json.loads(res['result']['content'][0]['text'])
            for it in obj['images']:
                urls[it['page_index']] = it['image_url']
            print('ok', batch, 'total', len(urls))
            break
        except Exception as e:
            print('parse err', batch, e); time.sleep(5)
assert len(urls) == N, f'{len(urls)}/{N}'

# 2. 下载
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/126 Safari/537.36',
           'Referer': 'https://docs.qq.com/'}
files = []
for seq, idx in enumerate(PAGES, 1):
    out = f'/tmp/b5_{seq:02d}.png'
    for attempt in range(4):
        try:
            req = urllib.request.Request(urls[idx], headers=headers)
            open(out, 'wb').write(urllib.request.urlopen(req, timeout=90).read())
            files.append(out)
            break
        except Exception as e:
            print('retry', idx, attempt, repr(e)); time.sleep(3)
assert len(files) == N, f'{len(files)}/{N}'

# 3. 拼版
GRAY = (0.55, 0.55, 0.55)
WM_W, WM_H = 2169 / 20 * K, 172 / 20 * K
WM_X, WM_Y = FOLD - WM_W / 2, PAGE_Y + PH - (6.6 + 172 / 20) * K

doc = pymupdf.open()
for i in range(1, N // 2 + 1):
    left, right = (N + 1 - i, i) if i % 2 else (i, N + 1 - i)
    page = doc.new_page(width=SW, height=SH)
    page.insert_image(pymupdf.Rect(PAGE_X, PAGE_Y, PAGE_X + PW, PAGE_Y + PH), filename=files[left - 1])
    page.insert_image(pymupdf.Rect(FOLD, PAGE_Y, FOLD + PW, PAGE_Y + PH), filename=files[right - 1])
    y = PAGE_Y + PH - 8.5 * K
    page.draw_line(pymupdf.Point(PAGE_X + 24 * K, y), pymupdf.Point(FOLD - 61.17 * K, y), color=GRAY, width=0.5)
    page.draw_line(pymupdf.Point(FOLD + 61.17 * K, y), pymupdf.Point(FOLD + PW - 24 * K, y), color=GRAY, width=0.5)
    page.insert_image(pymupdf.Rect(WM_X, WM_Y, WM_X + WM_W, WM_Y + WM_H),
                      filename='/tmp/watermark.png', overlay=True)
    print(f'sheet {i}: left=p{left} right=p{right}')

perm = pymupdf.PDF_PERM_ACCESSIBILITY | pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ
out = '/tmp/december_booklet_b5.pdf'
doc.save(out, garbage=4, deflate=True,
         encryption=pymupdf.PDF_ENCRYPT_AES_256,
         owner_pw='<OWNER_PW>', user_pw='', permissions=perm)
doc.close()
print('SAVED', out)

d = pymupdf.open(out); d.authenticate('<OWNER_PW>')
print('pages:', d.page_count, '| sheet pt:', round(d[0].rect.width, 2), round(d[0].rect.height, 2))
print('imgs per page:', [len(p.get_images(full=True)) for p in d])
print('fonts:', [f for p in d for f in p.get_fonts(full=True)] or 'NONE')
for i in range(d.page_count):
    d[i].get_pixmap(dpi=60).save(f'/tmp/b5_s{i+1}.png')
print('renders ok')
