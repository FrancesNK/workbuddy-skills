import pymupdf

MM = 72 / 25.4
SW, SH = 210 * MM, 297 * MM          # A4 纵向印张 595.28 x 841.89
PW, PH = 105 * MM, 148.5 * MM        # A6 成品页  297.64 x 420.94
K = PW / 421.0                       # A5(421pt) -> A6 缩放比 ≈0.707

GRAY = (0.55, 0.55, 0.55)
WM_W, WM_H = 2169 / 20 * K, 172 / 20 * K

# 每个 PDF 页 = 一张 A4 的一个印面：上半幅 2 页 + 下半幅 2 页
# 顺序：sheet1 正面 / sheet1 反面 / sheet2 正面 / sheet2 反面
SIDES = [
    [(16, 1), (14, 3)],   # A4 第1张 正面：上半幅=最外层折手，下半幅=第2层
    [(2, 15), (4, 13)],   # A4 第1张 反面
    [(12, 5), (10, 7)],   # A4 第2张 正面
    [(6, 11), (8, 9)],    # A4 第2张 反面
]

COL = [0, PW]
ROW = [0, PH]


def draw_half(page, x_off, y_off, left_page, right_page):
    page.insert_image(pymupdf.Rect(x_off, y_off, x_off + PW, y_off + PH),
                      filename=f'/tmp/dec_{left_page:02d}.png')
    page.insert_image(pymupdf.Rect(x_off + PW, y_off, x_off + 2 * PW, y_off + PH),
                      filename=f'/tmp/dec_{right_page:02d}.png')
    fold = x_off + PW
    half_r = x_off + 2 * PW
    bottom = y_off + PH
    y = bottom - 8.5 * K
    page.draw_line(pymupdf.Point(x_off + 24 * K, y), pymupdf.Point(fold - 61.17 * K, y),
                   color=GRAY, width=0.5)
    page.draw_line(pymupdf.Point(fold + 61.17 * K, y), pymupdf.Point(half_r - 24 * K, y),
                   color=GRAY, width=0.5)
    wm_y = bottom - (6.6 + 8.6) * K
    page.insert_image(pymupdf.Rect(fold - WM_W / 2, wm_y, fold + WM_W / 2, wm_y + WM_H),
                      filename='/tmp/watermark.png', overlay=True)


def scissors(page, x, y, s, color, mirror=False):
    """在 (x, y) 画一个小剪刀，box 宽高约 s；mirror=True 时刃口朝左。"""
    def X(v):
        return x + s - v if mirror else x + v
    c1 = pymupdf.Point(X(0.20 * s), y - 0.17 * s)
    c2 = pymupdf.Point(X(0.20 * s), y + 0.17 * s)
    page.draw_circle(c1, 0.11 * s, color=color, width=0.5)
    page.draw_circle(c2, 0.11 * s, color=color, width=0.5)
    page.draw_line(c1, pymupdf.Point(X(0.95 * s), y + 0.13 * s), color=color, width=0.5)
    page.draw_line(c2, pymupdf.Point(X(0.95 * s), y - 0.13 * s), color=color, width=0.5)


doc = pymupdf.open()
for side in SIDES:                       # [(top_l, top_r), (bottom_l, bottom_r)]
    page = doc.new_page(width=SW, height=SH)
    (tl, tr), (bl, br) = side
    draw_half(page, COL[0], ROW[0], tl, tr)     # 上半幅
    draw_half(page, COL[0], ROW[1], bl, br)     # 下半幅
    # 裁剪线：A4 横向正中（两张 A5 的分界），淡灰虚线 + 两端小剪刀
    y_cut = PH
    SC = 11
    page.draw_line(pymupdf.Point(SC + 7, y_cut), pymupdf.Point(SW - SC - 7, y_cut),
                   color=(0.72, 0.72, 0.72), width=0.4, dashes="[3 3] 0")
    scissors(page, 4, y_cut, SC, (0.58, 0.58, 0.58))
    scissors(page, SW - 4 - SC, y_cut, SC, (0.58, 0.58, 0.58), mirror=True)
    print('side ->', 'top', (tl, tr), '| bottom', (bl, br))

perm = pymupdf.PDF_PERM_ACCESSIBILITY | pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ
out = '/tmp/december_booklet_a6.pdf'
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
    d[i].get_pixmap(dpi=110).save(f'/tmp/a6_s{i+1}.png')
print('renders ok')
