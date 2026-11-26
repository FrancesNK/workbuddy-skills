import pymupdf

W, H, HALF, N = 841.8898, 595.2756, 420.9449, 16
GRAY = (0.55, 0.55, 0.55)

# watermark PNG: 2169x172 @20x -> 108.45 x 8.6 pt; ink-left offset 1.0pt, baseline 9.2pt from img top
WM_W, WM_H = 2169 / 20, 172 / 20
WM_X = 367.7698 - 1.0          # ink left aligns to 367.77 like reference
WM_Y = H - 6 - 9.2             # baseline at H-6

def compose(prefix, out):
    doc = pymupdf.open()
    for i in range(1, N // 2 + 1):
        left, right = (N + 1 - i, i) if i % 2 else (i, N + 1 - i)
        page = doc.new_page(width=W, height=H)
        page.insert_image(pymupdf.Rect(0, 0, HALF, H), filename=f'/tmp/{prefix}_slide_{left-1:02d}.png')
        page.insert_image(pymupdf.Rect(HALF, 0, W, H), filename=f'/tmp/{prefix}_slide_{right-1:02d}.png')
        y = H - 8.5
        page.draw_line(pymupdf.Point(24, y), pymupdf.Point(359.7698, y), color=GRAY, width=0.5)
        page.draw_line(pymupdf.Point(482.1199, y), pymupdf.Point(817.8898, y), color=GRAY, width=0.5)
        page.insert_image(pymupdf.Rect(WM_X, WM_Y, WM_X + WM_W, WM_Y + WM_H),
                          filename='/tmp/watermark.png', overlay=True)
    perm = pymupdf.PDF_PERM_ACCESSIBILITY | pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ
    doc.save(out, garbage=4, deflate=True,
             encryption=pymupdf.PDF_ENCRYPT_AES_256,
             owner_pw='<OWNER_PW>', user_pw='', permissions=perm)
    doc.close()
    print('SAVED', out)

compose('oct', '/tmp/october_booklet_v4.pdf')
compose('nov', '/tmp/november_booklet_v4.pdf')
