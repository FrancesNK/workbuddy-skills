"""
Generate saddle-stitch booklet PDF using reportlab + PIL (high-res PNG source).
Stable and reliable — no PDF merge issues.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image

# Register Chinese font for copyright watermark
FONT_PATH = 'C:/Windows/Fonts/msyh.ttc'
FONT_NAME = 'MSYH'
try:
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
    print(f'Registered font: {FONT_NAME}')
except Exception as e:
    print(f'WARNING: Could not register font {FONT_PATH}: {e}')
    FONT_NAME = 'Helvetica'  # fallback

SRC_DIR = 'C:/Users/Jihai/WorkBuddy/2026-07-12-20-49-17/work/slides_new'
OUT = 'C:/Users/Jihai/WorkBuddy/2026-07-12-20-49-17/手账打印版_骑马钉装订.pdf'

# A4 landscape in points: 842 x 595 pt (297 x 210 mm)
# Two A5 pages side by side, each ~421 x 595 pt
A4_W, A4_H = landscape(A4)  # 841.89, 595.28

# Saddle-stitch pairing for 16 pages (0-indexed)
page_pairs = [
    (15, 0),   # Sheet 1 front: left=16(封底) right=1(封面). 对折后第1页在封面
    (1, 14),   # Sheet 1 back:  left=2 right=15
    (13, 2),   # Sheet 2 front: left=14 right=3
    (3, 12),   # Sheet 2 back:  left=4 right=13
    (11, 4),   # Sheet 3 front: left=12 right=5
    (5, 10),   # Sheet 3 back:  left=6 right=11
    (9, 6),    # Sheet 4 front: left=10 right=7
    (7, 8),    # Sheet 4 back:  left=8 right=9 (中心跨页)
]

# Load all slide images
slides = []
for i in range(1, 17):
    path = os.path.join(SRC_DIR, f'slide_{i:02d}.png')
    slides.append(Image.open(path))
    print(f'Loaded slide {i}: {slides[-1].size}')

# Check consistency
first_w, first_h = slides[0].size
for i, img in enumerate(slides):
    if img.size != (first_w, first_h):
        print(f'  WARNING: slide {i+1} size {img.size} differs from {first_w}x{first_h}')

print(f'\nGenerating saddle-stitch booklet ({len(page_pairs)} A4 pages)...')

c = canvas.Canvas(OUT, pagesize=landscape(A4))

for idx, (left_i, right_i) in enumerate(page_pairs):
    sheet = idx // 2 + 1
    side = "正面" if idx % 2 == 0 else "背面"
    
    # Left half: x=0 to x=A4_W/2
    left_img = slides[left_i]
    c.drawImage(left_img.filename, 0, 0, width=A4_W/2, height=A4_H, 
                preserveAspectRatio=True, anchor='c')
    
    # Right half: x=A4_W/2 to x=A4_W
    right_img = slides[right_i]
    c.drawImage(right_img.filename, A4_W/2, 0, width=A4_W/2, height=A4_H,
                preserveAspectRatio=True, anchor='c')
    
    # Copyright watermark at bottom center with decorative lines
    c.setFont(FONT_NAME, 8)
    c.setFillColorRGB(0.51, 0.51, 0.51)  # gray
    copyright_text = "版权所有 © 2026 改变从整理开始"
    text_width = c.stringWidth(copyright_text, FONT_NAME, 8)
    y = 17
    c.drawCentredString(A4_W / 2, y, copyright_text)
    
    # Decorative lines on both sides
    c.setStrokeColorRGB(0.51, 0.51, 0.51)
    c.setLineWidth(0.5)
    margin = 12
    gap = 14
    line_y = y + 1  # slightly above baseline
    # Left line: from left margin to just before text
    left_end = A4_W / 2 - text_width / 2 - gap
    c.line(margin, line_y, left_end, line_y)
    # Right line: from just after text to right margin
    right_start = A4_W / 2 + text_width / 2 + gap
    c.line(right_start, line_y, A4_W - margin, line_y)
    
    print(f'  Page {idx+1} (Sheet {sheet} {side}): left=slide {left_i+1}, right=slide {right_i+1}')
    c.showPage()

c.save()

# Close all images
for img in slides:
    img.close()

# --- Add encryption: allow reading, prevent modification ---
from pypdf import PdfReader, PdfWriter
from pypdf.constants import UserAccessPermissions

OWNER_PASSWORD = "Change@20260327"
reader = PdfReader(OUT)
writer = PdfWriter()

for page in reader.pages:
    writer.add_page(page)

# user_password="" means anyone can open without a password
# owner_password restricts editing permissions
# Allowed: printing (high-res), copying text, screen reading
# Disallowed: modifying, annotating, form filling, assembling
allowed = (
    UserAccessPermissions.PRINT
    | UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
    | UserAccessPermissions.PRINT_TO_REPRESENTATION
)
writer.encrypt(user_password="", owner_password=OWNER_PASSWORD, permissions_flag=allowed)

# Write to temporary file, then replace
encrypted_path = OUT.replace('.pdf', '_encrypted.pdf')
with open(encrypted_path, 'wb') as f:
    writer.write(f)

os.replace(encrypted_path, OUT)
print(f'\nPDF encrypted: readable by anyone, owner password = "{OWNER_PASSWORD}"')
# --- End encryption ---

# Get output size
size = os.path.getsize(OUT)
print(f'Output: {OUT}')
print(f'Size: {size/1024:.0f} KB, {len(page_pairs)} A4 landscape pages ({len(page_pairs)//2} sheets)')
print(f'Print duplex, short-edge flip, fold in half, staple at spine.')
