#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Saddle-stitch booklet generator: A5 PNG slides -> A4 landscape PDF -> encrypted.

Pipeline position: Step 4-6 of the pptx-booklet-pdf skill.

    python make_booklet.py --png-dir ./png --out "./2026年9月手账本@改变从整理开始.pdf"

The page order is computed automatically from the number of PNGs, so nothing needs
to be hardcoded per project. Page count must be a multiple of 4.

Print settings afterwards:
    A4 landscape -> double-sided -> short-edge flip -> 100% scale
    fold each sheet in half, stack 4->3->2->1, staple 2-3 times on the fold line
"""
import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
from pypdf.constants import UserAccessPermissions

# ---------------------------------------------------------------- config
FONT_PATH = r"C:\Windows\Fonts\msyh.ttc"
FONT_NAME = "MSYH"
WATERMARK = "版权所有 © 2026 改变从整理开始"
DEFAULT_OWNER_PWD = "Change@20260327"

# Watermark decoration: lines must be flush with the content width, not short dashes.
EDGE_MARGIN = 24    # distance from the page edge, pt (echoes the A5 inner margin)
LINE_GAP = 8        # gap between the line and the text, pt
LINE_W = 0.5        # line thickness, pt
WM_Y = 6            # watermark baseline, pt
GRAY = 0.55         # grey level for text + lines


def build_page_order(n):
    """Saddle-stitch imposition order for n pages (1-indexed page numbers).

    Sheet i front = (n - 2(i-1), 2(i-1) + 1)
    Sheet i back  = (2i,         n - 2i + 1)

    Sanity check: the first entry must be (n, 1) - back cover left, front cover right.
    """
    sheets = n // 4
    order = []
    for i in range(1, sheets + 1):
        order.append((n - 2 * (i - 1), 2 * (i - 1) + 1))   # front
        order.append((2 * i, n - 2 * i + 1))               # back
    return order


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--png-dir", required=True, help="directory holding slide_01.png ...")
    ap.add_argument("--out", required=True, help="output PDF path")
    ap.add_argument("--owner-password", default=DEFAULT_OWNER_PWD)
    ap.add_argument("--watermark", default=WATERMARK)
    ap.add_argument("--no-encrypt", action="store_true", help="skip PDF encryption")
    args = ap.parse_args()

    pages = sorted(f for f in os.listdir(args.png_dir) if f.lower().endswith(".png"))
    n = len(pages)
    if n == 0:
        sys.exit(f"no PNGs found in {args.png_dir}")
    assert n % 4 == 0, f"page count {n} is not a multiple of 4 - cannot saddle-stitch"

    order = build_page_order(n)
    assert order[0] == (n, 1), f"page order sanity check failed: {order[0]} != {(n, 1)}"

    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))

    page_w, page_h = landscape(A4)   # 841.89 x 595.28 pt
    half_w = page_w / 2              # one A5 per half, A5 is the same ratio as A4 half
    img_w, img_h = half_w, page_h

    tmp_pdf = os.path.splitext(args.out)[0] + "_raw.pdf"
    c = canvas.Canvas(tmp_pdf, pagesize=(page_w, page_h))

    for left_no, right_no in order:
        # left half
        c.drawImage(os.path.join(args.png_dir, pages[left_no - 1]), 0, 0,
                    width=img_w, height=img_h)
        # right half
        c.drawImage(os.path.join(args.png_dir, pages[right_no - 1]), half_w, 0,
                    width=img_w, height=img_h)

        # copyright footer, centred, with lines spanning the content width
        c.setFont(FONT_NAME, 7)
        c.setFillGray(GRAY)
        c.drawCentredString(page_w / 2, WM_Y, args.watermark)

        tw = pdfmetrics.stringWidth(args.watermark, FONT_NAME, 7)
        text_left = page_w / 2 - tw / 2
        text_right = page_w / 2 + tw / 2
        y_mid = WM_Y + 2.5  # visual midline of a 7pt glyph
        c.setStrokeGray(GRAY)
        c.setLineWidth(LINE_W)
        c.line(EDGE_MARGIN, y_mid, text_left - LINE_GAP, y_mid)
        c.line(text_right + LINE_GAP, y_mid, page_w - EDGE_MARGIN, y_mid)
        c.showPage()

    c.save()
    print(f"raw booklet: {len(order)} A4 pages")

    if not args.no_encrypt:
        # readable by anyone, printable, copyable; editing forbidden
        reader = PdfReader(tmp_pdf)
        writer = PdfWriter()
        for p in reader.pages:
            writer.add_page(p)
        writer.encrypt(
            user_password="",          # pypdf >= 6 requires this explicitly
            owner_password=args.owner_password,
            permissions_flag=(
                UserAccessPermissions.PRINT
                | UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
                | UserAccessPermissions.PRINT_TO_REPRESENTATION
            ),
        )
        with open(args.out, "wb") as f:
            writer.write(f)
        os.remove(tmp_pdf)
        print(f'Owner password: "{args.owner_password}" (user password: empty)')
    else:
        os.replace(tmp_pdf, args.out)

    print(f"Done. PDF: {args.out}")
    print(f"Size: {os.path.getsize(args.out) / 1024:.0f} KB, "
          f"{len(order)} A4 pages ({n // 4} sheets)")
    print("Print: A4 landscape, double-sided, short-edge flip, 100% scale.")


if __name__ == "__main__":
    main()
