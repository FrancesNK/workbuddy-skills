#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Render verification previews of the finished booklet PDF (PyMuPDF).

    python render_preview.py --pdf "2026年9月手账本@改变从整理开始.pdf" --out-dir preview

Produces:
    wm_01.png        zoomed crop of the page footer (watermark + flanking lines)
    page{N}_zoom.png zoomed crop of a full PDF page (spot-check any imposition page)
    preview_NN.png   full-page renders

Imposition reminder for a 16-page booklet:
    PDF p1 = (16, 1)   p2 = (2, 15)   p3 = (14, 3)   p4 = (4, 13)
    PDF p5 = (12, 5)   p6 = (6, 11)   p7 = (10, 7)   p8 = (8, 9)
So slide 5 lives on the left half of PDF page 5.
"""
import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

import fitz  # PyMuPDF


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out-dir", default="preview")
    ap.add_argument("--pages", type=int, nargs="*", default=[0],
                    help="0-indexed PDF pages to render in full (default: 0)")
    ap.add_argument("--zoom-pages", type=int, nargs="*", default=[],
                    help="0-indexed PDF pages to render as a zoomed crop")
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    doc = fitz.open(args.pdf)
    made = []

    # footer strip of page 1: watermark and its two horizontal rules
    page = doc[0]
    r = page.rect
    clip = fitz.Rect(0, r.height - 40, r.width, r.height)
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip)
    pix.save(os.path.join(args.out_dir, "wm_01.png"))
    made.append("wm_01.png")

    # zoomed crops
    for idx in args.zoom_pages:
        page = doc[idx]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        fn = f"page{idx + 1}_zoom.png"
        pix.save(os.path.join(args.out_dir, fn))
        made.append(fn)

    # full pages
    for idx in args.pages:
        pix = doc[idx].get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
        fn = f"preview_{idx + 1:02d}.png"
        pix.save(os.path.join(args.out_dir, fn))
        made.append(fn)

    print("previews saved:", made)


if __name__ == "__main__":
    main()
