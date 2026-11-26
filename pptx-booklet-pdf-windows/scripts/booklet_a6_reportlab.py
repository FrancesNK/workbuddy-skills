# -*- coding: utf-8 -*-
"""A6 booklet composer (Windows / reportlab).

A4 portrait sheet holds 4 A6 pages (2x2). After printing, cut horizontally
along the center dashed line into two A5 halves, nest, fold, saddle-stitch.

This is the A6 twin of make_booklet.py: same PNG input (1754x2480 A5@300DPI
exports are reused as-is), same watermark text/gray, same encryption rules.

Usage:
  python booklet_a6_reportlab.py --png-dir C:\\work\\png --out "C:\\work\\booklet_a6.pdf" ^
      [--owner-password <pw>] [--font-path C:\\Windows\\Fonts\\msyh.ttc]

Output naming convention: <year>年<M>月手账本@改变从整理开始(A6版).pdf
"""
import argparse
import glob
import math
import os
import sys

MM = 72 / 25.4
SW, SH = 210 * MM, 297 * MM          # A4 portrait sheet: 595.28 x 841.89 pt
PW, PH = 105 * MM, 148.5 * MM        # A6 finished page: 297.64 x 420.94 pt
K = PW / 421.0                       # A5(421pt) -> A6 scale ~0.707

GRAY = 0.55                          # watermark gray (matches A5 version)
CUT_GRAY = 0.72                      # dashed cut line
SCISSOR_GRAY = 0.58
SC = 11                              # scissor glyph size (pt)

DEFAULT_FONT = 'C:/Windows/Fonts/msyh.ttc'
DEFAULT_WM = '版权所有 © 2026 改变从整理开始'


def saddle_pairs(n):
    """1-based saddle-stitch pairs, sheet order like the A5 version.

    N=16 -> (16,1) (2,15) (14,3) (4,13) (12,5) (6,11) (10,7) (8,9)
    """
    pairs = []
    for i in range(1, n // 2 + 1):
        pairs.append((n + 1 - i, i) if i % 2 else (i, n + 1 - i))
    return pairs


def build_sides(n):
    """PDF page list. Each item = [(top_l, top_r), (bottom_l, bottom_r)].

    Sheet s: front = (pair[4s-3] top, pair[4s-1] bottom),
             back  = (pair[4s-2] top, pair[4s]   bottom).
    N=16 -> [[(16,1),(14,3)], [(2,15),(4,13)], [(12,5),(10,7)], [(6,11),(8,9)]]
    """
    pairs = saddle_pairs(n)
    assert pairs[0] == (n, 1), 'saddle order formula broken: first pair must be (N,1)'
    sides = []
    for s in range(n // 8):   # one A4 sheet = 2 sides x 4 A6 pages = 8 pages
        front = [pairs[4 * s], pairs[4 * s + 2]]
        back = [pairs[4 * s + 1], pairs[4 * s + 3]]
        sides.append(front)
        sides.append(back)
    return sides


def scissors(c, x, y, s, mirror=False):
    """Small vector scissors centered on the cut line at (x, y). No font dependency."""
    def X(v):
        return x + s - v if mirror else x + v
    c.setStrokeColorRGB(SCISSOR_GRAY, SCISSOR_GRAY, SCISSOR_GRAY)
    c.setLineWidth(0.5)
    c.circle(X(0.20 * s), y - 0.17 * s, 0.11 * s, stroke=1, fill=0)
    c.circle(X(0.20 * s), y + 0.17 * s, 0.11 * s, stroke=1, fill=0)
    c.line(X(0.20 * s), y - 0.17 * s, X(0.95 * s), y + 0.13 * s)
    c.line(X(0.20 * s), y + 0.17 * s, X(0.95 * s), y - 0.13 * s)


def draw_half(c, y_off, left_png, right_png, font_name, wm_text, tw):
    """One A5 half (reportlab origin = bottom-left of the A4 sheet).

    y_off = 0 (lower half) or PH (upper half). Watermark sits centered on the
    half's fold seam, baseline lifted (6.6+8.6)*K from the half's bottom edge;
    flanking lines span 24*K .. text-gap, mirrored — all scaled by K.
    """
    c.drawImage(left_png, 0, y_off, PW, PH)
    c.drawImage(right_png, PW, y_off, PW, PH)
    fold = PW
    bottom = y_off
    y = bottom + 8.5 * K
    c.setStrokeColorRGB(GRAY, GRAY, GRAY)
    c.setLineWidth(0.5)
    c.line(24 * K, y, fold - (tw / 2 + 8 * K), y)
    c.line(fold + (tw / 2 + 8 * K), y, 2 * PW - 24 * K, y)
    c.setFont(font_name, 7 * K)
    c.setFillGray(GRAY)
    c.drawCentredString(fold, bottom + 6.6 * K, wm_text)


def load_pngs(png_dir, n):
    files = sorted(glob.glob(os.path.join(png_dir, '*.png')))
    if len(files) < n:
        raise SystemExit(f'need {n} PNGs, found {len(files)} in {png_dir}')
    return files[:n]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--png-dir', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--owner-password', default='')
    ap.add_argument('--font-path', default=DEFAULT_FONT)
    ap.add_argument('--watermark', default=DEFAULT_WM)
    ap.add_argument('--pad-to-multiple-of-8', action='store_true',
                    help='pad with blank white pages if count is not a multiple of 8 (A6 sheet holds 8 pages)')
    args = ap.parse_args()

    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from pypdf import PdfWriter, PdfReader
    from pypdf.constants import UserAccessPermissions
    from PIL import Image, ImageDraw

    font_name = 'MSYH'
    try:
        pdfmetrics.registerFont(TTFont(font_name, args.font_path, subfontIndex=0))
    except Exception as e:
        sys.exit(f'cannot register font {args.font_path}: {e}')

    files = sorted(glob.glob(os.path.join(args.png_dir, '*.png')))
    n = len(files)
    if n == 0:
        sys.exit(f'no PNGs in {args.png_dir}')
    if n % 8:
        if not args.pad_to_multiple_of_8:
            sys.exit(f'{n} pages is not a multiple of 8; pass --pad-to-multiple-of-8 to pad blank pages')
        pad = os.path.join(args.png_dir, '_blank_pad.png')
        Image.new('RGB', (1754, 2480), 'white').save(pad)
        while n % 8:
            files.append(pad)
            n += 1

    sides = build_sides(n)
    wm_size = 7 * K
    tw = pdfmetrics.stringWidth(args.watermark, font_name, wm_size)

    tmp_pdf = args.out + '.tmp.pdf'
    c = canvas.Canvas(tmp_pdf, pagesize=(SW, SH))
    for side in sides:
        (tl, tr), (bl, br) = side
        # reportlab origin is bottom-left: lower half first, upper half second
        draw_half(c, 0, files[bl - 1], files[br - 1], font_name, args.watermark, tw)
        draw_half(c, PH, files[tl - 1], files[tr - 1], font_name, args.watermark, tw)
        # cut line: horizontal center of the A4 sheet, dashed + scissors at both ends
        y_cut = PH
        c.setStrokeColorRGB(CUT_GRAY, CUT_GRAY, CUT_GRAY)
        c.setLineWidth(0.4)
        c.setDash(3, 3)
        c.line(SC + 7, y_cut, SW - SC - 7, y_cut)
        c.setDash()
        scissors(c, 4, y_cut, SC)
        scissors(c, SW - 4 - SC, y_cut, SC, mirror=True)
        c.showPage()
        print('side -> top', (tl, tr), '| bottom', (bl, br))
    c.save()

    # encrypt: empty user password (opens freely), owner password restricts edit/copy
    perm = (UserAccessPermissions.PRINT
            | UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
            | UserAccessPermissions.PRINT_TO_REPRESENTATION)
    reader = PdfReader(tmp_pdf)
    used = ''
    for alg in ('AES-256', 'AES-128', 'RC4-128'):   # AES needs cryptography>=3.1
        try:
            w = PdfWriter()
            for p in reader.pages:
                w.add_page(p)
            w.encrypt(user_password='', owner_password=args.owner_password,
                      permissions_flag=perm, algorithm=alg)
            # AES may fail only at write time when `cryptography` is missing,
            # so the write lives inside the try as well
            with open(args.out, 'wb') as f:
                w.write(f)
            used = alg
            break
        except Exception as e:
            print(f'encrypt with {alg} failed: {e}')
    if not used:
        sys.exit('encryption failed for every algorithm')
    os.remove(tmp_pdf)
    print('SAVED', args.out, f'({n // 8} A4 sheets, {n} content pages, {used})')
    if used.startswith('RC4'):
        print('WARN: install `cryptography` (pip install cryptography) to get AES-256 like the A5 version')


if __name__ == '__main__':
    main()
