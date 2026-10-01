---
name: pptx-booklet-pdf
description: Convert a PPT/PPTX slideshow (A5-sized pages) into a saddle-stitch printable PDF booklet on A4 landscape paper. This skill should be used when the user wants to print a booklet from a PPT, 手账本打印, 骑马钉装订, or generate a print-ready saddle-stitch PDF from presentation slides.
agent_created: true
---

# PPTX → Saddle-Stitch Booklet PDF

Convert a PPT/PPTX file (A5-sized slides) into a saddle-stitch printable PDF for A4 double-sided printing. Each A4 sheet holds 2 A5 pages. After printing double-sided (short-edge flip), folding in half, and stapling at the spine, the page order matches the original PPT sequence.

## Prerequisites

- Windows with Microsoft PowerPoint installed (COM automation)
- Python 3 with `reportlab`, `Pillow`, `pypdf` installed in the venv at `~/.workbuddy/binaries/python/envs/default/`
- Chinese font `msyh.ttc` at `C:/Windows/Fonts/msyh.ttc` (used for copyright watermark)

> ⚠️ **Skill self-validates prerequisites on first run.** If `~/.workbuddy/binaries/python/envs/default/Scripts/python.exe` does not exist, create the venv first and install deps:
> ```bash
> "C:/Users/Jihai/.workbuddy/binaries/python/versions/3.13.12/python.exe" -m venv "C:/Users/Jihai/.workbuddy/binaries/python/envs/default"
> "C:/Users/Jihai/.workbuddy/binaries/python/envs/default/Scripts/python.exe" -m pip install reportlab Pillow pypdf
> ```

## 7-Step Workflow

### Step 1: Obtain the PPTX File

If the user provides a Tencent Docs URL (`docs.qq.com/slide/...`):
- Use the `tencent-docs` skill to export the file
  - `manage.export_file` with `{"file_id": "<id>"}` → returns `task_id`
  - Poll `manage.export_progress` with `{"task_id": "..."}` every ~5s; when `progress=100` returns a 30-min signed `file_url`
  - `curl -L -o "name.pptx" "<file_url>"` to download
- If MCP connector is unavailable, ask the user to download PPTX locally

If the user provides a local `.pptx` path, use it directly. **For PowerPoint COM reliability, copy any non-ASCII filename to a short ASCII filename** (e.g. `sep2026.pptx`) — PowerShell 5.1 chokes on Chinese paths in `.ps1` scripts (HRESULT 0x8007007B).

### Step 1.5: Pre-flight Repair (Tencent-exported PPTX)

> ⚠️ **CRITICAL** for `docs.qq.com` exports. Tencent's PPTX exporter often emits sentinel values `<a:off x="-9223372036854775808" y="-9223372036854775808"/><a:ext cx="-2147483648" cy="-2147483648"/>` in a `p:grpSpPr`/`<a:xfrm>` block. PowerPoint COM rejects the whole file with the unhelpful error "PowerPoint could not open the file."

**Repair** by replacing each bad `<a:xfrm>` with the unit transform (use the matching `chOff`/`chExt` values that follow in the same block — preserves layout exactly):

```python
import zipfile, re
zin = zipfile.ZipFile('source.pptx')
data = {n: zin.read(n) for n in zin.namelist() if not n.endswith('/')}
BAD = '<a:off x="-9223372036854775808" y="-9223372036854775808"/><a:ext cx="-2147483648" cy="-2147483648"/>'
fixed_count = 0
for n in list(data):
    if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
        s = data[n].decode('utf-8')
        for m in re.finditer(r'<a:off x="-9223372036854775808" y="-9223372036854775808"/><a:ext cx="-2147483648" cy="-2147483648"/><a:chOff x="(-?\d+)" y="(-?\d+)"/><a:chExt cx="(-?\d+)" cy="(-?\d+)"/>', s):
            repl = f'<a:off x="{m.group(1)}" y="{m.group(2)}"/><a:ext cx="{m.group(3)}" cy="{m.group(4)}"/><a:chOff x="{m.group(1)}" y="{m.group(2)}"/><a:chExt cx="{m.group(3)}" cy="{m.group(4)}"/>'
            s = s.replace(m.group(0), repl, 1)
            fixed_count += 1
        data[n] = s.encode('utf-8')
# repack with [Content_Types].xml first (safer)
ordered = ['[Content_Types].xml', '_rels/.rels'] + [n for n in data if n not in ('[Content_Types].xml', '_rels/.rels')]
with zipfile.ZipFile('repaired.pptx', 'w', zipfile.ZIP_DEFLATED) as zout:
    for n in ordered:
        zout.writestr(n, data[n])
print(f'fixed {fixed_count} bad xfrm blocks')
```

**If PowerPoint still refuses after this repair** (e.g. damage in master/theme rather than a slide), bisect the slides: build variants keeping slide 1 + slides [2..k] and try opening each; first k that fails identifies a problematic slide in 2..k. Continue bisecting to find the single offender.

### Step 2: Export Slides as 300 DPI PNGs

Write a **.ps1 script that logs to a file** (PowerShell COM output is sometimes swallowed by the host):

```powershell
$log = "<work_dir>\export_log.txt"
$outDir = "<work_dir>\png"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open("<pptx_path>", $true, $false, $true)
    Add-Content $log "SlideCount=$($pres.Slides.Count)"
    foreach ($slide in $pres.Slides) {
        $slide.Export("$outDir\slide_$($slide.SlideNumber.ToString('00')).png", "PNG", 1754, 2480)
    }
    $pres.Close()
    $ppt.Quit()
    Add-Content $log "Done. Files: $((Get-ChildItem $outDir -Filter *.png).Count)"
} catch {
    Add-Content $log "ERROR: $($_.Exception.Message)"
}
```

Run it: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<script>.ps1"`, then read the log.

- Resolution: 1754 x 2480 pixels (A5 at 300 DPI)
- Output: one PNG per slide, named `slide_01.png` through `slide_NN.png`
- Open args: `ReadOnly=$true, Untitled=$false, WithWindow=$true`
- Use **ASCII filenames** in the script — Chinese characters in the path crash with HRESULT 0x8007007B

### Step 3: Verify Slide Count

Confirm the number of slides is a multiple of 4 (e.g., 8, 12, 16, 20). Each saddle-stitch signature requires 4 pages per sheet.

### Step 4: Compute Saddle-Stitch Page Pairs

For N slides (0-indexed), the pairing follows this pattern:

- Sheet 1 front: (N-1, 0) — back cover left, front cover right
- Sheet 1 back:  (1, N-2)
- Sheet 2 front: (N-3, 2)
- Sheet 2 back:  (3, N-4)
- ...continuing inward until reaching the center spread

The left side of each A4 page gets the first index, the right side gets the second. After folding, the cover (slide 1 / index 0) appears on the front.

### Step 5: Generate the A4 Landscape PDF

Use `scripts/booklet_reportlab.py` as a template. Key parameters to update in the script:

| Variable | Description |
|----------|-------------|
| `SRC_DIR` | Path to the exported PNGs directory |
| `OUT` | Output PDF path |
| `page_pairs` | Saddle-stitch pair list (computed in Step 4) |
| `OWNER_PASSWORD` | PDF owner password for encryption |
| `copyright_text` | Footer copyright text (default: "版权所有 © 2026 改变从整理开始") |

The script uses `reportlab` to:
- Place two A5 images side-by-side on A4 landscape (842 x 595 pt)
- Add a footer with copyright text and decorative horizontal lines on both sides
- Watermark: 7pt Microsoft YaHei (msyh.ttc), gray 0.55, text "版权所有 © 2026 改变从整理开始"
- Decorative lines: **span the full content width** — from `EDGE_MARGIN` (24pt) to the text, and mirrored on the right. Line thickness 0.5pt, gray 0.55, aligned to the text's visual midline (`WM_Y + 2.5` for a 7pt font). Use `pdfmetrics.stringWidth(text, font, size)` to compute the text bounds — never hardcode widths.

```python
# Watermark with flanking lines spanning the content width
EDGE_MARGIN = 24   # distance from page edge (echoes the A5 inner margin)
LINE_GAP = 8       # gap between line and text
WM_Y = 6           # baseline y
GRAY = 0.55

c.setFont("MSYH", 7)
c.setFillGray(GRAY)
c.drawCentredString(page_w / 2, WM_Y, WATERMARK)
tw = pdfmetrics.stringWidth(WATERMARK, "MSYH", 7)
text_left  = page_w / 2 - tw / 2
text_right = page_w / 2 + tw / 2
y_mid = WM_Y + 2.5
c.setStrokeGray(GRAY)
c.setLineWidth(0.5)
c.line(EDGE_MARGIN,          y_mid, text_left - LINE_GAP,  y_mid)
c.line(text_right + LINE_GAP, y_mid, page_w - EDGE_MARGIN, y_mid)
```

> 📐 **The user's confirmed preference (Aug 2026): the watermark lines must be flush with the content width, not short dashes.** With `EDGE_MARGIN = 24` each line ends up ≈324pt, visually aligning with the A5 images above. Keep 24pt for all future months.

### Step 6: Encrypt the PDF

The script uses `pypdf` to encrypt the output:
- **User password**: empty (anyone can open and read)
- **Owner password**: `Change@20260327` (default, change as needed)
- **Allowed permissions**: PRINT, EXTRACT_TEXT_AND_GRAPHICS, PRINT_TO_REPRESENTATION
- **Disallowed**: MODIFY, ADD_OR_MODIFY, FILL_FORM_FIELDS, ASSEMBLE_DOC

### Step 7: Name and Deliver Output

**File naming convention**: `2026年X月手账本@改变从整理开始.pdf`

The month should be extracted from the PPT content — the first slide typically contains "2026年 · X月" (Chinese numerals like 捌月、玖月). Use Arabic numerals in the filename (e.g., 8月).

Present the generated PDF and provide printing instructions:

> Print settings: A4 landscape, double-sided, short-edge flip, 100% scale.
> After printing, fold each sheet in half (short edge to short edge).
> Stack sheets: inner (sheet N/4) → ... → outer (sheet 1).
> Staple 2-3 times along the fold line.

## Bundled Scripts

Copy this skill's scripts into a **fresh working directory per run** (never regenerate in place — the user explicitly requires a clean full regeneration, not patching an old output).

| Script | Purpose |
|--------|---------|
| `scripts/repair_pptx.py <src> <dst>` | Fix the Tencent-exported sentinel `xfrm` values (Step 1.5). Rebuilds the zip and validates XML well-formedness. |
| `scripts/export_png.ps1` | PowerPoint COM export of every slide to 300 DPI PNG (Step 2). Edit the three path variables at the top. |
| `scripts/booklet_reportlab.py` | Self-contained A5 PNG → A4 landscape saddle-stitch PDF + encryption (Steps 4-6). Edit the config block. |
| `scripts/render_preview.py` | PyMuPDF verification previews: watermark zoom, a specific slide's zoom, full-page renders. |

## Pitfalls Checklist

1. **Tencent PPTX won't open in PowerPoint** → run `repair_pptx.py`. The sentinel values are usually on one slide, but scan all slides — don't assume a number.
2. **`0x8007007B` from PowerShell** → a Chinese character is in the path *inside the .ps1 file*. Copy the PPTX to a short ASCII name (`sep2026.pptx`) and use ASCII paths in the script.
3. **No output from PowerShell COM** → COM output is swallowed by the host; always `Add-Content $log` to a file and read that file.
4. **Never run Python with f-strings inline via `bash -c`** — the shell eats quotes and produces corrupt XML. Always write a `.py` file, then run it.
5. **`pypdf` ≥ 6 error: `missing 1 required positional argument: 'user_password'`** → pass `user_password=""` explicitly (empty string = opens without password).
6. **Encrypted PDFs cannot be patched page-by-page** → every revision means a full regeneration. Say so plainly when the user asks whether a change was "just page replacement".
7. **Saddle-stitch order sanity check**: sheet 1 front must be `(N, 1)` (back cover left, front cover right). If the cover lands anywhere else the formula is wrong — the output file size is a quick tell (a correct 16-page booklet is ~3.0 MB).
