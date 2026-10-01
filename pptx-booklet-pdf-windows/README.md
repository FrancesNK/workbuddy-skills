# pptx-booklet-pdf (Windows)

把一份 A5 尺寸的 PPT/PPTX 演示文稿转成**可直接双面打印、对折装订的骑马订 PDF**。

> 目录名带 `-windows` 后缀，是因为整条流水线依赖 Windows + 本机安装的 Microsoft PowerPoint（通过 COM 自动化导出高清图）。macOS / Linux 上跑不通第 2 步。

## 它做什么

```
PPTX  ──►  修复  ──►  逐页导 300DPI PNG  ──►  骑马订拼版  ──►  加密 PDF
         腾讯导出     PowerPoint COM          A4 横版双面     可打印 / 禁编辑
         哨兵值修复   1754×2480                 +版权水印
```

输入 16 页 A5 幻灯片 → 输出 8 页 A4 横版（= 4 张纸），打印后对折、装订，翻页顺序与原始 PPT 完全一致。

## 快速开始

```bash
# 0) 依赖（一次性）
python -m venv venv
venv/Scripts/pip install reportlab Pillow pypdf pymupdf

# 1) 腾讯文档导出（或直接准备本地 pptx，跳过这步）
#    manage.export_file -> 轮询 manage.export_progress -> 下载 file_url

# 2) 修复腾讯导出 PPTX 的哨兵值（PowerPoint 否则拒开）
python scripts/repair_pptx.py sep2026.pptx sep2026_repaired.pptx

# 3) 每页导出 300 DPI PNG（Windows + PowerPoint）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/export_png.ps1 \
    -Pptx "C:\work\sep2026_repaired.pptx" -OutDir "C:\work\png" -Log "C:\work\export_log.txt"

# 4) 拼版 + 水印 + 加密
python scripts/make_booklet.py --png-dir "C:\work\png" --out "C:\work\2026年9月手账本@改变从整理开始.pdf"

# 5) 渲染预览自检
python scripts/render_preview.py --pdf "C:\work\2026年9月手账本@改变从整理开始.pdf" --pages 0 --zoom-pages 4
```

## 脚本清单

| 脚本 | 作用 |
|------|------|
| `scripts/repair_pptx.py <src> <dst>` | 修复腾讯导出 PPTX 的非法 `xfrm` 哨兵值（`-9223372036854775808`），并做 XML 良构校验 |
| `scripts/export_png.ps1` | PowerPoint COM 批量导出 300 DPI PNG（1754×2480 = A5@300DPI） |
| `scripts/make_booklet.py` | 骑马订拼版 + 版权水印 + PDF 加密（页序按 PNG 数量自动推算） |
| `scripts/booklet_reportlab.py` | 早期模板版：手动配置页序列表，适合手工微调场景 |
| `scripts/render_preview.py` | PyMuPDF 渲染水印特写 / 指定页特写 / 整页预览，用于交付前自检 |

## 骑马订页序

对 N 页（页码从 1 开始）：

```
第 i 张纸 正面 = (N - 2(i-1),  2(i-1) + 1)
第 i 张纸 背面 = (2i,          N - 2i + 1)
```

第一条必须是 `(N, 1)`——左半是封底、右半是封面。若封面跑到别处，就是公式错了。
`make_booklet.py` 里内置了这条断言，页序不对会直接报错。

## 输出规格

| 项目 | 值 |
|------|-----|
| 页面 | A4 横版 842×595 pt，每张两个 A5 |
| 水印 | 7pt 微软雅黑，灰 0.55，`版权所有 © 2026 改变从整理开始` |
| 水印横线 | 与内容**同宽**（左右各距页边 24pt），0.5pt 粗 |
| 加密 | 打开密码为空；owner 密码见 `--owner-password`（默认 `Change@20260327`） |
| 权限 | 允许打印、允许复制文字；禁止修改 / 批注 / 重组 |

## 打印设置

> A4 横版 → 双面 → 短边翻转 → 100% 缩放
> 打印后每张对折，按 4→3→2→1 叠好，折线处订 2-3 颗钉

## 踩坑清单

1. **PowerPoint 报"无法打开"** → 跑 `repair_pptx.py`。哨兵值通常只出现在一页，但必须扫全部页；若还打不开，用二分法定位（逐步保留前 k 页生成变体试开）。
2. **PowerShell 报 `0x8007007B`** → `.ps1` 文件里出现了中文路径。把 PPTX 复制成短 ASCII 名（`sep2026.pptx`）再用纯 ASCII 路径。
3. **PowerShell COM 没有任何输出** → COM 输出会被宿主吞掉，一律 `Add-Content` 写日志文件再读。
4. **不要用 `bash -c "python -c '...f-string...'"`** → shell 会吃掉引号，生成的 XML 直接损坏。永远先写 `.py` 文件再执行。
5. **`pypdf` ≥ 6 报 `missing 1 required positional argument: 'user_password'`** → 必须显式传 `user_password=""`（空串 = 打开无需密码）。
6. **加密 PDF 无法逐页替换** → 每次修订都必须整本重新生成。源文档一改就要重跑全流程。
7. **页数不是 4 的倍数** → 骑马订需要每张纸 4 个页位，补空白页或调整版式。

## 相关

- 母技能定义见 [`SKILL.md`](./SKILL.md)
- 腾讯文档的导出/编辑工具链见同仓库的 `tencent-docs` 相关技能
