# pptx-booklet-pdf (Windows)

把一份 A5 尺寸的 PPT/PPTX 演示文稿转成**可直接双面打印、对折装订的骑马订 PDF**。**每次运行固定产出两个版本**：

- **A5 版**：A4 横版，每张 2 页，对折 + 骑马钉
- **A6 版**：A4 纵版 2×2，每张 4 页，中间虚线裁开 + 两端小剪刀，套叠对折 + 骑马钉

页图只导出一次，两个版本共用。

> 目录名带 `-windows` 后缀，是因为整条流水线依赖 Windows + 本机安装的 Microsoft PowerPoint（通过 COM 自动化导出高清图）。macOS / Linux 上跑不通第 2 步。

## 它做什么

```
PPTX  ──►  修复  ──►  逐页导 300DPI PNG  ──►  骑马订拼版        ──►  加密 PDF
         腾讯导出     PowerPoint COM          A5: A4横版 2-up      可打印 / 禁编辑
         哨兵值修复   1754×2480               A6: A4纵版 4-up+裁剪线
                                              （两版共用同一套 PNG）
```

输入 16 页 A5 幻灯片 → 输出 A5 版 8 页 A4 横版（4 张纸）+ A6 版 4 页 A4 纵版（2 张纸），打印后裁切、对折、装订，翻页顺序与原始 PPT 完全一致。

## 快速开始

```bash
# 0) 依赖（一次性；cryptography 使加密升级为 AES-256）
python -m venv venv
venv/Scripts/pip install reportlab Pillow pypdf cryptography

# 1) 腾讯文档导出（或直接准备本地 pptx，跳过这步）
#    manage.export_file -> 轮询 manage.export_progress -> 下载 file_url

# 2) 修复腾讯导出 PPTX 的哨兵值（PowerPoint 否则拒开）
python scripts/repair_pptx.py sep2026.pptx sep2026_repaired.pptx

# 3) 每页导出 300 DPI PNG（Windows + PowerPoint，两版共用）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/export_png.ps1 \
    -Pptx "C:\work\sep2026_repaired.pptx" -OutDir "C:\work\png" -Log "C:\work\export_log.txt"

# 4a) A5 版：拼版 + 水印 + 加密
python scripts/make_booklet.py --png-dir "C:\work\png" --out "C:\work\2026年9月手账本@改变从整理开始(A5版).pdf"

# 4b) A6 版：同一套 PNG，2×2 拼版 + 裁剪线 + 剪刀 + 水印 + 加密
python scripts/booklet_a6_reportlab.py --png-dir "C:\work\png" \
    --out "C:\work\2026年9月手账本@改变从整理开始(A6版).pdf" --owner-password "<OWNER_PW>"

# 5) 渲染预览自检
python scripts/render_preview.py --pdf "C:\work\2026年9月手账本@改变从整理开始(A5版).pdf" --pages 0 --zoom-pages 4
```

## 脚本清单

| 脚本 | 作用 |
|------|------|
| `scripts/repair_pptx.py <src> <dst>` | 修复腾讯导出 PPTX 的非法 `xfrm` 哨兵值（`-9223372036854775808`），并做 XML 良构校验 |
| `scripts/export_png.ps1` | PowerPoint COM 批量导出 300 DPI PNG（1754×2480 = A5@300DPI） |
| `scripts/make_booklet.py` | A5 版：骑马订拼版 + 版权水印 + PDF 加密（页序按 PNG 数量自动推算） |
| `scripts/booklet_a6_reportlab.py` | A6 版：同一套 PNG → A4 纵版 2×2，虚线裁剪线 + 矢量小剪刀，水印按 K=0.707 缩放，AES-256 加密（缺 cryptography 自动降级） |
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

| 项目 | A5 版 | A6 版 |
|------|-------|-------|
| 页面 | A4 横版 842×595 pt，每张两个 A5 | A4 纵版 595×842 pt，2×2 四个 A6 |
| 页数换算 | N 页内容 → N/4 张 A4 | N 页内容 → N/8 张 A4（须为 8 的倍数） |
| 水印 | 7pt 微软雅黑，灰 0.55，`版权所有 © 2026 改变从整理开始` | 同左，字号/坐标按 K=0.707 缩放，每半幅一条 |
| 水印横线 | 与内容**同宽**（左右各距页边 24pt），0.5pt 粗 | 同结构，全部 ×K |
| 裁剪线 | 无 | 横向正中虚线 0.4pt 灰 0.72 + 两端矢量小剪刀（SC=11pt） |
| 加密 | 打开密码为空；owner 密码见 `--owner-password`（不在文档中落盘） | 同左，AES-256 优先，缺 cryptography 降级 RC4-128 |
| 权限 | 允许打印、允许复制文字；禁止修改 / 批注 / 重组 | 同左 |

## 打印设置

> A5 版：A4 横版 → 双面 → **短边翻转** → 100% 缩放
> 打印后每张对折，按 4→3→2→1 叠好，折线处订 2-3 颗钉
>
> A6 版：A4 纵版 → 双面 → **长边翻转** → 100% 缩放
> 打印后沿中间虚线裁成两张 A5，套叠、对折、钉 2-3 颗钉
> 自检：A4 第 1 张打完像翻书翻过来，反面左上角应是页码 2、右上角 15

## 踩坑清单

1. **PowerPoint 报"无法打开"** → 跑 `repair_pptx.py`。哨兵值通常只出现在一页，但必须扫全部页；若还打不开，用二分法定位（逐步保留前 k 页生成变体试开）。
2. **PowerShell 报 `0x8007007B`** → `.ps1` 文件里出现了中文路径。把 PPTX 复制成短 ASCII 名（`sep2026.pptx`）再用纯 ASCII 路径。
3. **PowerShell COM 没有任何输出** → COM 输出会被宿主吞掉，一律 `Add-Content` 写日志文件再读。
4. **不要用 `bash -c "python -c '...f-string...'"`** → shell 会吃掉引号，生成的 XML 直接损坏。永远先写 `.py` 文件再执行。
5. **`pypdf` ≥ 6 报 `missing 1 required positional argument: 'user_password'`** → 必须显式传 `user_password=""`（空串 = 打开无需密码）。
6. **加密 PDF 无法逐页替换** → 每次修订都必须整本重新生成。源文档一改就要重跑全流程。
7. **页数不是 4 的倍数（A5 版）/ 不是 8 的倍数（A6 版）** → 骑马订需要整印张，补空白页或调整版式。A6 版脚本用 `--pad-to-multiple-of-8` 自动补。
8. **pypdf 报 `cryptography>=3.1 is required`** → AES 加密缺依赖，`pip install cryptography`；缺包时 A6 脚本自动降级 RC4-128 并告警，仍可用但强度低。
9. **owner 密码不要写进提交的文件** → 文档里一律用 `<OWNER_PW>` 占位符，真实密码运行时经 `--owner-password` 传入。

## 相关

- 母技能定义见 [`SKILL.md`](./SKILL.md)
- 腾讯文档的导出/编辑工具链见同仓库的 `tencent-docs` 相关技能
