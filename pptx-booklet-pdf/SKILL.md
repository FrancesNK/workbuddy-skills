---
name: pptx-booklet-pdf
description: 把腾讯文档 PPT（docs.qq.com/slide）转成骑马钉装订的手账本 PDF：服务端渲染高清页图→按骑马钉顺序拼版，每期固定产出 A5 版（2-up 折页）与 A6 版（4-up 裁切含剪刀裁剪线）两个 PDF→底部版权水印→AES 加密。适用于「整理疗愈手账本」月度系列。
agent_created: true
---

# pptx-booklet-pdf：腾讯文档 PPT → 骑马钉手账本 PDF

## 两版规格速查

| | A5 版 | A6 版（A5 的一半） |
|---|---|---|
| 成品 | 148×210mm | 105×148.5mm |
| 印张 | A4 **横向** 841.89×595.28 | A4 **纵向** 595.28×841.89 |
| 一张印张 | 2 页（左右） | 4 页（2×2） |
| 印面数（16页内容） | 8 | 4 |
| 装订方式 | 对折 → 骑马钉 | 横向正中裁成 2 张 A5 → 套叠对折 → 骑马钉 |
| 双面设置 | 长边翻转 | 长边翻转（若上下颠倒则改短边） |
| 裁剪线 | 无（只折） | 有（虚线 + 两端小剪刀） |
| 文件名 | `…@改变从整理开始(A5版).pdf` | `…@改变从整理开始(A6版).pdf` |

两版的页图是同一套（300dpi 导出图可直接复用），差异只在拼版与水印缩放。

> **默认规则（用户明确要求）：每期生成时 A5 版和 A6 版两个 PDF 都要产出**，即使本轮只被要求了其中一个版本。页图只导出一次，两版共用；除非用户当期明确说只要某一版。

## 输入

- 用户给出 PPT 链接（如 `https://docs.qq.com/slide/<FILE_ID>`），取 `FILE_ID`。
- 参考9月成品规格：`/Users/yufengs/Downloads/2026年9月手账本@改变从整理开始.pdf`（可解析验证）。

## 环境（关键坑）

1. CLI 位置：`/Users/yufengs/.workbuddy/plugins/cache/workbuddy-builtin/tencent-docs-plugin/<版本>/skills/tencent-docs/tencentdocs.py`。**版本会升级导致硬编码路径失效**——用 glob 匹配 `tencent-docs-plugin/*/skills/tencent-docs` 并按版本号取最大者（2026-11 时为 5.11.3）。
   ```python
   import glob, re
   paths = glob.glob('/Users/yufengs/.workbuddy/plugins/cache/workbuddy-builtin/tencent-docs-plugin/*/skills/tencent-docs')
   key = lambda p: tuple(int(x) for x in re.search(r'/(\d+)\.(\d+)\.(\d+)-', p).groups())
   sys.path.insert(0, max(paths, key=key))
   ```
2. **必须覆盖 API base**：凭据网关下发的 `www-docs.workbuddy.cn` 公网 NXDOMAIN，所有调用前设置：
   ```bash
   export TDOC_API_BASE_URL=https://docs.qq.com
   ```
3. 本地代理偶发 `Tunnel connection failed: 502`（瞬时），重试即可。
4. 沙箱内 pip 长安装易被 SIGTERM(137)，用 run_in_background；`mkdir/cat/ls` 的 brokered shim 可能损坏，用 `/bin/mkdir` 等绝对路径。
5. 依赖：`pymupdf`（拼版+加密，自带AES），装在 venv `/Users/yufengs/.workbuddy/binaries/python/envs/default`。

## 流程

> **最快路径**：复制 skill 目录里的 `dec_compose_a6.py`（A6）或 `compose_v4.py`（A5），只改 `FID` / `PAGES` / 输出文件名即可跑完「导出→下载→拼版→水印→加密」；导出与拼版在同一脚本里，约 2–4 分钟。
>
> **每期固定产出两个版本**：页图导出一次后，先跑 A5 拼版再跑 A6 拼版（或合并进同一脚本，导出图片和隐藏页检测结果全程复用），交付 `…(A5版).pdf` + `…(A6版).pdf` 两个文件。

### 1. 读页数与页脚页码

```bash
python3 tencentdocs.py tdoc_call slide-mcp slide_get_info '{"file_id":"<ID>"}'
# → slide_count, 页面尺寸(w_pt≈421, h_pt≈595)
```
页脚页码用 `slide_get_page_info`（page_index 0-based）读每个 shape 的 text/bounds.y 最大者。手账本页脚页码 = 幻灯片顺序（封面=1 无页码，内页2~N-1，封底=N）。

**必须同时排除隐藏页**（用户会临时隐藏某一页不计入）：接口都读不到隐藏状态，导出图也不跳过。唯一可靠方法——导出 PPTX 解压查 `show="0"`：

```python
res, _ = td.call_tool('tencent-docs', 'manage.export_file', {"file_id": FID})
task = json.loads(res['result']['content'][0]['text'])['task_id']
# 轮询 manage.export_progress 直到 progress=='100' 且 file_url 以 http 开头、长度>20
# 解压后按 presentation.xml 的 sldIdLst + rels 顺序遍历 ppt/slides/slideN.xml，根节点 show="0" 即隐藏
```

辅助手段：隐藏页常带「重复的页脚页码」（两页同为 6），可先扫页脚找断点，再用 PPTX 法确认。剔除后页脚页码应连续 2~N-1。

### 2. 分批导出高清页图（300dpi）

**一次超过约5页会服务端 xvfb 超时**，按 4 页一批：

```bash
python3 tencentdocs.py tdoc_call slide-mcp slide_export_pages_to_image_urls \
  '{"file_id":"<ID>","long_edge":2480,"page_indices":[0,1,2,3]}'
# 返回 images:[{page_index, image_url, width:1755, height:2480}]
```

下载时**必须带 UA + Referer**，否则 403：

```python
headers = {'User-Agent': 'Mozilla/5.0 ... Chrome/126 ...', 'Referer': 'https://docs.qq.com/'}
urllib.request.Request(url, headers=headers)
```

保存为 `/tmp/slide_XX.png`（XX=page_index 两位数）。

### 3. 水印：必须走「位图」路线（唯一不漏乱码的做法）

> ⚠️ 踩过两次坑：① 内置 `china-s` → 乱码；② 抽 `PingFang.ttc` 里的 ttf 嵌入 → 本机 pymupdf 渲染正常、`get_text()` 也正常，但用户用 WPS/预览打开仍是乱码（PingFang 是 **CFF2 可变字体**，PDF 里显示 `/Subtype /Type0` + `otf`，第三方阅读器解析不了）。
> **结论：水印不要嵌任何字体，先把文字渲染成透明 PNG，再当图片贴进 PDF。** 像素在本地定稿，任何阅读器都不可能乱码。

**步骤 A：渲染水印 PNG**（PIL + FreeType 能正常解析 PingFang 的 CFF2）

```python
from PIL import Image, ImageDraw, ImageFont
TEXT, SCALE, GRAY = "版权所有 © 2026 改变从整理开始", 20, 140   # 20x 超采样，55%灰
font = ImageFont.truetype('/tmp/PingFangSC-Regular.ttf', 7 * SCALE)  # 缺字体时退回系统 ttc
bbox = ImageDraw.Draw(Image.new('RGBA', (10, 10))).textbbox((0, 0), TEXT, font=font)
w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
img = Image.new('RGBA', (w + 40, h + 40), (0, 0, 0, 0))
ImageDraw.Draw(img).text((20 - bbox[0], 20 - bbox[1]), TEXT, font=font, fill=(GRAY, GRAY, GRAY, 255))
img.save('/tmp/watermark.png')      # → 2169x172，即 108.45 x 8.6 pt
```

若 `/tmp/PingFangSC-Regular.ttf` 不在（/tmp 会被清理），用 fonttools 从 `/System/Library/Fonts/PingFang.ttc` 抽 SC Regular 后保存一份到 `~/.workbuddy/skills/pptx-booklet-pdf/PingFangSC-Regular.ttf` 复用。

**步骤 B：拼版并贴水印**

```python
import pymupdf
W, H, HALF, N = 841.8898, 595.2756, 420.9449, 16  # N=页数
GRAY = (0.55, 0.55, 0.55)
WM_W, WM_H = 2169/20, 172/20            # 位图尺寸(pt)
WM_X = 367.7698 - 1.0                   # 位图内墨迹左留白 1pt → 墨迹起点对齐 367.77
WM_Y = H - 6 - 9.2                      # 位图内基线距顶 9.2pt → 基线落在距底 6pt
doc = pymupdf.open()
for i in range(1, N//2 + 1):
    left, right = (N+1-i, i) if i % 2 else (i, N+1-i)   # 骑马钉顺序
    page = doc.new_page(width=W, height=H)
    page.insert_image(pymupdf.Rect(0, 0, HALF, H), filename=f'/tmp/slide_{left-1:02d}.png')
    page.insert_image(pymupdf.Rect(HALF, 0, W, H), filename=f'/tmp/slide_{right-1:02d}.png')
    y = H - 8.5
    page.draw_line(pymupdf.Point(24, y), pymupdf.Point(359.7698, y), color=GRAY, width=0.5)
    page.draw_line(pymupdf.Point(482.1199, y), pymupdf.Point(817.8898, y), color=GRAY, width=0.5)
    page.insert_image(pymupdf.Rect(WM_X, WM_Y, WM_X + WM_W, WM_Y + WM_H),
                      filename='/tmp/watermark.png', overlay=True)
```

水印坐标全部取自9月参考版实测：文字墨迹起点 x=367.77、基线距底 6pt、7pt 字号、墨迹总宽 106.1pt；横线 y=8.5、0.5pt 宽、24→359.77 与 482.12→817.89。

骑马钉顺序示例（N=16）：`[16,1] [2,15] [14,3] [4,13] [12,5] [6,11] [10,7] [8,9]`（第1张=封底|封面）。

### 4. 加密保存

```python
perm = pymupdf.PDF_PERM_ACCESSIBILITY | pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ
doc.save(out, garbage=4, deflate=True,
         encryption=pymupdf.PDF_ENCRYPT_AES_256,
         owner_pw='<OWNER_PW>', user_pw='', permissions=perm)
```

- 空用户密码 = 正常打开即读，但限制编辑/复制；**owner 密码以用户当次指定为准**（owner 密码以用户当期指定为准，文档中脱敏为 `<OWNER_PW>`）。
- **只改密码/只改名时不要重跑导出**：`pymupdf.open(旧pdf).save(新, garbage=4, deflate=True, encryption=..., owner_pw=<新>, user_pw='', permissions=perm)` 即可。
- 输出命名（2026-11 起）：A5 版 `~/Downloads/<年>年<月>月手账本@改变从整理开始(A5版).pdf`，A6 版 `...@改变从整理开始(A6版).pdf`。旧的无版本后缀文件名保留不动，覆盖 Downloads 里已存在的同名文件会被文件策略拦截。

### 5. 验证

- `pypdf` 读 pages/size（AES 校验需 cryptography 包，缺失时改用 pymupdf 验证）。
- `sips -s format png --resampleWidth 1600 <pdf> --out /tmp/check.png` 渲染第1张，Read 检查：封底在左、封面在右、底部水印与横线齐全。
- 注意：`sips`/预览只渲染**第1页**；检查其他页用 pymupdf `get_pixmap`。

## 附：A6 版（A5 的一半，A4 一张印 4 页）

A6 版已与 A5 版并列为**每期默认交付**（不再只在用户点名「A5 的一半」时才做）。成品是 **A6 105×148.5mm**。做法与 A5 版不同：A4 纵向 2×2 摆 4 个 A6 页 → 横向从**正中裁成两张 A5** → 4 张 A5 套叠 → 对折 → 骑马钉（贴齐书口）。

几何（pt）：印张 A4 = 595.28×841.89；页 A6 = 297.64×420.95，正好四等分；缩放比 `K = 297.64/421 ≈ 0.707`（水印、横线、内缩量全部乘 K，视觉比例与 A5 版一致）。

每张 A4 = 2 个印面，每面上下两幅（各 2 页）：

```python
SIDES = [
    [(16, 1), (14, 3)],   # A4 第1张 正面：上=最外层折手，下=第2层
    [(2, 15), (4, 13)],   # A4 第1张 反面
    [(12, 5), (10, 7)],   # A4 第2张 正面
    [(6, 11), (8, 9)],    # A4 第2张 反面
]
# 上半幅 y:0..PH，下半幅 y:PH..2PH；每个半幅内 左页 x:x0..x0+PW，右页 x:x0+PW..x0+2PW，折缝在 x0+PW
```

- **必须提示双面打印选「长边翻转／书籍式」**（A4 纵向），否则反面上下颠倒。自检：打完 A4 第 1 张，像翻书一样翻过来，反面左上角应是页码 2、右上角 15。
- **裁剪线**：画在 `y = PH`（横向正中），淡灰 `(0.72,0.72,0.72)`、0.4pt 虚线 `dashes="[3 3] 0"`，从 `SC+7` 画到 `SW-SC-7`；**两端各一个矢量小剪刀**（用户明确要求，比纯实线更直观）。版权水印照旧保留，每半幅一条（居中于该半幅的折缝），不能因为版面小就省掉。
- 剪刀用矢量绘制（**不要用文字 ✂**，避免字体依赖）：两个圆环手柄 + 交叉双刃，`SC=11pt`、0.58 灰、0.5pt 线，右侧镜像。

```python
def scissors(page, x, y, s, color, mirror=False):
    def X(v): return x + s - v if mirror else x + v
    c1 = pymupdf.Point(X(0.20*s), y - 0.17*s)
    c2 = pymupdf.Point(X(0.20*s), y + 0.17*s)
    page.draw_circle(c1, 0.11*s, color=color, width=0.5)
    page.draw_circle(c2, 0.11*s, color=color, width=0.5)
    page.draw_line(c1, pymupdf.Point(X(0.95*s), y + 0.13*s), color=color, width=0.5)
    page.draw_line(c2, pymupdf.Point(X(0.95*s), y - 0.13*s), color=color, width=0.5)
```

- 页图可复用 A5 版的 2480px 导出图（1755×2480，落 A6 上约 420dpi），不必重新导出。
- 文件名：`~/Downloads/<年>年<月>月手账本@改变从整理开始(A6版).pdf`。
- 参考实现：`pptx-booklet-pdf/dec_compose_a6.py`（同目录另有 `dec_compose_b5.py`：真 B5 176×250 拼 A3 横向 420×297，2×B5 居中，K=1.185，用户暂不需要但客户要大开本时可用）。

## 交付与协作约定（踩过的坑，务必遵守）

1. **不要改用户的 PPT 文稿**。用 `slide_find_replace_text` 动文字会重建文本块，把该行的**单行加粗/高亮样式清掉**（12月封底「12月 全年复盘」踩过，用户要求撤销）。文案要改就让用户在编辑器里改，我只负责重新生成 PDF。若已被要求改，还原方式：替换回原文本 + `slide_set_text_property(bold=True)` 补样式。
2. **不要擅自放大/加深水印**。55% 灰、7pt 是参考版规格；用户说「看不清」时先渲染验证颜色（像素 140/255）再解释，别自己改设计。
3. **不覆盖 Downloads 里已有文件**。`cp` 覆盖会被文件策略拦成「删除」请求并可能被拒；一律写新文件名（带 `(A5版)`/`(A6版)` 后缀），旧文件留给用户自己删。
4. **交付前必做的三件事**：① 逐印面渲染核对页码配对；② `page.get_fonts()` 为空、每页图片数对（A5=3、A6=6）；③ `strings | grep AESV3` 确认加密。
5. **每期交付两个文件**：`…(A5版).pdf` 和 `…(A6版).pdf`（默认规则，见顶部速查表）。
6. 每次生成后更新工作日志；用户每月都要一轮，流程与脚本保持可复用。

## 说明

- 9月参考PDF内部实为服务器渲染的 300dpi 位图（1754×2480 Image XObject），非矢量 PPT→PDF；本流程与其规格完全一致。
- 页数不是 4 的倍数时需先补空白页到 4 的倍数再拼版（N 为补后页数）。
- 隐藏页判定细节与代码见「流程 1」，**每期都要查一次**（用户常在中间插入或隐藏一页）。
- `manage.export_progress` 在 progress=50 时可能返回**半截 file_url**（如 `"https:"`），必须等 `progress=='100'` 且 url 以 `http` 开头、长度 > 20 再下载。
- **乱码只能用位图解决**：`page.get_text()` 回读正常 ≠ 用户那里正常（CFF2 字体陷阱）。验收标准是「用图片形式落地」，成品正确时 `page.get_fonts()` 应为空、每页 3 个图片（2 张幻灯片 + 1 张水印）。
- 原版「版权所有 © 2026 改变从整理开始」的宽间距来自雅黑字形边距；PingFang SC 单空格渲染墨迹宽 106.45pt ≈ 原版 106.1pt，视觉一致。
- `/tmp` 会被系统定期清理，重跑前先确认 `/tmp/slide_XX.png`、`/tmp/watermark.png` 是否还在；水印可由 skill 目录里的 `make_wm_png.py` + `PingFangSC-Regular.ttf` 重建。
- 加密方案（用户确认）：只设 owner 密码（限制编辑/复制），user 密码留空——买家双击即可打开、可打印。能防编辑复制，防不住截图。
- skill 目录自带可复用资产：`PingFangSC-Regular.ttf`、`make_wm_png.py`、`compose_v4.py`、`dec_compose_b5.py`、`dec_compose_a6.py`。
