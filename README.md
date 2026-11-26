# workbuddy-skills

自用 WorkBuddy 技能合集。每个技能一个目录，目录内 `SKILL.md` 是技能定义，`scripts/` 放配套脚本。

## 技能索引

| 技能 | 说明 | 平台 |
|------|------|------|
| [pptx-booklet-pdf-windows](./pptx-booklet-pdf-windows/) | PPT/PPTX → 骑马订装订 PDF，每期固定产出 A5 版（2-up 折页）与 A6 版（4-up 裁切含剪刀裁剪线），含腾讯导出修复、300DPI 高清导出、版权水印、AES 加密 | Windows + PowerPoint |
| [pptx-booklet-pdf](./pptx-booklet-pdf/) | 腾讯文档 PPT（docs.qq.com/slide）→ 骑马钉手账本 PDF，每期固定产出 A5 版（2-up 折页）与 A6 版（4-up 裁切含剪刀裁剪线）两个 PDF，位图版权水印、AES 加密 | macOS + 腾讯文档 MCP |

## 安装

把技能目录复制到 WorkBuddy 的技能目录即可：

```bash
cp -r pptx-booklet-pdf-windows ~/.workbuddy/skills/
```

或直接把本仓库克隆到 `~/.workbuddy/skills/` 下按需取用。
