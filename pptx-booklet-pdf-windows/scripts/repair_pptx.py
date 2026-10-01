# -*- coding: utf-8 -*-
"""修复腾讯导出 PPTX 中的 INT64_MIN/INT32_MIN 哨兵值 xfrm（全新编写）
用法: python repair_pptx.py <src> <dst>
修复策略：把非法 off/ext 替换为同块 chOff/chExt（恒等变换，布局零偏差）
"""
import zipfile, re, sys, io
from xml.dom import minidom

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

src, dst = sys.argv[1], sys.argv[2]

zin = zipfile.ZipFile(src)
data = {n: zin.read(n) for n in zin.namelist() if not n.endswith('/')}

pat = re.compile(
    r'<a:off x="-9223372036854775808" y="-9223372036854775808"/>'
    r'<a:ext cx="-2147483648" cy="-2147483648"/>'
    r'<a:chOff x="(-?\d+)" y="(-?\d+)"/>'
    r'<a:chExt cx="(\d+)" cy="(\d+)"/>'
)

fixed = 0
for n in sorted(data):
    if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
        s = data[n].decode('utf-8')
        def repl(m):
            global fixed
            fixed += 1
            return (f'<a:off x="{m.group(1)}" y="{m.group(2)}"/>'
                    f'<a:ext cx="{m.group(3)}" cy="{m.group(4)}"/>'
                    f'<a:chOff x="{m.group(1)}" y="{m.group(2)}"/>'
                    f'<a:chExt cx="{m.group(3)}" cy="{m.group(4)}"/>')
        s2 = pat.sub(repl, s)
        if s2 != s:
            # 良构校验
            minidom.parseString(s2)
            data[n] = s2.encode('utf-8')
            print(f'  fixed {n}')

# 重打包：[Content_Types].xml 在前
ordered = ['[Content_Types].xml', '_rels/.rels'] + \
          [n for n in data if n not in ('[Content_Types].xml', '_rels/.rels')]
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
    for n in ordered:
        zout.writestr(n, data[n])

print(f'done: fixed {fixed} bad xfrm block(s) -> {dst}')
