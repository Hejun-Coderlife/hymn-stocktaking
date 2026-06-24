# -*- coding: utf-8 -*-
"""把「赫眉 / HYMN」用字体轮廓导出为 SVG <path>，生成自包含 logo（不依赖任何字体）。
并自动写入 static/hymn-logo.svg，以及替换 app.py 中的内联 LOGO_SVG。"""
import re, json
from fontTools.ttLib import TTFont
from fontTools.pens.svgPathPen import SVGPathPen

CJK_FONT = "/System/Library/Fonts/STHeiti Medium.ttc"
LAT_FONT = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"

def open_font(path):
    try:
        return TTFont(path, fontNumber=0)
    except Exception:
        return TTFont(path)

def glyph(font, ch):
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    gs = font.getGlyphSet()
    g = gs[cmap[ord(ch)]]
    pen = SVGPathPen(gs)
    g.draw(pen)
    return pen.getCommands(), g.width, upm

def run(font, text, em_px, x0, baseline):
    """返回 (paths_markup, end_x)。每字一个 <path>，已缩放并翻转 y。"""
    parts = []
    x = x0
    for ch in text:
        d, adv, upm = glyph(font, ch)
        s = em_px / upm
        if d.strip():
            parts.append(
                f'<path d="{d}" transform="translate({x:.2f} {baseline:.2f}) scale({s:.5f} {-s:.5f})" fill="#221D17"/>'
            )
        x += adv * s
    return "".join(parts), x

cjk = open_font(CJK_FONT)
lat = open_font(LAT_FONT)

# 布局参数
EM_CJK = 108        # 赫眉 字号(px)
EM_LAT = 80         # HYMN 字号(px)
PADX   = 34         # 边框内左右留白
BASE_CJK = 136      # 赫眉 基线
BASE_LAT = 130      # HYMN 基线
BOX_X, BOX_Y, BOX_H = 14, 38, 124
GAP_AFTER_BOX = 40
RIGHT_PAD = 26

# 先量出赫眉宽度
_, cjk_end_tmp = run(cjk, "赫眉", EM_CJK, 0, 0)
cjk_w = cjk_end_tmp
box_w = cjk_w + 2 * PADX
x0_cjk = BOX_X + PADX
cjk_paths, _ = run(cjk, "赫眉", EM_CJK, x0_cjk, BASE_CJK)

# HYMN
box_right = BOX_X + box_w
x0_lat = box_right + GAP_AFTER_BOX
lat_paths, lat_end = run(lat, "HYMN", EM_LAT, x0_lat, BASE_LAT)

VW = lat_end + RIGHT_PAD
VH = 200

box = f'<rect x="{BOX_X}" y="{BOX_Y}" width="{box_w:.2f}" height="{BOX_H}" rx="2" fill="none" stroke="#221D17" stroke-width="3"/>'
inner = box + cjk_paths + lat_paths
svg = f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {VW:.0f} {VH}" role="img" aria-label="赫眉 HYMN">{inner}</svg>'

# 1) 写 standalone（带换行美化）
pretty = svg.replace("><", ">\n  <").replace("<svg", "<svg").replace("\n  </svg>", "\n</svg>")
open("static/hymn-logo.svg", "w", encoding="utf-8").write(pretty + "\n")

# 2) 替换 app.py 内联 LOGO_SVG
src = open("app.py", encoding="utf-8").read()
new_assign = "const LOGO_SVG = " + json.dumps(svg, ensure_ascii=False) + ";\n"
src2 = re.sub(r"const LOGO_SVG = .*?';\n", new_assign, src, count=1, flags=re.S)
assert src2 != src, "未找到 LOGO_SVG 进行替换"
open("app.py", "w", encoding="utf-8").write(src2)

print(f"viewBox 0 0 {VW:.0f} 200 | 赫眉宽 {cjk_w:.0f} 框宽 {box_w:.0f} | 路径数 cjk={cjk_paths.count('<path')} lat={lat_paths.count('<path')}")
print("✅ 已写 static/hymn-logo.svg 并替换 app.py 内联 LOGO_SVG")
