# -*- coding: utf-8 -*-
"""把 PPT 每页的内容整体合并为一个对象（组），再整体等比缩小；
标题文字与页面右侧的公司图标不参与合并/缩放，并置于最上层保证不被遮挡。

用法：
  python group_and_scale.py --file 合并.pptx [--scale 0.85] [--skip-slides 1]
可选：
  --margin-left 0.3 --margin-top 0.8   缩放后的左上留白（英寸）
  --out 输出.pptx                       默认原地保存（会先生成 *_分组前备份.pptx）
  --title-pt 20                         判定标题的最小字号
  --logo-zone 0.88                      判定右侧 Logo：右缘 > 0.88 × 页宽

原理：
  组的 a:xfrm 中 a:ext（屏幕实际尺寸）= scale × a:chExt（原始坐标系尺寸），
  a:xfrm 位于 p:grpSpPr 内部，必须用 .//a:xfrm 取（直接 find 会取到 None）。
  组缩放会连带缩放组内文字，故组内文字与容器始终同比例。
"""
import argparse
import os
import shutil

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EMU_PER_IN = 914400
TOP_TAGS = ("sp", "pic", "grpSp", "graphicFrame", "cxnSp")


def local(el):
    return etree.QName(el.tag).localname


def max_run_pt(shape):
    if not shape.has_text_frame:
        return 0.0
    sizes = [r.font.size.pt for p in shape.text_frame.paragraphs for r in p.runs if r.font.size]
    return max(sizes) if sizes else 0.0


def find_titles(slide, title_pt):
    """标题 = 顶层文本框，且最大字号 >= title_pt 且宽度 > 0.4in；找不到则回退最靠上的宽文本。"""
    tree = slide.shapes._spTree
    cands = []
    for el in tree:
        if local(el) != "sp":
            continue
        shp = slide.shapes._shape_factory(el)
        if not shp.has_text_frame or not shp.text_frame.text.strip():
            continue
        cands.append((max_run_pt(shp), Emu(shp.top).inches, Emu(shp.width).inches, shp))
    titles = [c[3] for c in cands if c[0] >= title_pt and c[2] > 0.4]
    if not titles:
        wide = [c for c in cands if c[2] > 1.5]
        if wide:
            wide.sort(key=lambda c: c[1])
            titles = [wide[0][3]]
    return titles


def find_logos(slide, page_w_in, logo_zone):
    """右侧公司图标 = 顶层图片，且左缘+宽度 > logo_zone × 页宽。"""
    out = []
    for el in slide.shapes._spTree:
        if local(el) != "pic":
            continue
        shp = slide.shapes._shape_factory(el)
        right = Emu(shp.left).inches + Emu(shp.width).inches
        if right > logo_zone * page_w_in:
            out.append(shp)
    return out


def process_slide(slide, prs, scale, margin_left, margin_top, title_pt, logo_zone):
    W_emu, H_emu = prs.slide_width, prs.slide_height
    W_in, H_in = Emu(W_emu).inches, Emu(H_emu).inches
    tree = slide.shapes._spTree

    titles = find_titles(slide, title_pt)
    logos = find_logos(slide, W_in, logo_zone)
    keep = {id(s._element) for s in titles + logos}

    content = [slide.shapes._shape_factory(el) for el in tree
               if local(el) in TOP_TAGS and id(el) not in keep]
    if not content:
        return {"titles": len(titles), "logos": len(logos), "content": 0, "scale": None}

    # 1) 建组并把内容搬进去（lxml 移动语义，保持原有顺序）
    grp = slide.shapes.add_group_shape(content)

    # 2) 收敛缩放比：留白内可容纳的最大比例，且不超过 0.9
    s = min(scale,
            (W_in - margin_left - 0.3) / W_in,
            (H_in - margin_top - 0.3) / H_in,
            0.9)

    # 3) 写组变换：off=留白起点，ext=缩放后尺寸，chOff/chExt=原始整页坐标系
    xfrm = grp._element.find(".//" + NS_A + "xfrm")
    xfrm.find(NS_A + "off").set("x", str(int(margin_left * EMU_PER_IN)))
    xfrm.find(NS_A + "off").set("y", str(int(margin_top * EMU_PER_IN)))
    xfrm.find(NS_A + "ext").set("cx", str(int(W_emu * s)))
    xfrm.find(NS_A + "ext").set("cy", str(int(H_emu * s)))
    xfrm.find(NS_A + "chOff").set("x", "0")
    xfrm.find(NS_A + "chOff").set("y", "0")
    xfrm.find(NS_A + "chExt").set("cx", str(W_emu))
    xfrm.find(NS_A + "chExt").set("cy", str(H_emu))

    # 4) 标题/Logo 移到 spTree 末尾 = 最上层，保证不被组遮挡
    for shp in titles + logos:
        el = shp._element
        tree.remove(el)
        tree.append(el)

    return {"titles": len(titles), "logos": len(logos), "content": len(content), "scale": round(s, 3)}


def main():
    ap = argparse.ArgumentParser(description="整页合并为一个对象并等比缩小")
    ap.add_argument("--file", required=True)
    ap.add_argument("--out", default=None, help="默认原地保存（会先备份）")
    ap.add_argument("--scale", type=float, default=0.85)
    ap.add_argument("--skip-slides", type=int, default=1, help="前 N 页不处理（封面）")
    ap.add_argument("--margin-left", type=float, default=0.3)
    ap.add_argument("--margin-top", type=float, default=0.8)
    ap.add_argument("--title-pt", type=float, default=20.0)
    ap.add_argument("--logo-zone", type=float, default=0.88)
    args = ap.parse_args()

    out = args.out or args.file
    if out == args.file:
        bak = os.path.splitext(args.file)[0] + "_分组前备份.pptx"
        shutil.copyfile(args.file, bak)
        print("已备份 ->", bak)

    prs = Presentation(args.file)
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, 1):
        if i <= args.skip_slides:
            print("第%d页: 跳过（封面/前 %d 页）" % (i, args.skip_slides))
            continue
        info = process_slide(slide, prs, args.scale, args.margin_left, args.margin_top,
                             args.title_pt, args.logo_zone)
        if info["scale"] is None:
            print("第%d页: 无内容可分组，已跳过" % i)
        else:
            print("第%d页: 标题%d个 / Logo%d个 / 内容%d个 -> 合并为1个对象，缩放比=%.3f"
                  % (i, info["titles"], info["logos"], info["content"], info["scale"]))

    prs.save(out)
    print("已保存 %s（共 %d 页）" % (out, total))


if __name__ == "__main__":
    main()
