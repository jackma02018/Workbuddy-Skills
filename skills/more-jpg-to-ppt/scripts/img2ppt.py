#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
img2ppt.py —— 把图片文件夹转成「一页一图」的 PPTX。

用法:
    python img2ppt.py <图片文件夹> [输出.pptx] [选项]

选项:
    --size 16:9|4:3        幻灯片比例（默认 16:9）
    --background #RRGGBB   背景色（默认 FFFFFF）
    --margin 英寸          图片与页面边缘留白（默认 0.0）
    --page-numbers         每页右下角显示 n / N 页码
    --title 文本           首张图片页前插入一张标题页

行为:
    - 图片按文件名「自然顺序」排序（img2 排在 img10 之前）
    - 只处理顶层文件，不递归子文件夹
    - 图片始终等比放大铺满页面（等价 PowerPoint「适应幻灯片」），不裁切
    - webp 在 python-pptx 原生不支持，自动转 PNG 后插入（临时文件用后即删）
    - 以 . 开头的隐藏文件跳过
    - 损坏图片打警告并跳过该页，不中断整体转换
    - 不传输出路径时，默认输出到文件夹同级、与文件夹同名的 .pptx

依赖:
    python-pptx, pillow
"""

import argparse
import os
import re
import sys
import tempfile

try:
    from pptx import Presentation
    from pptx.util import Inches, Emu
    from PIL import Image
except ImportError as e:
    sys.exit(f"缺少依赖: {e}\n请先安装: pip install python-pptx pillow")

SUPPORTED = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
# python-pptx 不认 webp，需要先转 PNG
NEEDS_CONVERT = {".webp"}

SIZES = {
    "16:9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
}


def natural_key(name):
    """自然排序：img2 排在 img10 之前。"""
    return [int(t) if t.isdigit() else t.lower()
            for t in re.split(r"(\d+)", name)]


def collect_images(folder):
    files = []
    for entry in sorted(os.listdir(folder)):
        if entry.startswith("."):
            continue
        full = os.path.join(folder, entry)
        if not os.path.isfile(full):
            continue
        if os.path.splitext(entry)[1].lower() in SUPPORTED:
            files.append(full)
    files.sort(key=lambda p: natural_key(os.path.basename(p)))
    return files


def normalize_color(value, default="FFFFFF"):
    if not value:
        return default
    v = value.lstrip("#").strip().upper()
    if len(v) == 3:
        v = "".join(c * 2 for c in v)
    if len(v) != 6 or any(c not in "0123456789ABCDEF" for c in v):
        raise ValueError(f"非法颜色值: {value}")
    return v


def fit_box(img_w_in, img_h_in, page_w, page_h, margin):
    """等比缩放并居中，返回 (left, top, width, height)。"""
    avail_w = page_w - 2 * margin
    avail_h = page_h - 2 * margin
    if avail_w <= 0 or avail_h <= 0:
        raise ValueError("留白过大，页面可用区域为负")
    scale = min(avail_w / img_w_in, avail_h / img_h_in)
    w = img_w_in * scale
    h = img_h_in * scale
    left = margin + (avail_w - w) / 2
    top = margin + (avail_h - h) / 2
    return left, top, w, h


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("folder", help="图片文件夹")
    ap.add_argument("out", nargs="?", help="输出 .pptx 路径（可选）")
    ap.add_argument("--size", default="16:9", choices=list(SIZES.keys()))
    ap.add_argument("--background", default="FFFFFF")
    ap.add_argument("--margin", type=float, default=0.0)
    ap.add_argument("--page-numbers", action="store_true")
    ap.add_argument("--title", default=None)
    args = ap.parse_args()

    folder = os.path.abspath(args.folder)
    if not os.path.isdir(folder):
        sys.exit(f"不是文件夹: {folder}")

    images = collect_images(folder)
    if not images:
        sys.exit(f"未找到支持的图片: {folder}\n支持格式: {', '.join(sorted(SUPPORTED))}")

    out = args.out or (folder.rstrip(os.sep) + ".pptx")
    out = os.path.abspath(out)
    bg = normalize_color(args.background)
    page_w, page_h = SIZES[args.size]

    prs = Presentation()
    prs.slide_width = Inches(page_w)
    prs.slide_height = Inches(page_h)
    blank = prs.slide_layouts[6]
    tmp_files = []
    skipped = []

    def paint(slide):
        fill = slide.background.fill
        fill.solid()
        from pptx.dml.color import RGBColor
        fill.fore_color.rgb = RGBColor.from_string(bg)

    total = len(images) + (1 if args.title else 0)
    n = 0

    # 标题页（可选）
    if args.title:
        n += 1
        slide = prs.slides.add_slide(blank)
        paint(slide)
        box = slide.shapes.add_textbox(Inches(0.8), Inches(page_h / 2 - 0.8),
                                       Inches(page_w - 1.6), Inches(1.6))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = 2  # center
        run = p.add_run()
        run.text = args.title
        run.font.size = Inches(0.55)
        run.font.bold = True
        from pptx.dml.color import RGBColor
        run.font.color.rgb = RGBColor.from_string("1F3A5F")

    for idx, path in enumerate(images, 1):
        n += 1
        src = path
        ext = os.path.splitext(path)[1].lower()

        # webp → PNG
        if ext in NEEDS_CONVERT:
            try:
                im = Image.open(path).convert("RGBA")
                fd, tmp = tempfile.mkstemp(suffix=".png")
                os.close(fd)
                im.save(tmp, "PNG")
                tmp_files.append(tmp)
                src = tmp
            except Exception as e:
                print(f"  [!] 跳过（webp 转换失败）: {os.path.basename(path)} — {e}")
                skipped.append(os.path.basename(path))
                continue

        # 读尺寸
        try:
            with Image.open(src) as im:
                iw, ih = im.size
        except Exception as e:
            print(f"  [!] 跳过（无法读取）: {os.path.basename(path)} — {e}")
            skipped.append(os.path.basename(path))
            continue

        if iw <= 0 or ih <= 0:
            print(f"  [!] 跳过（尺寸异常）: {os.path.basename(path)}")
            skipped.append(os.path.basename(path))
            continue

        # DPI 信息不全时按 96dpi 折算英寸
        try:
            dpi = Image.open(src).info.get("dpi", (96, 96))
            dpi_x = dpi[0] if dpi and dpi[0] else 96
            dpi_y = dpi[1] if dpi and dpi[1] else 96
        except Exception:
            dpi_x = dpi_y = 96

        img_w_in = iw / dpi_x
        img_h_in = ih / dpi_y

        left, top, w, h = fit_box(img_w_in, img_h_in, page_w, page_h, args.margin)

        slide = prs.slides.add_slide(blank)
        paint(slide)
        slide.shapes.add_picture(src, Inches(left), Inches(top),
                                 Inches(w), Inches(h))
        print(f"  [{n - (1 if args.title else 0)}/{len(images)}] {os.path.basename(path)}")

        if args.page_numbers:
            tb = slide.shapes.add_textbox(
                Inches(page_w - 1.5), Inches(page_h - 0.6), Inches(1.2), Inches(0.4))
            p = tb.text_frame.paragraphs[0]
            p.alignment = 3  # right
            run = p.add_run()
            run.text = f"{n} / {total}"
            run.font.size = Inches(0.14)
            from pptx.dml.color import RGBColor
            run.font.color.rgb = RGBColor.from_string("808080")

    prs.save(out)

    # 清理临时文件
    for t in tmp_files:
        try:
            os.remove(t)
        except OSError:
            pass

    added = len(images) - len(skipped)
    print(f"\n完成! 共 {added} 张图片 → {out}")
    print(f"幻灯片: {added + (1 if args.title else 0)} 页, 比例 {args.size}")
    if skipped:
        print(f"跳过 {len(skipped)} 个: {', '.join(skipped)}")


if __name__ == "__main__":
    main()
