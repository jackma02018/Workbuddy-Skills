# -*- coding: utf-8 -*-
"""文字自适应容器：超出容器时自动缩小字号，仍超出则减少字数（补 …），
并写入 a:normAutofit，令后续在 PowerPoint/WPS 中改字时也会自动缩小。

用法：
  python fit_text.py --file 目标.pptx [--min-pt 10] [--logo-gap 0.2] [--scope top|all]
可选：
  --scope top   只处理顶层文本框（默认，= 标题/副标题等固定位元素）
  --scope all   顶层 + 组内全部文本框（谨慎：原始海报设计本就贴合，可能过度缩字）
  --logo-gap 0.2        标题容器右缘与最左 Logo 左缘的最小间距（英寸）
  --logo-zone 0.88      判定右侧 Logo：右缘 > 0.88 × 页宽
  --out 输出.pptx       默认原地保存（会先备份 *_文字自适应前备份.pptx）

处理顺序（固定）：
  1) 容器宽度封顶（不压 Logo） 2) 高度保底（消除默认内边距造成的假溢出）
  3) 横向溢出按比例缩字号，下限 --min-pt  4) 仍溢出则尾部删字补 …
  5) 写入 a:normAutofit
"""
import argparse
import math
import os
import shutil
import unicodedata

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EMU_PER_IN = 914400
EMU_PER_PT = 12700


def local(el):
    return etree.QName(el.tag).localname


def char_w(ch):
    """粗略字宽（em）：全角 1.0；空格 0.32；窄标点 0.30；其余 0.55。"""
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return 1.0
    if ch == " ":
        return 0.32
    if ch in ".,:;'\"!|()[]":
        return 0.30
    return 0.55


def est_width_pt(tf):
    total = 0.0
    for p in tf.paragraphs:
        for r in p.runs:
            sz = r.font.size.pt if r.font.size else 18.0
            total += sum(char_w(c) for c in r.text) * sz
    return total


def insets(bodyPr):
    return (int(bodyPr.get("lIns", 91440)), int(bodyPr.get("rIns", 91440)),
            int(bodyPr.get("tIns", 45720)), int(bodyPr.get("bIns", 45720)))


def set_autofit(bodyPr):
    """写入 normAutofit（位于 prstTxWarp 之后、scene3d 之前）。"""
    for tag in ("normAutofit", "spAutoFit", "noAutofit"):
        for el in bodyPr.findall(NS_A + tag):
            bodyPr.remove(el)
    fit = etree.Element(NS_A + "normAutofit")
    idx = len(bodyPr)
    for i, child in enumerate(list(bodyPr)):
        if local(child) in ("scene3d", "sp3d", "flatTx", "extLst"):
            idx = i
            break
    bodyPr.insert(idx, fit)


def has_autofit(bodyPr):
    return bodyPr.find(NS_A + "normAutofit") is not None


def fit_text_shape(shape, min_pt, logo_left_in, logo_gap):
    tf = shape.text_frame
    bodyPr = tf._txBody.find(NS_A + "bodyPr")
    lI, rI, tI, bI = insets(bodyPr)
    wrap = bodyPr.get("wrap", "square")
    note = []

    # 1) 宽度封顶：不越过 Logo 区
    if logo_left_in is not None:
        right_in = Emu(shape.left).inches + Emu(shape.width).inches
        cap_right = logo_left_in - logo_gap
        if right_in > cap_right and cap_right > Emu(shape.left).inches + 1.0:
            shape.width = Emu(int((cap_right - Emu(shape.left).inches) * EMU_PER_IN))
            note.append("容器右缘收窄至%.2fin" % cap_right)

    # 2) 高度保底
    maxsz = max([r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size] or [18.0])
    need_h = int((1.25 * maxsz + (tI + bI) / EMU_PER_PT) * EMU_PER_PT)
    if shape.height < need_h:
        shape.height = need_h
        note.append("容器加高")

    # 3) 横向溢出 → 缩字号（向下取整到 0.1pt，绝不上取，否则会因临界值触发多余减字）
    changed = False
    TOL = 1.005          # 0.5% 容差：估算误差内不算溢出，避免"差一点点就减字"
    if wrap == "none":
        usable_pt = Emu(shape.width).inches * 72.0 - (lI + rI) / EMU_PER_PT
        w = est_width_pt(tf)
        if w > usable_pt * TOL > 0:
            factor = usable_pt / w
            new_pt = max(min_pt, math.floor(maxsz * factor * 10) / 10.0)
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(new_pt)
            note.append("字号 %.1f→%.1fpt" % (maxsz, new_pt))
            changed = True
            # 4) 仍溢出 → 减字（预留省略号宽度后再收字）
            w = est_width_pt(tf)
            if w > usable_pt * TOL:
                runs = [r for p in tf.paragraphs for r in p.runs]
                original = tf.text
                target = usable_pt - char_w("…") * new_pt
                while runs and w > target and sum(len(r.text) for r in runs) > 4:
                    for r in reversed(runs):
                        if r.text:
                            r.text = r.text[:-1]
                            break
                    for r in reversed(runs):
                        if r.text.endswith(" "):
                            r.text = r.text[:-1]
                            break
                    w = est_width_pt(tf)
                if runs:
                    runs[-1].text = runs[-1].text.rstrip() + "…"
                note.append("减字：%s → %s" % (original[:16], tf.text[:16]))

    # 5) 写入 autofit
    if not has_autofit(bodyPr):
        set_autofit(bodyPr)
        if not changed:
            note.append("已启用 normAutofit")
    return note


def iter_shapes(slide, scope):
    tree = slide.shapes._spTree
    if scope == "top":
        for el in tree:
            if local(el) == "sp":
                yield slide.shapes._shape_factory(el)
        return
    for el in tree.iter():
        if local(el) == "sp":
            try:
                yield slide.shapes._shape_factory(el)
            except Exception:
                continue


def logo_left_of(slide, page_w_in, logo_zone):
    best = None
    for el in slide.shapes._spTree:
        if local(el) != "pic":
            continue
        shp = slide.shapes._shape_factory(el)
        if Emu(shp.left).inches + Emu(shp.width).inches > logo_zone * page_w_in:
            l = Emu(shp.left).inches
            best = l if best is None else min(best, l)
    return best


def main():
    ap = argparse.ArgumentParser(description="文字自适应容器")
    ap.add_argument("--file", required=True)
    ap.add_argument("--out", default=None, help="默认原地保存（会先备份）")
    ap.add_argument("--scope", choices=("top", "all"), default="top")
    ap.add_argument("--min-pt", type=float, default=10.0)
    ap.add_argument("--logo-gap", type=float, default=0.2)
    ap.add_argument("--logo-zone", type=float, default=0.88)
    args = ap.parse_args()

    out = args.out or args.file
    if out == args.file:
        bak = os.path.splitext(args.file)[0] + "_文字自适应前备份.pptx"
        shutil.copyfile(args.file, bak)
        print("已备份 ->", bak)

    prs = Presentation(args.file)
    page_w_in = Emu(prs.slide_width).inches
    touched = 0
    for i, slide in enumerate(prs.slides, 1):
        logo_left = logo_left_of(slide, page_w_in, args.logo_zone)
        for shp in iter_shapes(slide, args.scope):
            if not shp.has_text_frame or not shp.text_frame.text.strip():
                continue
            note = fit_text_shape(shp, args.min_pt, logo_left, args.logo_gap)
            if note:
                touched += 1
                print("第%d页 '%s' -> %s" % (i, shp.text_frame.text.strip()[:14], "；".join(note)))
    prs.save(out)
    print("已保存 %s；共处理 %d 处文本容器" % (out, touched))


if __name__ == "__main__":
    main()
