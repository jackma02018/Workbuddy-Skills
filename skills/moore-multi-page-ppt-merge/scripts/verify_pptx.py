# -*- coding: utf-8 -*-
"""PPT 加工结果校验：图片完整性 / 组结构 / 标题与 Logo 是否在组外顶层 / 文字是否入框 / autofit。

用法：
  python verify_pptx.py --file 结果.pptx --baseline 改动前备份.pptx
可选：
  --title-pt 20 --logo-zone 0.88 --font 黑体

退出码 0 = 全部通过；1 = 有检查项未通过（便于流程中断）。
"""
import argparse
import hashlib
import re
import sys
import unicodedata
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def local(el):
    """兼容传入元素或标签字符串。"""
    tag = el.tag if hasattr(el, "tag") else el
    return etree.QName(tag).localname


def sha1(b):
    return hashlib.sha1(b).hexdigest()


def media_map(path):
    """返回 {partname: sha1}；排除以 / 结尾的 zip 目录占位项（python-pptx 存盘会剔除）。"""
    z = zipfile.ZipFile(path)
    out = {}
    for n in z.namelist():
        if re.search(r"ppt/.*media/.*", n) and not n.endswith("/"):
            out[n] = sha1(z.read(n))
    z.close()
    return out


def char_w(ch):
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return 1.0
    if ch == " ":
        return 0.32
    if ch in ".,:;'\"!|()[]":
        return 0.30
    return 0.55


def main():
    ap = argparse.ArgumentParser(description="PPT 加工结果校验")
    ap.add_argument("--file", required=True)
    ap.add_argument("--baseline", default=None, help="改动前备份，用于图片 sha1 比对")
    ap.add_argument("--title-pt", type=float, default=20.0)
    ap.add_argument("--logo-zone", type=float, default=0.88)
    args = ap.parse_args()

    ok = True
    prs = Presentation(args.file)
    W_in = Emu(prs.slide_width).inches
    H_in = Emu(prs.slide_height).inches
    print("文件 %s  共 %d 页  画布 %.2f×%.2fin" % (args.file, len(prs.slides), W_in, H_in))

    # 1) 图片完整性
    if args.baseline:
        try:
            base = media_map(args.baseline)
            cur = media_map(args.file)
            missing = sorted(k for k in base if k not in cur)
            changed = sorted(k for k in base if k in cur and base[k] != cur[k])
            extra = sorted(k for k in cur if k not in base)
            good = not missing and not changed
            ok &= good
            print("[1] 图片完整性: 结果=%d 张 / 改动前=%d 张  缺失=%d 变更=%d 新增=%d  ->  %s"
                  % (len(cur), len(base), len(missing), len(changed), len(extra), "OK" if good else "FAIL"))
            for k in (missing + changed)[:5]:
                print("      异常: %s" % k)
        except Exception as e:                      # noqa: BLE001
            print("[1] 图片完整性: 无法比对 (%s)" % e)

    # 2~5) 逐页
    for i, slide in enumerate(prs.slides, 1):
        tree = slide.shapes._spTree
        grps = [c for c in tree if local(c.tag) == "grpSp"]
        scale_txt = []
        grp_children = 0
        for g in grps:
            xfrm = g.find(".//" + NS_A + "xfrm")     # a:xfrm 在 p:grpSpPr 内
            e, ce = xfrm.find(NS_A + "ext"), xfrm.find(NS_A + "chExt")
            if e is not None and ce is not None and int(ce.get("cx")):
                scale_txt.append("%.3f" % (int(e.get("cx")) / int(ce.get("cx"))))
            grp_children += sum(1 for _ in g.iter() if local(_.tag) in ("sp", "pic", "grpSp"))

        # 标题 / Logo 是否在组外顶层
        top_titles, top_logos, top_text_shapes = [], [], []
        for c in tree:
            tag = local(c)
            if tag == "sp":
                shp = slide.shapes._shape_factory(c)
                if not shp.has_text_frame or not shp.text_frame.text.strip():
                    continue
                top_text_shapes.append(shp)
                mx = max([r.font.size.pt for p in shp.text_frame.paragraphs for r in p.runs if r.font.size] or [0])
                if mx >= args.title_pt and Emu(shp.width).inches > 0.4:
                    top_titles.append(shp)
            elif tag == "pic":
                shp = slide.shapes._shape_factory(c)
                if Emu(shp.left).inches + Emu(shp.width).inches > args.logo_zone * W_in:
                    top_logos.append(shp)

        # 组内是否混入顶层文本
        grp_txt = set()
        for g in grps:
            for el in g.iter():
                if local(el) == "sp":
                    t = "".join(x.text or "" for x in el.iter(NS_A + "t")).strip()
                    if t:
                        grp_txt.add(t)
        leaked = [s.text_frame.text.strip()[:12] for s in top_text_shapes if s.text_frame.text.strip() in grp_txt]

        # 文字是否入框 + autofit（检查全部顶层文本框）
        over, nofit = [], []
        for shp in top_text_shapes:
            tf = shp.text_frame
            bodyPr = tf._txBody.find(NS_A + "bodyPr")
            wrap = bodyPr.get("wrap", "square")
            lI = int(bodyPr.get("lIns", 91440))
            rI = int(bodyPr.get("rIns", 91440))
            if wrap == "none":
                usable = Emu(shp.width).inches * 72.0 - (lI + rI) / 12700.0
                need = sum(char_w(ch) for p in tf.paragraphs for r in p.runs for ch in r.text)
                mx = max([r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size] or [0])
                need *= mx
                if need > usable > 0:
                    over.append((tf.text.strip()[:12], round(need / usable, 2)))
            if bodyPr.find(NS_A + "normAutofit") is None:
                nofit.append(tf.text.strip()[:12])

        page_ok = not leaked and not over and not nofit
        ok &= page_ok
        print("[2] 第%d页 组=%d 缩放=%s 组内元素=%d  顶层文本=%d(标题%d) 顶层Logo=%d  文字越界=%s autofit缺失=%s 被并入组=%s  -> %s"
              % (i, len(grps), ",".join(scale_txt) or "-", grp_children, len(top_text_shapes), len(top_titles),
                 len(top_logos), over or "无", nofit or "无", leaked or "无", "OK" if page_ok else "FAIL"))

    print("\n总结论: %s" % ("全部通过 ✅" if ok else "存在未通过项 ❌"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
