# -*- coding: utf-8 -*-
"""把多份单页方案 PPT 合并进一份模版 PPT，并统一字体字号（黑体 / 标题20 / 正文12）。

用法：
  python merge_ppts.py --template 模版.pptx --sources A.pptx B.pptx --out 合并.pptx
可选：
  --font 黑体 --title-pt 20 --body-pt 12 --title-threshold 16
  --keep first|all     模版保留首页（默认）或全部页
  --layout match|min    新页使用的版式：按来源版式同名匹配（默认）/ 占位符最少的版式

关键点：
  * 图片走原始字节复制（Part + relate_to），绕过 PIL —— EMF/WMF/SVG 用 PIL 会识别失败；
  * 全文档图片按 sha1 去重，避免同一 Logo 被复制 N 份；
  * 中文字体必须同时写 a:latin / a:ea / a:cs 的 typeface。
"""
import argparse
import copy
import hashlib
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EMU_PER_IN = 914400


def local(el):
    return etree.QName(el.tag).localname


def rel_attrs(el):
    return [a for a in el.attrib if a.startswith(NS_R)]


class PartCopier:
    """把源包内的 Part 以原始字节复制进目标包，按 (reltype, sha1) 去重。"""

    def __init__(self, package):
        self.package = package
        self._names = {str(p.partname) for p in package.iter_parts()}
        self._by_key = {}

    def _alloc(self, ext):
        i = 1
        while True:
            cand = "/ppt/media/image%d%s" % (i, ext)
            if cand not in self._names:
                self._names.add(cand)
                return PackURI(cand)
            i += 1

    def copy(self, dst_part, src_target, reltype, blob):
        key = (reltype, hashlib.sha1(blob).hexdigest())
        if key in self._by_key:
            return self._by_key[key]
        ext = os.path.splitext(str(src_target.partname))[1] or ".bin"
        part = Part(self._alloc(ext), src_target.content_type, self.package, blob)
        rId = dst_part.relate_to(part, reltype)
        self._by_key[key] = rId
        return rId


def choose_layout(dst_prs, src_slide, mode):
    """为新页选择版式：match=按来源版式名匹配，min=占位符最少的版式（等价空白版式）。"""
    src_name = None
    try:
        src_name = src_slide.slide_layout.name
    except Exception:
        src_name = None
    if mode == "match" and src_name:
        for lay in dst_prs.slide_layouts:
            if lay.name == src_name:
                return lay
    best, fewest = None, None
    for lay in dst_prs.slide_layouts:
        n = len(list(lay.placeholders))
        if fewest is None or n < fewest:
            best, fewest = lay, n
    return best or dst_prs.slide_layouts[0]


def clone_slide(src_prs, dst_prs, copier, layout):
    """把源演示文稿第 1 页克隆进目标演示文稿的新页。"""
    src_slide = src_prs.slides[0]
    new_slide = dst_prs.slides.add_slide(layout)
    dst_tree = new_slide.shapes._spTree
    for el in list(dst_tree):
        if local(el) not in ("nvGrpSpPr", "grpSpPr"):
            dst_tree.remove(el)
    for el in src_slide.shapes._spTree:
        if local(el) in ("nvGrpSpPr", "grpSpPr"):
            continue
        dst_tree.append(copy.deepcopy(el))

    src_part, dst_part = src_slide.part, new_slide.part
    mapping = {}
    for el in dst_tree.iter():
        for attr in rel_attrs(el):
            rid = el.get(attr)
            if not rid:
                continue
            if rid not in mapping:
                try:
                    rel = src_part.rels[rid]
                except KeyError:
                    continue
                if rel.is_external:
                    mapping[rid] = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                else:
                    mapping[rid] = copier.copy(dst_part, rel.target_part, rel.reltype, rel.target_part.blob)
            el.set(attr, mapping[rid])
    return new_slide


def keep_first_slide(prs):
    """只保留首页（封面），删除模版其余页。"""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst)[1:]:
        rId = sldId.get(NS_R + "id")
        if rId:
            prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def _insert_ordered(parent, child, before_tags):
    for i, existing in enumerate(parent):
        if local(existing) in before_tags:
            parent.insert(i, child)
            return
    parent.append(child)


def normalize_fonts(slide, font, title_pt, body_pt, threshold):
    """统一字号与字体：原始字号 >= threshold 视为标题；同时写 latin/ea/cs。"""
    stat = {"title": 0, "body": 0}
    for el in slide.shapes._spTree.iter():
        tag = local(el)
        if tag not in ("rPr", "defRPr", "endParaRPr"):
            continue
        sz = el.get("sz")
        if sz is not None:
            try:
                pt = int(sz) / 100.0
            except ValueError:
                pt = None
            if pt is not None and pt > 0:
                if pt >= threshold:
                    el.set("sz", str(int(round(title_pt * 100))))
                    stat["title"] += 1
                else:
                    el.set("sz", str(int(round(body_pt * 100))))
                    stat["body"] += 1
        if tag in ("rPr", "defRPr"):
            for which in ("latin", "ea", "cs"):
                sub = el.find(NS_A + which)
                if sub is None:
                    sub = etree.Element(NS_A + which)
                    _insert_ordered(el, sub, ("sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"))
                sub.set("typeface", font)
    return stat


def main():
    ap = argparse.ArgumentParser(description="合并多份单页方案 PPT 并统一字体")
    ap.add_argument("--template", required=True, help="模版 pptx，首页作封面")
    ap.add_argument("--sources", required=True, nargs="+", help="来源单页 pptx（可多个）")
    ap.add_argument("--out", required=True, help="输出 pptx 路径")
    ap.add_argument("--font", default="黑体")
    ap.add_argument("--title-pt", type=float, default=20.0)
    ap.add_argument("--body-pt", type=float, default=12.0)
    ap.add_argument("--title-threshold", type=float, default=16.0)
    ap.add_argument("--keep", choices=("first", "all"), default="first")
    ap.add_argument("--layout", choices=("match", "min"), default="match")
    args = ap.parse_args()

    for p in [args.template] + args.sources:
        if not os.path.exists(p):
            sys.exit("找不到文件: %s" % p)

    dst_prs = Presentation(args.template)
    if args.keep == "first":
        keep_first_slide(dst_prs)
    copier = PartCopier(dst_prs.part.package)

    print("模版 %s 起始页数=%d" % (os.path.basename(args.template), len(dst_prs.slides)))
    for src in args.sources:
        src_prs = Presentation(src)
        layout = choose_layout(dst_prs, src_prs.slides[0], args.layout)
        slide = clone_slide(src_prs, dst_prs, copier, layout)
        stat = normalize_fonts(slide, args.font, args.title_pt, args.body_pt, args.title_threshold)
        print("  + %-46s 版式=%-14s 标题字号改%d处/正文字号改%d处"
              % (os.path.basename(src), layout.name, stat["title"], stat["body"]))

    dst_prs.save(args.out)
    out = Presentation(args.out)
    media = {str(p.partname) for p in out.part.package.iter_parts() if "/media/" in str(p.partname)}
    print("已保存 %s  共 %d 页  媒体文件 %d 个" % (args.out, len(out.slides), len(media)))


if __name__ == "__main__":
    main()
