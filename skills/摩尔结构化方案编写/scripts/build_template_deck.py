# -*- coding: utf-8 -*-
"""build_template_deck — 基于客户上传的 PPT 模板生成摩尔结构化方案（V2.0）

核心约束（V2.0 不可破）：
  1. 模板的封面 / 目录 / 封底 / 过渡页、标题占位位置、Logo(页眉组)位置 一律不改；
     仅把模板页眉组(含 Logo/色标)逐字节克隆到新增内容页的相同位置。
  2. 内容页复用 deck_kit 的四模块骨架渲染（业务需求/当前痛点/解决方案双视角/方案价值），
     圆点精确对齐文字、价值字号加大、默认不绘页脚数据来源/品牌标注。
  3. 默认零外部素材依赖、矢量可编辑，生成后由 build_deck 统一跑版式体检。

用法：
  python build_template_deck.py --content <内容.json> --template <PPT模板.pptx> \
        --out <输出.pptx> [--dividers <分隔页.json>]
  # 分隔页(过渡页)可选：list of {"before":<1-based内容页序号>, "title":.., "sub":..}
  # 也可在内容 JSON 的 meta.dividers 中声明。
"""
import io
import json
import os
import copy
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deck_kit as K
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

DEFAULT_TPL = "E:/02跟进项目/PPT模板.pptx"


# ---------------- 工具 ----------------
def to_rgb(v):
    if isinstance(v, RGBColor):
        return v
    if isinstance(v, (list, tuple)) and len(v) == 3:
        return RGBColor(int(v[0]), int(v[1]), int(v[2]))
    s = str(v).lstrip('#').strip()
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def norm_page(p, i):
    """与 build_deck 同口径：补骨架、转 RGB、元组化、结构校验。"""
    p = dict(p)
    p['bone'] = p.get('bone') or ['A', 'B', 'C', 'D', 'E'][(i - 1) % 5]
    if p['bone'] not in K.BONES:
        raise ValueError('第 %d 页骨架 %r 不存在，可选 %s' % (i, p['bone'], list(K.BONES)))
    p['need'] = [str(x) for x in p.get('need', [])]
    p['pain'] = [(a, b) for a, b in p.get('pain', [])]
    p['kpis'] = [(a, b) for a, b in p.get('kpis', [])]
    fl = p.get('flow') or {}
    fl['steps'] = [(a, b) for a, b in fl.get('steps', [])]
    p['flow'] = fl
    alt = p.get('alt') or {}
    if 'color' in alt:
        alt['color'] = to_rgb(alt['color'])
    cards = [tuple(list(c) + [''] * (3 - len(c))) for c in alt.get('cards', [])]
    alt['cards'] = cards
    p['alt'] = alt
    icons = list(p.get('flow_icons') or [])
    while len(icons) < len(fl['steps']):
        icons.append('check')
    p['flow_icons'] = icons[:len(fl['steps'])] or ['check']
    # 兼容旧字段：claim
    p.setdefault('claim', p.get('claim', ''))
    missing = [k for k in ('title', 'sub', 'need', 'pain', 'flow', 'alt', 'kpis')
               if not p.get(k)]
    if missing:
        raise ValueError('第 %d 页缺字段：%s' % (i, missing))
    if len(p['kpis']) not in (3, 4):
        raise ValueError('第 %d 页 kpis 需 3 或 4 项，当前 %d' % (i, len(p['kpis'])))
    return p


def clone_group(src_shape):
    """深拷贝组合(含子元素)，重置内部 id 避免冲突。"""
    el = copy.deepcopy(src_shape._element)
    c = [9000]
    for node in el.iter():
        tag = node.tag.split('}')[-1]
        if tag in ('sp', 'pic', 'grpSp', 'graphicFrame'):
            cn = node.find(qn('p:cNvPr'))
            if cn is not None and 'id' in cn.attrib:
                c[0] += 1
                cn.set('id', str(c[0]))
    return el


def find_header_group(slide):
    """取模板页眉/Logo 组合（取顶部最接近 0.75in 的组）。"""
    cands = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.GROUP]
    if not cands:
        return None
    cands.sort(key=lambda s: abs(s.top / 914400 - 0.75))
    return cands[0]


def find_source_slide(prs):
    """选一个含「标题占位 + 页眉组」的内容示例页作为克隆源；找不到则退化为第3页。"""
    slides = list(prs.slides)
    for idx in range(2, len(slides) - 1):
        sl = slides[idx]
        has_ph = any(sh.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER for sh in sl.shapes)
        has_grp = any(sh.shape_type == MSO_SHAPE_TYPE.GROUP for sh in sl.shapes)
        if has_ph and has_grp:
            return idx
    return 2 if len(slides) > 2 else 0


def find_title_ph(slide):
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            return sh
    return None


def add_title(slide, title, ph):
    """把标题写入模板标题占位的相同坐标（位置不变，仅替换文字）。"""
    box = slide.shapes.add_textbox(Inches(ph.left / 914400), Inches(ph.top / 914400),
                                    Inches(ph.width / 914400), Inches(ph.height / 914400))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = K.INK
    r.font.name = K.FONT
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        rPr.append(__import__('pptx').oxml.parse_xml(
            '<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="%s"/>' % K.FONT))
    else:
        ea.set('typeface', K.FONT)
    return box


def add_divider(slide, header_el, title, sub):
    """专题分隔/过渡页：克隆模板页眉组(含Logo)，居中写专题名。"""
    if header_el is not None:
        slide.shapes._spTree.append(copy.deepcopy(header_el))
    box = slide.shapes.add_textbox(Inches(0.5), Inches(2.7), Inches(12.333), Inches(1.1))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.size = Pt(34)
    r.font.bold = True
    r.font.color.rgb = K.INK
    r.font.name = K.FONT
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = sub
    r2.font.size = Pt(13)
    r2.font.color.rgb = K.GREY
    r2.font.name = K.FONT
    K.footer(slide)


def delete_slide(prs, index):
    """移除一页的 sldId 引用与关系（物理 part 作唯一命名的孤儿保留，不冲突）。"""
    sldIdLst = prs.slides._sldIdLst
    slide_id = list(sldIdLst)[index]
    rId = slide_id.get(qn('r:id'))
    sldIdLst.remove(slide_id)
    prs.part.drop_rel(rId)


# ---------------- 主构建 ----------------
def build(TPL, CONTENT, OUT, dividers=None, meta=None):
    if not os.path.exists(TPL):
        raise FileNotFoundError('模板不存在: %s' % TPL)
    prs = Presentation(TPL)
    src_idx = find_source_slide(prs)
    src_slide = list(prs.slides)[src_idx]
    header_el = clone_group(find_header_group(src_slide)) if find_header_group(src_slide) else None
    ph = find_title_ph(src_slide)

    cfg = json.load(io.open(CONTENT, encoding='utf-8'))
    meta = cfg.get('meta', {}) if meta is None else meta
    pages = [norm_page(p, i) for i, p in enumerate(cfg['pages'], 1)]
    for k in ('source', 'brand', 'corner', 'total', 'footer_anno'):
        if k in meta:
            K.META[k] = meta[k]
    K.META['total'] = len(pages)

    # 分隔页(过渡页)声明：--dividers 优先，否则 meta.dividers
    divs = []
    if dividers:
        divs = dividers
    elif 'dividers' in meta:
        divs = meta['dividers']
    divs = sorted([(int(d['before']), d.get('title', ''), d.get('sub', ''))
                   for d in divs if str(d.get('before', '')).isdigit()])

    orig = list(prs.slides._sldIdLst)
    n0 = len(orig)                      # 模板原始页数（新增页之前）
    cover_id, toc_id, back_id = orig[0], orig[1], orig[-1]
    blank = prs.slide_layouts[2]
    new_ids = []
    di = 0
    for idx, p in enumerate(pages):
        # 在该内容页之前插入对应专题分隔页
        while di < len(divs) and divs[di][0] == idx + 1:
            sl = prs.slides.add_slide(blank)
            add_divider(sl, header_el, divs[di][1], divs[di][2])
            new_ids.append(list(prs.slides._sldIdLst)[-1])
            di += 1
        sl = prs.slides.add_slide(blank)
        if header_el is not None:
            sl.shapes._spTree.append(copy.deepcopy(header_el))
        if ph is not None:
            add_title(sl, p['title'], ph)
        K.BONES[p['bone']][0](sl, p)
        K.footer(sl)
        new_ids.append(list(prs.slides._sldIdLst)[-1])

    # 删除模板自带示例内容页（保留封面/目录/封底）；范围基于模板原始页数 n0，
    # 必须在新增页之前计算，否则会误删新增内容页与封底。
    for i in sorted(range(2, max(2, n0 - 1)), reverse=True):
        delete_slide(prs, i)

    # 重排：封面 → 目录 → 新页(顺序) → 封底
    xml = prs.slides._sldIdLst
    for el in list(xml):
        xml.remove(el)
    xml.append(cover_id)
    xml.append(toc_id)
    for nid in new_ids:
        xml.append(nid)
    xml.append(back_id)

    os.makedirs(os.path.dirname(os.path.abspath(OUT)), exist_ok=True)
    prs.save(OUT)
    return len(pages)


def main():
    args = sys.argv[1:]
    if '--template' in args:
        TPL = args[args.index('--template') + 1]
    else:
        TPL = DEFAULT_TPL
    if '--content' not in args:
        print('缺少 --content <内容.json>')
        return
    CONTENT = args[args.index('--content') + 1]
    OUT = None
    if '--out' in args:
        OUT = args[args.index('--out') + 1]
    else:
        OUT = os.path.splitext(os.path.abspath(CONTENT))[0] + '_模板版.pptx'
    dividers = None
    if '--dividers' in args:
        dv = args[args.index('--dividers') + 1]
        dividers = json.load(io.open(dv, encoding='utf-8')) if os.path.exists(dv) else json.loads(dv)
    n = build(TPL, CONTENT, OUT, dividers=dividers)
    print('SAVED', OUT, 'content_pages=', n)


if __name__ == '__main__':
    main()
