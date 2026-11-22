# -*- coding: utf-8 -*-
"""
项目解读_重点需求与节点 —— 渲染主脚本（可复用骨架）
读取 content.json，套用摩尔方案模板 .potx，生成 6 页重点需求 PPT。

用法：
  python build_deck.py --tpl <模板.potx> --content <content.json> [--out <输出.pptx>]
  # out 缺省：<content.json 同目录>/YYYYMMDD_客户简称_重点需求.pptx
  #          自动检测 ~$ 锁文件 → 占用则递增 _vN 版本（不覆盖/不杀进程）

依赖：python-pptx（置于隔离 venv：C:/Users/majunpeng/.workbuddy/binaries/python/envs/default）
"""
import os, sys, json, zipfile, argparse
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.parts.slide import SlidePart

# ---------- 调色板（与摩尔方案模板对齐；换模板时到 inspect_template.py 核对） ----------
DARK  = RGBColor(0x44, 0x54, 0x6A)
BLUE  = RGBColor(0x44, 0x72, 0xC4)
LBLUE = RGBColor(0x5B, 0x9B, 0xD5)
ORANGE= RGBColor(0xED, 0x7D, 0x31)
GREEN = RGBColor(0x70, 0xAD, 0x47)
GRAY  = RGBColor(0xA5, 0xA5, 0xA5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK   = RGBColor(0x22, 0x22, 0x22)
LTBG  = RGBColor(0xF2, 0xF5, 0xFA)
FONT  = "微软雅黑"

def set_font(run, size=12, bold=False, color=INK, name=FONT):
    run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", name)

def add_text(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    first = True
    for item in lines:
        text, size, bold, color = item[0], item[1], item[2], item[3]
        p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        p.alignment = align
        if len(item) > 4 and item[4] is not None: p.space_after = Pt(item[4])
        p.space_before = Pt(0)
        r = p.add_run(); r.text = text; set_font(r, size, bold, color)
    return tb

def style_title(slide, text, sub=None):
    ph = slide.placeholders[0]; ph.text = text
    for p in ph.text_frame.paragraphs:
        for r in p.runs: set_font(r, 26, True, DARK)
    idxs = [pl.placeholder_format.idx for pl in slide.placeholders]
    if sub is not None and 1 in idxs:
        sp = slide.placeholders[1]; sp.text = sub
        for p in sp.text_frame.paragraphs:
            for r in p.runs: set_font(r, 13, False, GRAY)

def get_layout(prs, keywords):
    for kw in keywords:
        for lay in prs.slide_layouts:
            if kw in (lay.name or ""):
                return lay
    return prs.slide_layouts[0]

def clear_sample_slides(prs):
    """彻底删除模板自带示例幻灯片（sldIdLst 引用 + 孤儿 SlidePart 部件），避免 Duplicate name。"""
    xml_slides = prs.slides._sldIdLst
    for sldId in list(xml_slides):
        rId = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        xml_slides.remove(sldId)
        if rId is not None:
            try: prs.part.drop_rel(rId)
            except Exception: pass
    for part in list(prs.part.package.iter_parts()):
        if isinstance(part, SlidePart):
            prs.part.package._parts.discard(part)

def load_template(tpl):
    """复制 .potx 并把 content-type 的 template.main 改 presentation.main，使 python-pptx 可打开。"""
    with zipfile.ZipFile(tpl, "r") as zin:
        names = zin.namelist(); data = {n: zin.read(n) for n in names}
    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    data["[Content_Types].xml"] = ct.encode("utf-8")
    tmp = tpl + "._tmp.pptx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in data.items(): zout.writestr(n, b)
    return tmp

def overflow_scan(prs):
    SW, SH = prs.slide_width, prs.slide_height
    bad = 0
    for si, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            try:
                if sh.left is None: continue
                if sh.left < 0 or sh.top < 0 or (sh.left + sh.width) > SW + 10000 or (sh.top + sh.height) > SH + 10000:
                    bad += 1
                    print(f"  ⚠ 溢出 S{si}: R={Emu(sh.left+sh.width).inches:.2f} B={Emu(sh.top+sh.height).inches:.2f}")
            except Exception: pass
    return bad

# ============================================================
# 渲染 6 页
# ============================================================
def build(prs, C):
    m = C.get("meta", {})
    client = m.get("client", "")

    # Slide 1 封面（标题幻灯片版式）
    s = prs.slides.add_slide(get_layout(prs, ["标题幻灯片"]))
    s.placeholders[0].text = f"{client} MES 项目\n重点需求解析"
    for p in s.placeholders[0].text_frame.paragraphs:
        for r in p.runs: set_font(r, 32, True, DARK)
    if 1 in [pl.placeholder_format.idx for pl in s.placeholders]:
        s.placeholders[1].text = m.get("cover_subtitle", "")
        for p in s.placeholders[1].text_frame.paragraphs:
            for r in p.runs: set_font(r, 15, False, GRAY)
    if m.get("cover_date_line"):
        add_text(s, 0.7, 6.6, 8, 0.4, [(m["cover_date_line"], 11, False, GRAY)])

    # Slide 2 重点需求（标题和内容版式，enumerate 自动编号）
    s = prs.slides.add_slide(get_layout(prs, ["标题和内容"]))
    n_req = len(C.get("requirements", []))
    style_title(s, "一、重点需求", m.get("req_subtitle", f"共 {n_req} 项，编号自动生成"))
    body = s.placeholders[1]; tf = body.text_frame; tf.word_wrap = True
    if tf.paragraphs:
        tf.paragraphs[0].space_after = Pt(8)   # 副标题独占一段，与需求分离
    for idx, req in enumerate(C.get("requirements", []), start=1):
        p = tf.add_paragraph(); p.space_after = Pt(6)
        r1 = p.add_run(); r1.text = f"{idx}　{req.get('title','')}　"; set_font(r1, 13, True, BLUE)
        r2 = p.add_run(); r2.text = req.get("desc", ""); set_font(r2, 11.5, False, INK)

    # Slide 3 主要流程（仅标题版式 + 节点框 + 箭头 + 底部约束卡片）
    s = prs.slides.add_slide(get_layout(prs, ["仅标题"]))
    proc = C.get("process", {})
    style_title(s, "二、主要流程描述（按生产制程顺序）", proc.get("subtitle", ""))
    stages = proc.get("stages", []); n = len(stages)
    margin = 0.4; usable = 13.333 - 2 * margin; box_w = 1.55
    arr_w = (usable - n * box_w) / (n - 1); box_t = 1.55; box_h = 1.55; x = margin
    for i, st in enumerate(stages):
        name = st.get("name", ""); act = st.get("acts", "")
        shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(box_t), Inches(box_w), Inches(box_h))
        shp.fill.solid(); shp.fill.fore_color.rgb = BLUE if i % 2 == 0 else LBLUE
        shp.line.color.rgb = WHITE; shp.line.width = Pt(1); shp.shadow.inherit = False
        tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = name; set_font(r, 12.5, True, WHITE)
        for ln in act.split("\n"):
            p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run(); r2.text = ln; set_font(r2, 9.5, False, WHITE)
        if i < n - 1:
            ax = x + box_w + (arr_w - 0.18) / 2
            arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(ax), Inches(box_t + box_h / 2 - 0.09), Inches(0.18), Inches(0.18))
            arr.fill.solid(); arr.fill.fore_color.rgb = ORANGE; arr.line.fill.background(); arr.shadow.inherit = False
        x += box_w + arr_w
    callouts = proc.get("callouts", []); cx = margin; cw = usable / 3 - 0.15; ct_t = box_t + box_h + 0.45
    for co in callouts:
        box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(ct_t), Inches(cw), Inches(1.6))
        box.fill.solid(); box.fill.fore_color.rgb = LTBG; box.line.color.rgb = GRAY; box.line.width = Pt(0.75); box.shadow.inherit = False
        tf = box.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08); tf.margin_top = Inches(0.06)
        p = tf.paragraphs[0]; r = p.add_run(); r.text = "◆ " + co.get("title", ""); set_font(r, 11.5, True, ORANGE)
        p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = co.get("desc", ""); set_font(r2, 10, False, INK)
        cx += cw + 0.15

    # Slide 4 重要节点（仅标题版式 + 三列表格）
    s = prs.slides.add_slide(get_layout(prs, ["仅标题"]))
    nodes = C.get("nodes", {})
    style_title(s, "三、重要节点描述（重点管控节点）", nodes.get("subtitle", ""))
    rows = nodes.get("rows", []); nr = len(rows) + 1; nc = 3
    tbl_shape = s.shapes.add_table(nr, nc, Inches(0.4), Inches(1.5), Inches(12.53), Inches(5.4))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.7); tbl.columns[1].width = Inches(5.5); tbl.columns[2].width = Inches(4.33)
    tbl.rows[0].height = Inches(0.42)
    for ri in range(1, nr): tbl.rows[ri].height = Inches(0.54)
    for j, h in enumerate(["管控节点", "管控内容", "判定标准"]):
        c = tbl.cell(0, j); c.fill.solid(); c.fill.fore_color.rgb = BLUE
        c.margin_left = Inches(0.06); c.margin_right = Inches(0.06); c.margin_top = Inches(0.03); c.margin_bottom = Inches(0.03)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = h; set_font(r, 11.5, True, WHITE)
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate([row.get("node", ""), row.get("control", ""), row.get("criteria", "")]):
            c = tbl.cell(i, j); c.fill.solid(); c.fill.fore_color.rgb = WHITE if i % 2 == 1 else LTBG
            c.margin_left = Inches(0.06); c.margin_right = Inches(0.06); c.margin_top = Inches(0.03); c.margin_bottom = Inches(0.03)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = v; set_font(r, 10, (j == 0), (DARK if j == 0 else INK))

    # Slide 5 待确认（标题和内容版式）
    s = prs.slides.add_slide(get_layout(prs, ["标题和内容"]))
    style_title(s, "四、待确认事项（资料无依据，标注“待确认”）")
    body = s.placeholders[1]; tf = body.text_frame; tf.word_wrap = True
    first = True
    for t in C.get("todos", []):
        p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        p.space_after = Pt(7); r = p.add_run(); r.text = "• " + t; set_font(r, 12, False, INK)

    # Slide 6 结论与建议（仅标题版式 + 文本框）
    s = prs.slides.add_slide(get_layout(prs, ["仅标题"]))
    style_title(s, "五、结论与分阶段实施建议")
    tb = s.shapes.add_textbox(Inches(0.92), Inches(1.20), Inches(11.5), Inches(5.6))
    tf = tb.text_frame; tf.word_wrap = True; first = True
    for co in C.get("conclusion", []):
        p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        p.space_after = Pt(12)
        r1 = p.add_run(); r1.text = "▶ " + co.get("head", "") + "："; set_font(r1, 13, True, BLUE)
        r2 = p.add_run(); r2.text = co.get("text", ""); set_font(r2, 12, False, INK)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tpl", required=True, help="摩尔方案模板 .potx 路径")
    ap.add_argument("--content", required=True, help="content.json 路径")
    ap.add_argument("--out", default=None, help="输出 pptx（缺省：content.json 同目录/YYYYMMDD_客户简称_重点需求.pptx）")
    args = ap.parse_args()

    with open(args.content, encoding="utf-8") as f:
        C = json.load(f)
    m = C.get("meta", {})

    tmp = load_template(args.tpl)
    try:
        prs = Presentation(tmp)
        clear_sample_slides(prs)
        build(prs, C)

        if args.out:
            candidate = args.out
        else:
            out_dir = os.path.dirname(os.path.abspath(args.content))
            base = f"{m.get('date','')}_{m.get('client','')}_重点需求"; ext = ".pptx"
            lock = os.path.join(out_dir, "~$" + base + ext)
            if os.path.exists(lock):   # WPS 占用：递增版本，不覆盖/不杀进程
                ver = 2
                while os.path.exists(os.path.join(out_dir, f"{base}_v{ver}{ext}")) or \
                      os.path.exists(os.path.join(out_dir, "~$" + base + f"_v{ver}{ext}")):
                    ver += 1
                candidate = os.path.join(out_dir, f"{base}_v{ver}{ext}")
            else:
                candidate = os.path.join(out_dir, base + ext)

        prs.save(candidate)
        with zipfile.ZipFile(candidate) as z:
            bad = z.testzip()
        print("SAVED:", candidate)
        print("ZIP integrity:", "OK" if bad is None else "BAD:" + str(bad))
        print("Slide count:", len(prs.slides))
        ob = overflow_scan(prs)
        print("Overflow:", "无" if ob == 0 else f"{ob} 处")
    finally:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except Exception: pass

if __name__ == "__main__":
    main()
