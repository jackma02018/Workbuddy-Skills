# -*- coding: utf-8 -*-
"""deck_kit — 摩尔元数结构化方案 PPT 引擎（矢量可编辑 · 无第三方素材依赖）

结构：① 设计令牌  ② 12 栏栅格  ③ 基础图元  ④ 组件库  ⑤ 26 个矢量图标
      ⑥ 五套页面骨架 page1..page5(slide, content_dict)
内容与引擎解耦：内容以 JSON 传入 build_deck.py，本文件不改动。
"""

"""
设计语言（全页统一，不随页面变化）
  画布 13.333 x 7.5 in（16:9）   栅格 12 栏 / 页边距 0.50 / 栏间距 0.16
  顶部四色标识条 0.10 + 页眉带 1.05（导航位置固定）
  卡片组件：圆角 0.035 / 白底 / 0.75pt 边框 / 0.055 左侧色条 / 0.045 顶部色线
           标题 12.5pt + 英文 7.5pt / 0.68 分隔线
  字体：微软雅黑；正文 8.5~9.5pt；页脚 7.5pt
  模块色：需求 #1F6FEB / 痛点 #D9534F / 方案 #0B7285 / 价值 #D97706

五页骨架（结构互不相同）
  P1 右重双栏 5:7 + 底部通栏价值带
  P2 顶部通栏方案泳道 + 底部三栏 5:4:3
  P3 顶部需求条 + 中部 4:8 + 底部 8:4
  P4 双塔竖切 左三段 : 右通高 = 6:6
  P5 三竖栏聚焦 3:6:3 + 底部反色价值带
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# ---------------- 设计令牌 ----------------
BG      = RGBColor(0xF4, 0xF6, 0xF9)
INK     = RGBColor(0x0F, 0x27, 0x40)
NAVY    = RGBColor(0x14, 0x33, 0x54)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GREY    = RGBColor(0x6B, 0x7A, 0x8C)
GREY2   = RGBColor(0x9A, 0xA7, 0xB4)
LINE    = RGBColor(0xE4, 0xE9, 0xF0)
SEPL    = RGBColor(0xEE, 0xF1, 0xF5)
SUBTLE  = RGBColor(0x7E, 0x9B, 0xB8)
TXT     = RGBColor(0x33, 0x42, 0x52)
PANEL   = RGBColor(0xFA, 0xFB, 0xFC)

C_NEED  = RGBColor(0x1F, 0x6F, 0xEB)
C_PAIN  = RGBColor(0xD9, 0x53, 0x4F)
C_SOLVE = RGBColor(0x0B, 0x72, 0x85)
C_VALUE = RGBColor(0xD9, 0x77, 0x06)

FONT = '微软雅黑'
SW, SH = 13.333, 7.5

# ---------------- 栅格 ----------------
MARGIN, GUT = 0.50, 0.16
CONT_W = SW - 2 * MARGIN          # 12.333
COL_W = (CONT_W - 11 * GUT) / 12  # 0.881


def cx(c):                        # 第 c 栏左边界（0 起）
    return MARGIN + c * (COL_W + GUT)


def cw(span):                     # 跨 span 栏的宽度
    return span * COL_W + (span - 1) * GUT


TOP, BOT = 1.36, 6.96
BODY_H = BOT - TOP                # 5.60

# ---------------- 基础组件 ----------------
def rect(s, x, y, w, h, fill=None, line=None, lw=1.0,
         shape=MSO_SHAPE.RECTANGLE, adj=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if adj is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = adj
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    if sp.has_text_frame:
        sp.text_frame.word_wrap = True
    return sp


def shadow(sp, color='233A52', blur=12, dist=2, alpha=14000):
    spPr = sp._element.spPr
    old = spPr.find(qn('a:effectLst'))
    if old is not None:
        spPr.remove(old)
    spPr.append(parse_xml(
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:outerShdw blurRad="%d" dist="%d" dir="5400000" algn="tl" rotWithShape="0">'
        '<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>'
        % (int(blur * 12700), int(dist * 12700), color, alpha)))


def tb(s, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return box, tf


def _ea(r):
    r.font.name = FONT
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        rPr.append(parse_xml(
            '<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="%s"/>' % FONT))
    else:
        ea.set('typeface', FONT)


def _para(tf, first=False):
    if first or (len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs):
        return tf.paragraphs[0]
    return tf.add_paragraph()


def run(tf, text, size=9, color=GREY, bold=False, first=False,
        space_after=0, align=PP_ALIGN.LEFT, line=1.0):
    p = _para(tf, first)
    p.alignment = align
    p.space_after = Pt(space_after)
    p.line_spacing = line
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = color
    _ea(r)
    return p


def rich(tf, parts, first=False, space_after=0, line=1.0, align=PP_ALIGN.LEFT):
    """parts = [(text, color, bold, size), ...]"""
    p = _para(tf, first)
    p.alignment = align
    p.space_after = Pt(space_after)
    p.line_spacing = line
    for t, c, b, sz in parts:
        r = p.add_run(); r.text = t
        r.font.size = Pt(sz); r.font.bold = b; r.font.color.rgb = c
        _ea(r)
    return p


def card(s, x, y, w, h, color, title, en, head=0.68, dark=False):
    bg = NAVY if dark else WHITE
    bd = NAVY if dark else LINE
    c = rect(s, x, y, w, h, fill=bg, line=bd, lw=0.75,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.035)
    shadow(c)
    rect(s, x + 0.20, y + 0.19, 0.055, 0.30, fill=color)
    rect(s, x + 0.02, y, w - 0.04, 0.045, fill=color)
    tc = WHITE if dark else INK
    ec = SUBTLE if dark else GREY2
    _, f = tb(s, x + 0.36, y + 0.15, w - 0.75, 0.40)
    run(f, title, size=12.5, color=tc, bold=True, first=True)
    _, f = tb(s, x + 0.36, y + 0.42, w - 0.75, 0.22)
    run(f, en, size=7.5, color=ec, first=True)
    if head:
        rect(s, x + 0.22, y + head, w - 0.44, 0.008,
             fill=RGBColor(0x2A, 0x44, 0x62) if dark else SEPL)
    return c


def dot(s, x, y, color, size=0.075):
    rect(s, x, y, size, size, fill=color, shape=MSO_SHAPE.OVAL).shadow.inherit = False


# ---------- 原生矢量图标库（语义对齐 ima「2025 图标 部分流程」素材体系） ----------
def _g(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, adj=None, rot=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if adj is not None:
        try:
            sp.adjustments[0] = adj
        except Exception:
            pass
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    sp.shadow.inherit = False
    if rot:
        sp.rotation = rot
    return sp


def glyph(s, x, y, sz, kind, c, bg):
    """在 (x,y) 处绘制 sz×sz 的线性剪影图标；c=主色，bg=底色（用于挖空）"""
    U = sz
    P = lambda rx, ry, rw, rh, f=None, sh=MSO_SHAPE.RECTANGLE, a=None, r=None: \
        _g(s, x + rx * U, y + ry * U, rw * U, rh * U, f or c, sh, a, r)
    O = lambda rx, ry, rw, rh, f=None: \
        _g(s, x + rx * U, y + ry * U, rw * U, rh * U, f or c, MSO_SHAPE.OVAL)

    if kind == 'doc':                                    # 工单 / 单据
        P(.22, .10, .56, .80, None, MSO_SHAPE.ROUNDED_RECTANGLE, .10)
        P(.32, .26, .36, .06, bg); P(.32, .42, .36, .06, bg); P(.32, .58, .24, .06, bg)
    elif kind == 'split':                                # 工单拆解
        P(.04, .38, .34, .24, None, MSO_SHAPE.ROUNDED_RECTANGLE, .20)
        P(.44, .14, .52, .22, None, MSO_SHAPE.RIGHT_ARROW)
        P(.44, .64, .52, .22, None, MSO_SHAPE.RIGHT_ARROW)
    elif kind in ('check', 'ok'):                        # 校验 / 放行（像素阶梯勾，免旋转）
        for px_, py_ in ((.16, .46), (.27, .57), (.38, .68),
                         (.49, .57), (.60, .46), (.71, .35)):
            P(px_, py_, .14, .14, None, MSO_SHAPE.ROUNDED_RECTANGLE, .25)
    elif kind == 'people':                               # 人员 / 排班
        O(.33, .06, .34, .34)
        P(.18, .46, .64, .44, None, MSO_SHAPE.ROUNDED_RECTANGLE, .35)
    elif kind == 'scan':                                 # 扫码 / 采集
        P(.10, .16, .80, .68, None, MSO_SHAPE.ROUNDED_RECTANGLE, .10)
        for i in (.22, .38, .54, .70):
            P(i, .28, .06, .44, bg)
    elif kind == 'gear':                                 # 设备 / 参数 / 建模
        P(.08, .08, .84, .84, None, MSO_SHAPE.GEAR_6)
        O(.36, .36, .28, .28, bg)
    elif kind == 'chip':                                 # 程序 / 参数下发
        P(.26, .26, .48, .48, None, MSO_SHAPE.ROUNDED_RECTANGLE, .12)
        for i in (.34, .48, .62):
            P(i - .03, .14, .06, .12); P(i - .03, .74, .06, .12)
            P(.14, i - .03, .12, .06); P(.74, i - .03, .12, .06)
    elif kind == 'arrow':                                # 流转 / 回写 / 派单
        P(.04, .28, .92, .44, None, MSO_SHAPE.RIGHT_ARROW)
    elif kind == 'down':                                 # 下发 / 上料
        P(.26, .04, .48, .92, None, MSO_SHAPE.DOWN_ARROW)
    elif kind == 'clock':                                # 工时 / ECT
        O(.08, .08, .84, .84); O(.18, .18, .64, .64, bg)
        P(.46, .22, .08, .30); P(.46, .48, .26, .08)
    elif kind == 'chart':                                # 排产 / 分析
        P(.12, .54, .18, .32, None, MSO_SHAPE.ROUNDED_RECTANGLE, .12)
        P(.40, .34, .18, .52, None, MSO_SHAPE.ROUNDED_RECTANGLE, .12)
        P(.68, .16, .18, .70, None, MSO_SHAPE.ROUNDED_RECTANGLE, .12)
    elif kind == 'screen':                               # 看板 / ESOP
        P(.06, .10, .88, .62, None, MSO_SHAPE.ROUNDED_RECTANGLE, .10)
        P(.42, .74, .16, .10); P(.26, .84, .48, .06)
    elif kind == 'box':                                  # 物料 / 料箱
        P(.16, .16, .68, .70, None, MSO_SHAPE.CAN)
    elif kind == 'shelf':                                # 三级库存
        for yy in (.14, .44, .74):
            P(.08, yy, .84, .12, None, MSO_SHAPE.ROUNDED_RECTANGLE, .30)
        P(.10, .14, .08, .72); P(.82, .14, .08, .72)
    elif kind == 'truck':                                # 配送 / AGV
        P(.04, .30, .52, .34, None, MSO_SHAPE.ROUNDED_RECTANGLE, .12)
        P(.58, .42, .30, .22, None, MSO_SHAPE.ROUNDED_RECTANGLE, .18)
        O(.14, .66, .18, .18); O(.60, .66, .18, .18)
    elif kind == 'warn':                                 # 异常 / 预警
        P(.06, .12, .88, .74, None, MSO_SHAPE.ISOSCELES_TRIANGLE)
        P(.45, .36, .10, .26, bg); O(.45, .66, .10, .10, bg)
    elif kind == 'tower':                                # Andon 呼叫
        O(.26, .06, .48, .30); P(.42, .34, .16, .30)
        P(.28, .64, .44, .20, None, MSO_SHAPE.ROUNDED_RECTANGLE, .30)
    elif kind == 'network':                              # 集成总线 / 设备联网
        O(.04, .06, .26, .26); O(.70, .06, .26, .26); O(.37, .68, .26, .26)
        P(.14, .28, .06, .42); P(.80, .28, .06, .42); P(.30, .48, .40, .06)
    elif kind == 'cycle':                                # 循环拉动 / 双向追溯
        P(.06, .06, .88, .88, None, MSO_SHAPE.BLOCK_ARC, .5)
        P(.06, .06, .88, .88, None, MSO_SHAPE.BLOCK_ARC, .5, 180)
    elif kind == 'wrench':                               # 处置恢复
        P(.12, .10, .18, .66, None, MSO_SHAPE.RECTANGLE, None, 45)
        O(.44, .30, .44, .44); O(.55, .41, .22, .22, bg)
    elif kind == 'route':                                # 工艺路线 / 过站
        O(.04, .38, .20, .20); O(.40, .06, .20, .20); O(.76, .60, .20, .20)
        P(.18, .44, .28, .07); P(.54, .22, .28, .07); P(.54, .22, .07, .44)
    elif kind == 'gauge':                                # SPC / 监控
        P(.08, .18, .84, .62, None, MSO_SHAPE.BLOCK_ARC, .5, 180)
        P(.44, .30, .10, .32, None, MSO_SHAPE.RECTANGLE, None, 30)
        O(.42, .76, .16, .16)
    elif kind == 'book':                                 # 知识沉淀 / 复盘
        P(.12, .10, .76, .80, None, MSO_SHAPE.ROUNDED_RECTANGLE, .08)
        P(.20, .10, .10, .80, bg)
        P(.38, .26, .40, .07, bg); P(.38, .44, .40, .07, bg)
    elif kind == 'factory':                              # 工厂级
        P(.06, .46, .88, .42)
        P(.02, .18, .96, .32, None, MSO_SHAPE.ISOSCELES_TRIANGLE)
        P(.70, .04, .12, .24)
    elif kind == 'conveyor':                             # 产线级
        P(.06, .38, .88, .20, None, MSO_SHAPE.ROUNDED_RECTANGLE, .40)
        for i in (.10, .38, .66):
            P(i, .12, .24, .22, None, MSO_SHAPE.ROUNDED_RECTANGLE, .18)
    elif kind == 'bolt':                                 # 拉动信号
        P(.24, .06, .52, .88, None, MSO_SHAPE.LIGHTNING_BOLT)
    elif kind == 'calc':                                 # 算法 / 要素数据化
        P(.10, .08, .80, .84, None, MSO_SHAPE.ROUNDED_RECTANGLE, .10)
        P(.20, .16, .60, .20, bg)
        for r_ in (.46, .62, .78):
            for c_ in (.22, .44, .66):
                P(c_ - .04, r_ - .04, .10, .10, bg)
    else:                                                # 兜底
        O(.10, .10, .80, .80)


def icon(s, x, y, sz, kind, c, tint):
    """圆角方形浅底徽章 + 矢量图标"""
    b = rect(s, x, y, sz, sz, fill=tint, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.20)
    b.shadow.inherit = False
    glyph(s, x + sz * 0.17, y + sz * 0.17, sz * 0.66, kind, c, tint)
    return b


def num_circle(s, x, y, d, n, color, fs=8):
    o = rect(s, x, y, d, d, fill=color, shape=MSO_SHAPE.OVAL)
    o.shadow.inherit = False
    _, f = tb(s, x, y + 0.01, d, d, anchor=MSO_ANCHOR.MIDDLE)
    run(f, str(n), size=fs, color=WHITE, bold=True, first=True, align=PP_ALIGN.CENTER)


def chip(s, x, y, w, h, fill, line_c, adj=0.10):
    c = rect(s, x, y, w, h, fill=fill, line=line_c, lw=0.75,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=adj)
    c.shadow.inherit = False
    return c


def kpi_grid(s, x, y, w, kpis, cols, rows, cgap=0.12, rgap=0.12,
             vfs=16, lfs=9, chip_h=0.80):
    cwid = (w - (cols - 1) * cgap) / cols
    for i, (v, lb) in enumerate(kpis):
        r, c = divmod(i, cols)
        px = x + c * (cwid + cgap)
        py = y + r * (chip_h + rgap)
        chip(s, px, py, cwid, chip_h, RGBColor(0xFF, 0xF8, 0xEC), RGBColor(0xF2, 0xDC, 0xB4), 0.12)
        _, f = tb(s, px, py + chip_h * 0.14, cwid, chip_h * 0.48, anchor=MSO_ANCHOR.MIDDLE)
        run(f, v, size=vfs, color=C_VALUE, bold=True, first=True, align=PP_ALIGN.CENTER)
        _, f = tb(s, px, py + chip_h * 0.60, cwid, chip_h * 0.34, anchor=MSO_ANCHOR.MIDDLE)
        run(f, lb, size=lfs, color=RGBColor(0x8A, 0x6D, 0x3B), first=True, align=PP_ALIGN.CENTER)


def flow_h(s, x, y, w, steps, node_w=None, node_h=0.52, gap=0.10,
           ncolor=C_SOLVE, nfs=9, dfs=8, cir=0.185):
    """横向流程节点链（单行）"""
    n = len(steps)
    if node_w is None:
        node_w = (w - (n - 1) * gap) / n
    for i, (nm, ds) in enumerate(steps):
        px = x + i * (node_w + gap)
        chip(s, px, y, node_w, node_h, WHITE, LINE, 0.12)
        rect(s, px, y + 0.09, 0.045, node_h - 0.18, fill=ncolor)
        _, f = tb(s, px + 0.16, y + 0.06, node_w - 0.28, 0.22)
        run(f, nm, size=nfs, color=INK, bold=True, first=True)
        _, f = tb(s, px + 0.16, y + 0.28, node_w - 0.28, 0.22)
        run(f, ds, size=dfs, color=GREY, first=True)
    return node_w


def flow_grid(s, x, y, w, steps, cols, cell_h=0.50, cgap=0.12, rgap=0.12,
              ncolor=C_SOLVE, dfs=8.5):
    """网格化流程（多行）"""
    n = len(steps)
    rows = (n + cols - 1) // cols
    cwid = (w - (cols - 1) * cgap) / cols
    for i, (nm, ds) in enumerate(steps):
        r, c = divmod(i, cols)
        px = x + c * (cwid + cgap)
        py = y + r * (cell_h + rgap)
        chip(s, px, py, cwid, cell_h, WHITE, LINE, 0.12)
        num_circle(s, px + 0.13, py + (cell_h - 0.185) / 2, 0.185, i + 1, ncolor, 7.5)
        _, f = tb(s, px + 0.37, py + 0.04, cwid - 0.50, cell_h - 0.08,
                  anchor=MSO_ANCHOR.MIDDLE)
        rich(f, [(nm, INK, True, 9), ('  ' + ds, GREY, False, dfs)], first=True, line=1.0)
    return rows * (cell_h + rgap) - rgap


def flow_v(s, x, y, w, steps, color=C_SOLVE, step=0.40, dfs=8.5):
    """竖向时间轴流程（左竖线 + 圆节点）"""
    n = len(steps)
    for i, (nm, ds) in enumerate(steps):
        yy = y + i * step
        cy = yy + 0.035
        if i < n - 1:
            rect(s, x + 0.088, cy + 0.185, 0.012, step - 0.185 + 0.02,
                 fill=RGBColor(0xCF, 0xD8, 0xE2))
        num_circle(s, x, cy, 0.185, i + 1, color)
        _, f = tb(s, x + 0.28, yy - 0.015, w - 0.28, step)
        rich(f, [(nm, INK, True, 9), ('  ·  ', RGBColor(0xC3, 0xCB, 0xD4), False, 8.5),
                 (ds, GREY, False, dfs)], first=True, line=1.0)
    return n * step


def _line_centers(tf):
    """每个段落首行竖直中心（相对 textbox 顶，英寸）。

    与 check_layout 同口径：单行高度 = fs * 1.30 * ls（ls 为 float 行距倍数）。
    用于把彩色圆点(业务需求绿点 / 当前痛点红点)精确对齐到对应文字首行中心，
    杜绝逐行递增的漂移。
    """
    centers, pos = [], 0.0
    for p in tf.paragraphs:
        if not p.runs:
            continue
        fs = max((r.font.size.pt if r.font.size else 12) for r in p.runs)
        ls = p.line_spacing
        ls = ls if isinstance(ls, float) else 1.0
        lh = fs * 1.30 * ls / 72.0
        centers.append(pos + lh / 2.0)
        sa = (p.space_after.pt if p.space_after is not None else 0.0)
        pos += lh + sa / 72.0
    return centers


def bullets(s, x, y, w, items, color, size=8.5, step=None, gap=0.18,
            dot_size=0.085, first=False):
    """彩色圆点项目符号列表（业务需求绿点）。

    圆点按各段真实行高对齐到文字首行中心；文本框高度按内容自适应，
    避免溢出（满足"文字不超边界 + 标点精确对齐"）。
    """
    box, f = tb(s, x + gap, y, w - gap, 0.4)
    for i, it in enumerate(items):
        run(f, it, size=size, color=TXT, first=(i == 0),
            space_after=4.0, line=1.06)
    centers = _line_centers(f)
    if centers:
        last_lh = (size * 1.30 * 1.06 / 72.0) / 2.0
        box.height = Inches(centers[-1] + last_lh + 0.14)
    for c in centers:
        dot(s, x + 0.02, y + c - dot_size / 2.0, color, dot_size)
    return f


def pain_block(s, x, y, w, items, color, step=0.55, title_sz=9, desc_sz=8,
               gap=0.18, dot_size=0.085, tb_h=0.52):
    """当前痛点红点列表：每条 = 粗体标题 + 灰色说明，圆点对齐标题首行中心。

    tb_h 为每条文本框高度（默认 0.52，与旧版一致：文字居顶、框可略大于步距，
    不影响版式体检；体检按文字实测高度判定，故不溢出）。
    """
    for i, (t, ds) in enumerate(items):
        yy = y + i * step
        box, f = tb(s, x + gap, yy, w - gap, tb_h)
        run(f, t, size=title_sz, color=INK, bold=True, first=True,
            space_after=1.5, line=1.0)
        run(f, ds, size=desc_sz, color=GREY, line=1.05)
        c0 = _line_centers(f)[0]
        dot(s, x + 0.02, yy + c0 - dot_size / 2.0, color, dot_size)
    return


def panel(s, x, y, w, h, color):
    p = rect(s, x, y, w, h, fill=PANEL, line=LINE, lw=0.75,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
    shadow(p, blur=6, dist=1, alpha=8000)
    rect(s, x, y, 0.05, h, fill=color)
    rect(s, x + 0.05, y, w - 0.05, 0.045, fill=color)
    return p


def flow_grid_icon(s, x, y, w, steps, icons, cols, cell_h=0.52, cgap=0.12, rgap=0.10,
                   ncolor=C_SOLVE, dfs=8.5, tint=RGBColor(0xF2, 0xF8, 0xF9)):
    """网格化流程：图标徽章 + 单行「名称 描述」"""
    n = len(steps)
    cwid = (w - (cols - 1) * cgap) / cols
    for i, ((nm, ds), ik) in enumerate(zip(steps, icons)):
        r, c = divmod(i, cols)
        px = x + c * (cwid + cgap)
        py = y + r * (cell_h + rgap)
        chip(s, px, py, cwid, cell_h, WHITE, LINE, 0.12)
        isz = cell_h * 0.74
        icon(s, px + 0.10, py + (cell_h - isz) / 2, isz, ik, ncolor, tint)
        _, f = tb(s, px + 0.10 + isz + 0.08, py + 0.02,
                  cwid - 0.20 - isz - 0.08, cell_h - 0.04, anchor=MSO_ANCHOR.MIDDLE)
        rich(f, [(nm, INK, True, 9), ('  ' + ds, GREY, False, dfs)], first=True, line=1.0)
    rows = (n + cols - 1) // cols
    return rows * (cell_h + rgap) - rgap


def flow_h_icon(s, x, y, w, steps, icons, node_h=0.68, gap=0.10,
                ncolor=C_SOLVE, tint=RGBColor(0xF2, 0xF8, 0xF9)):
    """横向泳道：图标居中在上，名称 + 描述居中在下"""
    n = len(steps)
    nw = (w - (n - 1) * gap) / n
    for i, ((nm, ds), ik) in enumerate(zip(steps, icons)):
        px = x + i * (nw + gap)
        chip(s, px, y, nw, node_h, WHITE, LINE, 0.12)
        isz = node_h * 0.42
        icon(s, px + (nw - isz) / 2, y + 0.07, isz, ik, ncolor, tint)
        _, f = tb(s, px + 0.06, y + isz + 0.10, nw - 0.12, 0.22)
        run(f, nm, size=8.5, color=INK, bold=True, first=True, align=PP_ALIGN.CENTER)
        _, f = tb(s, px + 0.06, y + isz + 0.30, nw - 0.12, 0.22)
        run(f, ds, size=7.5, color=GREY, first=True, align=PP_ALIGN.CENTER)
    return node_h


def flow_v_icon(s, x, y, w, steps, icons, color=C_SOLVE, step=0.38, dfs=8.5,
                tint=RGBColor(0xF2, 0xF8, 0xF9)):
    """竖向时间轴：图标徽章替代序号圆"""
    n = len(steps)
    for i, ((nm, ds), ik) in enumerate(zip(steps, icons)):
        yy = y + i * step
        cy = yy + 0.03
        if i < n - 1:
            rect(s, x + 0.13, cy + 0.26, 0.012, step - 0.26 + 0.02,
                 fill=RGBColor(0xCF, 0xD8, 0xE2))
        icon(s, x, cy, 0.28, ik, color, tint)
        _, f = tb(s, x + 0.36, yy - 0.015, w - 0.36, step)
        rich(f, [(nm, INK, True, 9), ('  ·  ', RGBColor(0xC3, 0xCB, 0xD4), False, 8.5),
                 (ds, GREY, False, dfs)], first=True, line=1.0)
    return n * step


def icon_cards(s, x, y, w, items, cols, color, tint, cell_h=0.44, cgap=0.12, rgap=0.08,
               lfs=8.5, sfs=7.5):
    """图标卡网格：items = [(icon, 主标签, 副标签), ...]"""
    rows = (len(items) + cols - 1) // cols
    cwid = (w - (cols - 1) * cgap) / cols
    for i, (ik, lb, sub) in enumerate(items):
        r, c = divmod(i, cols)
        px = x + c * (cwid + cgap)
        py = y + r * (cell_h + rgap)
        chip(s, px, py, cwid, cell_h, WHITE, LINE, 0.12)
        isz = cell_h * 0.72
        icon(s, px + 0.09, py + (cell_h - isz) / 2, isz, ik, color, tint)
        tx = px + 0.09 + isz + 0.08
        _, f = tb(s, tx, py + 0.05, cwid - 0.17 - isz - 0.08, cell_h - 0.08,
                  anchor=MSO_ANCHOR.MIDDLE)
        run(f, lb, size=lfs, color=TXT, bold=True, first=True, line=1.0)
        if sub:
            run(f, sub, size=sfs, color=GREY, line=1.05)
    return rows * (cell_h + rgap) - rgap


def panel_head(s, x, y, w, name, sub, color):
    _, f = tb(s, x + 0.16, y + 0.09, w - 0.32, 0.24)
    run(f, name, size=10.5, color=color, bold=True, first=True)
    _, f = tb(s, x + 0.16, y + 0.325, w - 0.32, 0.20)
    run(f, sub, size=8, color=GREY2, first=True)


def note_bar(s, x, y, w, text, color):
    chip(s, x, y, w, 0.30, RGBColor(0xF2, 0xF8, 0xF9), RGBColor(0xCF, 0xE2, 0xE5), 0.30)
    _, f = tb(s, x + 0.14, y + 0.02, w - 0.28, 0.26, anchor=MSO_ANCHOR.MIDDLE)
    run(f, text, size=8, color=color, bold=True, first=True)


# ---------------- 页眉 / 页脚（导航位置统一） ----------------
# 每套方案仅需在 build_deck.py 里覆盖这几项，引擎本身不改
META = {
    'total': 5,                                                        # 总页数（页眉 "01 / 05"）
    'corner': '客户名 · MES 建设方案',                                   # 页眉右上角项目名
    # V2.0 默认不绘制页脚标注（数据来源 / 品牌）；需恢复旧行为时置 footer_anno=True 并填 source/brand
    'source': None,
    'brand': None,
    'footer_anno': False,
}


def header(s, no, title, sub, band_h=1.05):
    rect(s, 0, 0, SW, SH, fill=BG)
    seg = SW / 4
    for k, col in enumerate([C_NEED, C_PAIN, C_SOLVE, C_VALUE]):
        rect(s, k * seg, 0, seg, 0.10, fill=col)
    rect(s, 0, 0.10, SW, band_h, fill=NAVY)
    rect(s, 0.50, 0.36, 0.075, 0.52, fill=RGBColor(0x4D, 0xA3, 0xFF))
    _, f = tb(s, 0.72, 0.30, 9.6, 0.44)
    run(f, title, size=21, color=WHITE, bold=True, first=True)
    _, f = tb(s, 0.72, 0.76, 9.6, 0.30)
    run(f, sub, size=10, color=SUBTLE, first=True)
    _, f = tb(s, 10.9, 0.28, 1.93, 0.46)
    rich(f, [(no, WHITE, True, 24), (' / %02d' % META['total'], RGBColor(0x6F, 0x8E, 0xAE), False, 11)],
         first=True, align=PP_ALIGN.RIGHT)
    _, f = tb(s, 10.4, 0.78, 2.43, 0.26)
    run(f, META['corner'], size=8.5, color=SUBTLE, first=True, align=PP_ALIGN.RIGHT)


def footer(s):
    """V2.0 默认不绘制页脚标注（数据来源 / 品牌）。
    仅当 META['footer_anno']=True 且填了 source/brand 时恢复，保留原能力。"""
    if not K_META().get('footer_anno'):
        return
    rect(s, 0.50, 7.06, SW - 1.0, 0.008, fill=LINE)
    if K_META().get('source'):
        _, f = tb(s, 0.50, 7.14, 10.5, 0.24)
        run(f, K_META()['source'], size=7.5, color=GREY2, first=True)
    if K_META().get('brand'):
        _, f = tb(s, 11.0, 7.14, 1.83, 0.24)
        run(f, K_META()['brand'], size=7.5, color=GREY2, first=True, align=PP_ALIGN.RIGHT)


def K_META():
    return META
# ================= 页面版式 =================
def page1(s, d):
    """骨架 A：右重双栏 5:7 + 底部通栏价值带（总览页，横向流程链为主视觉）"""
    lw, rw = cw(5), cw(7)
    rx = cx(5)
    hs = 4.05
    yv = TOP + hs + GUT
    hv = BOT - yv
    # 需求 / 痛点（左，竖叠；痛点为紧凑单行，避免与底部通栏争高）
    hn1 = 1.80
    card(s, MARGIN, TOP, lw, hn1, C_NEED, '业务需求', 'BUSINESS REQUIREMENT')
    bullets(s, MARGIN + 0.42, TOP + 0.74, lw - 0.46, d['need'], C_NEED,
            size=9, gap=0.18)
    yp = TOP + hn1 + GUT
    hp = (yv - GUT) - yp
    card(s, MARGIN, yp, lw, hp, C_PAIN, '当前痛点', 'CURRENT PAIN POINTS')
    bullets(s, MARGIN + 0.42, yp + 0.74, lw - 0.46,
            ['%s · %s' % (t, ds) for t, ds in d['pain']], C_PAIN, size=9, gap=0.18)
    # 解决方案（右，大卡）
    card(s, rx, TOP, rw, hs, C_SOLVE, '解决方案', 'SOLUTION · 双思维视角')
    inner_x, inner_w = rx + 0.22, rw - 0.44
    panel_head(s, inner_x, TOP + 0.72, inner_w,
               d['flow']['name'], d['flow']['sub'], C_SOLVE)
    nw = (inner_w - 2 * 0.12) / 3
    flow_grid_icon(s, inner_x, TOP + 1.06, inner_w, d['flow']['steps'], d['flow_icons'], 3,
                   cell_h=0.52, cgap=0.12, rgap=0.12, ncolor=C_SOLVE, dfs=8.5)
    y2 = TOP + 1.06 + 2 * (0.52 + 0.12) + 0.02
    panel_head(s, inner_x, y2, inner_w, d['alt']['name'], d['alt']['sub'], d['alt']['color'])
    icon_cards(s, inner_x, y2 + 0.54, inner_w, d['alt']['cards'], 2,
               d['alt']['color'], RGBColor(0xF4, 0xF4, 0xFD),
               cell_h=0.44, cgap=0.14, rgap=0.09, lfs=8.5, sfs=7.5)
    # 价值（底部通栏：左标题 + 中主张 + 右四 KPI）
    card(s, MARGIN, yv, CONT_W, hv, C_VALUE, '方案价值', 'EXPECTED VALUE', head=0)
    _, f = tb(s, MARGIN + 2.60, yv + 0.20, 3.80, hv - 0.34, anchor=MSO_ANCHOR.MIDDLE)
    run(f, d['claim'], size=12, color=RGBColor(0x8A, 0x6D, 0x3B), bold=True,
        first=True, line=1.25)
    rect(s, MARGIN + 6.62, yv + 0.20, 0.012, hv - 0.40, fill=RGBColor(0xF0, 0xDC, 0xB8))
    kpi_grid(s, MARGIN + 6.90, yv + 0.26, CONT_W - 6.90 - 0.22, d['kpis'], 4, 1,
             cgap=0.13, chip_h=0.86, vfs=17, lfs=9)


def page2(s, d):
    """骨架 B：顶部通栏方案泳道 + 底部三栏 5:4:3（流程主导页，第一落点在顶部横带）"""
    hs = 3.10
    card(s, MARGIN, TOP, CONT_W, hs, C_SOLVE, '解决方案', 'SOLUTION · 双思维视角')
    ly = TOP + 0.72
    lw8 = cw(8)
    panel_head(s, MARGIN + 0.22, ly, lw8, d['flow']['name'], d['flow']['sub'], C_SOLVE)
    flow_h_icon(s, MARGIN + 0.22, ly + 0.56, lw8, d['flow']['steps'], d['flow_icons'],
                node_h=0.68, gap=0.10, ncolor=C_SOLVE)
    note_bar(s, MARGIN + 0.22, ly + 1.52, lw8, d['flow']['note'], C_SOLVE)
    # 右：另一思维（图标卡 2×3）
    rx4 = cx(8)
    w4 = cw(4)
    rect(s, rx4 - GUT / 2, ly - 0.02, 0.012, hs - 0.74, fill=SEPL)
    panel_head(s, rx4, ly, w4, d['alt']['name'], d['alt']['sub'], d['alt']['color'])
    icon_cards(s, rx4, ly + 0.58, w4 - 0.16, d['alt']['cards'], 2,
               d['alt']['color'], RGBColor(0xF1, 0xF7, 0xF4),
               cell_h=0.42, cgap=0.10, rgap=0.08, lfs=8, sfs=7)
    # 底部三栏
    yb = TOP + hs + GUT
    hb = BOT - yb
    card(s, cx(0), yb, cw(5), hb, C_NEED, '业务需求', 'BUSINESS REQUIREMENT')
    bullets(s, cx(0) + 0.42, yb + 0.80, cw(5) - 0.46, d['need'], C_NEED, size=9.5, gap=0.18)
    card(s, cx(5), yb, cw(4), hb, C_PAIN, '当前痛点', 'CURRENT PAIN POINTS')
    pain_block(s, cx(5) + 0.22, yb + 0.78, cw(4) - 0.44, d['pain'], C_PAIN,
               step=0.35, title_sz=9, desc_sz=8)
    card(s, cx(9), yb, cw(3), hb, C_VALUE, '方案价值', 'EXPECTED VALUE')
    kpi_grid(s, cx(9) + 0.22, yb + 0.76, cw(3) - 0.44, d['kpis'], 2, 2,
             cgap=0.12, rgap=0.12, chip_h=0.70, vfs=15, lfs=8.5)


def page3(s, d):
    """骨架 C：顶部需求条 + 中部 4:8 + 底部 8:4（矩阵页，网格化块状节奏）"""
    hn = 0.92
    card(s, MARGIN, TOP, CONT_W, hn, C_NEED, '业务需求', 'BUSINESS REQUIREMENT', head=0)
    nw = (CONT_W - 1.95 - 3 * 0.13) / 4
    for i, t in enumerate(d['need']):
        px = MARGIN + 1.95 + i * (nw + 0.13)
        chip(s, px, TOP + 0.15, nw, 0.62, RGBColor(0xF3, 0xF7, 0xFD),
             RGBColor(0xCF, 0xDF, 0xFA), 0.14)
        dot(s, px + 0.13, TOP + 0.40, C_NEED, 0.065)
        _, f = tb(s, px + 0.28, TOP + 0.15, nw - 0.42, 0.62, anchor=MSO_ANCHOR.MIDDLE)
        run(f, t, size=9, color=TXT, first=True, line=1.05)
    # 中部
    ym = TOP + hn + GUT
    hm = 3.08
    card(s, cx(0), ym, cw(4), hm, C_PAIN, '当前痛点', 'CURRENT PAIN POINTS')
    pain_block(s, cx(0) + 0.22, ym + 0.80, cw(4) - 0.44, d['pain'], C_PAIN,
               step=0.56, title_sz=9, desc_sz=8)
    card(s, cx(4), ym, cw(8), hm, C_SOLVE, '解决方案', 'SOLUTION · 双思维视角')
    ix, iw = cx(4) + 0.22, cw(8) - 0.44
    panel_head(s, ix, ym + 0.72, iw, d['flow']['name'], d['flow']['sub'], C_SOLVE)
    flow_grid_icon(s, ix, ym + 1.06, iw, d['flow']['steps'], d['flow_icons'], 3,
                   cell_h=0.40, cgap=0.12, rgap=0.08, ncolor=C_SOLVE, dfs=8)
    y2 = ym + 1.06 + 2 * (0.40 + 0.08) + 0.06
    panel_head(s, ix, y2, iw, d['alt']['name'], d['alt']['sub'], d['alt']['color'])
    ac = d['alt']['color']
    cw6 = (iw - 5 * 0.08) / 6
    for i, c in enumerate(d['alt']['cards'][:6]):
        ik, lb = c[0], c[1]
        px = ix + i * (cw6 + 0.08)
        py = y2 + 0.52
        chip(s, px, py, cw6, 0.46, WHITE, LINE, 0.14)
        icon(s, px + (cw6 - 0.24) / 2, py + 0.05, 0.24, ik, ac, RGBColor(0xFB, 0xF2, 0xE9))
        _, f = tb(s, px + 0.08, py + 0.29, cw6 - 0.16, 0.16, anchor=MSO_ANCHOR.MIDDLE)
        run(f, lb, size=8, color=TXT, first=True, align=PP_ALIGN.CENTER, line=1.0)
    # 底部：价值 8 栏 + 卡控点提示块 4 栏
    yb = ym + hm + GUT
    hb = BOT - yb
    card(s, cx(0), yb, cw(8), hb, C_VALUE, '方案价值', 'EXPECTED VALUE')
    kpi_grid(s, cx(0) + 0.22, yb + 0.72, cw(8) - 0.44, d['kpis'], 4, 1,
             cgap=0.14, chip_h=0.54, vfs=15, lfs=8.5)
    bl = chip(s, cx(8), yb, cw(4), hb, RGBColor(0xF2, 0xF8, 0xF9),
              RGBColor(0xCF, 0xE2, 0xE5), 0.08)
    bl.shadow.inherit = False
    rect(s, cx(8), yb + 0.16, 0.05, hb - 0.32, fill=C_SOLVE)
    _, f = tb(s, cx(8) + 0.22, yb + 0.14, cw(4) - 0.40, hb - 0.24, anchor=MSO_ANCHOR.MIDDLE)
    run(f, '卡控点即质量门', size=10, color=C_SOLVE, bold=True, first=True, space_after=3)
    run(f, d['flow']['note'].replace('关键卡控点：', '').replace(' · ', ' / '),
        size=8, color=GREY, line=1.15)


def page4(s, d):
    """骨架 D：双塔竖切 左三段 : 右通高 = 6:6（对称双塔，右侧流程塔为绝对主角）"""
    lw6 = cw(6)
    rl = cx(6)
    wr = cw(6)
    # 左塔：需求
    h1 = 1.80
    card(s, MARGIN, TOP, lw6, h1, C_NEED, '业务需求', 'BUSINESS REQUIREMENT')
    bullets(s, MARGIN + 0.42, TOP + 0.74, lw6 - 0.46, d['need'], C_NEED, size=9, gap=0.18)
    # 左塔：痛点（2x2 网格）
    y2 = TOP + h1 + GUT
    h2 = 1.80
    card(s, MARGIN, y2, lw6, h2, C_PAIN, '当前痛点', 'CURRENT PAIN POINTS')
    cellw = (lw6 - 0.44 - 0.14) / 2
    cellh = (h2 - 0.78 - 0.10) / 2
    for i, (t, ds) in enumerate(d['pain']):
        r, c = divmod(i, 2)
        px = MARGIN + 0.22 + c * (cellw + 0.14)
        py = y2 + 0.78 + r * (cellh + 0.10)
        chip(s, px, py, cellw, cellh, RGBColor(0xFD, 0xF5, 0xF5),
             RGBColor(0xF3, 0xD6, 0xD5), 0.10)
        _, f = tb(s, px + 0.30, py + 0.06, cellw - 0.44, cellh - 0.10)
        run(f, t, size=9, color=INK, bold=True, first=True, space_after=1.5, line=1.0)
        run(f, ds, size=8, color=GREY, line=1.05)
        c0 = _line_centers(f)[0]
        dot(s, px + 0.10, py + c0 - 0.085 / 2.0, C_PAIN, 0.085)
    # 左塔：价值
    y3 = y2 + h2 + GUT
    h3 = BOT - y3
    card(s, MARGIN, y3, lw6, h3, C_VALUE, '方案价值', 'EXPECTED VALUE')
    kpi_grid(s, MARGIN + 0.22, y3 + 0.74, lw6 - 0.44, d['kpis'], 4, 1,
             cgap=0.12, chip_h=0.80, vfs=16, lfs=8.5)
    # 右塔：解决方案通高（竖向时间轴 + 要点）
    card(s, rl, TOP, wr, BODY_H, C_SOLVE, '解决方案', 'SOLUTION · 双思维视角')
    ix, iw = rl + 0.22, wr - 0.44
    panel_head(s, ix, TOP + 0.72, iw, d['flow']['name'], d['flow']['sub'], C_SOLVE)
    used = flow_v_icon(s, ix + 0.16, TOP + 1.06, iw - 0.16, d['flow']['steps'],
                       d['flow_icons'], C_SOLVE, step=0.34, dfs=8.5)
    note_bar(s, ix, TOP + 1.06 + used + 0.04, iw, d['flow']['note'], C_SOLVE)
    y2r = TOP + 1.06 + used + 0.40
    panel_head(s, ix, y2r, iw, d['alt']['name'], d['alt']['sub'], d['alt']['color'])
    icon_cards(s, ix, y2r + 0.54, iw, d['alt']['cards'], 2,
               d['alt']['color'], RGBColor(0xF1, 0xF6, 0xFD),
               cell_h=0.40, cgap=0.12, rgap=0.07, lfs=8, sfs=7)


def page5(s, d):
    """骨架 E：三竖栏聚焦 3:6:3 + 底部反色价值带（剧场式，方案居中为唯一焦点）"""
    hcol = 4.55
    # 左：需求（窄竖卡）
    card(s, cx(0), TOP, cw(3), hcol, C_NEED, '业务需求', 'BUSINESS REQUIREMENT')
    bullets(s, cx(0) + 0.40, TOP + 0.80, cw(3) - 0.42, d['need'], C_NEED,
            size=8.5, gap=0.18)
    # 中：解决方案（唯一焦点）
    card(s, cx(3), TOP, cw(6), hcol, C_SOLVE, '解决方案', 'SOLUTION · 双思维视角')
    ix, iw = cx(3) + 0.22, cw(6) - 0.44
    panel_head(s, ix, TOP + 0.72, iw, d['flow']['name'], d['flow']['sub'], C_SOLVE)
    used = flow_grid_icon(s, ix, TOP + 1.06, iw, d['flow']['steps'], d['flow_icons'], 2,
                          cell_h=0.42, cgap=0.12, rgap=0.08, ncolor=C_SOLVE, dfs=8)
    note_bar(s, ix, TOP + 1.06 + used + 0.04, iw, d['flow']['note'], C_SOLVE)
    y2 = TOP + 1.06 + used + 0.38
    panel_head(s, ix, y2, iw, d['alt']['name'], d['alt']['sub'], d['alt']['color'])
    icon_cards(s, ix, y2 + 0.50, iw, d['alt']['cards'], 2,
               d['alt']['color'], RGBColor(0xF5, 0xF1, 0xFA),
               cell_h=0.42, cgap=0.12, rgap=0.08, lfs=8, sfs=7)
    # 右：痛点（窄竖卡）
    card(s, cx(9), TOP, cw(3), hcol, C_PAIN, '当前痛点', 'CURRENT PAIN POINTS')
    pain_block(s, cx(9) + 0.20, TOP + 0.80, cw(3) - 0.40, d['pain'], C_PAIN,
               step=0.62, title_sz=9, desc_sz=8)
    # 底部反色价值带
    yv = TOP + hcol + GUT
    hv = BOT - yv
    card(s, MARGIN, yv, CONT_W, hv, C_VALUE, '方案价值', 'EXPECTED VALUE',
         head=0, dark=True)
    _, f = tb(s, MARGIN + 0.30, yv + 0.10, 2.10, hv - 0.20, anchor=MSO_ANCHOR.MIDDLE)
    run(f, '价值度量', size=13, color=WHITE, bold=True, first=True, line=1.0)
    run(f, 'QCDS 四维改善', size=9.5, color=SUBTLE, line=1.0)
    kx = MARGIN + 2.55
    kw = (CONT_W - 2.55 - 0.30 - 3 * 0.16) / 4
    for i, (v, lb) in enumerate(d['kpis']):
        px = kx + i * (kw + 0.16)
        _, f = tb(s, px, yv + 0.14, kw, 0.38, anchor=MSO_ANCHOR.MIDDLE)
        run(f, v, size=17, color=C_VALUE, bold=True, first=True, align=PP_ALIGN.CENTER)
        _, f = tb(s, px, yv + 0.56, kw, 0.26, anchor=MSO_ANCHOR.MIDDLE)
        run(f, lb, size=9.5, color=RGBColor(0xA9, 0xBD, 0xD4), first=True, align=PP_ALIGN.CENTER)
        if i < 3:
            rect(s, px + kw + 0.07, yv + 0.24, 0.012, hv - 0.48, fill=RGBColor(0x2A, 0x44, 0x62))


# ================= 骨架注册表 =================
BONES = {
    'A': (page1, '右重双栏 5:7 + 底部通栏价值带', '总览/体系类：左侧需求+痛点，右侧主视觉流程链，底部价值横带'),
    'B': (page2, '上通栏 + 下三栏 5:4:3', '计划/排产类：顶部横向端到端流程链，下部三栏拆解执行要素'),
    'C': (page3, '上条 + 中 4:8 + 底 8:4', '工艺/质量类：中区主次分明，底部双区做闭环与防错'),
    'D': (page4, '双塔竖切 左三段 : 右通高 6:6', '物料/追溯类：左右对望，弱化横向流程、强调并行对比'),
    'E': (page5, '三竖栏 3:6:3 + 底反色带', '综合/价值类：中间主叙事，两侧辅助，底部反色带收口'),
}
