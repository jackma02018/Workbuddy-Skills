# -*- coding: utf-8 -*-
"""为方案 PPT 的每页注入「系统界面截图」悬浮弹层 + 点击逐张切换动画。

用法：
    python add_screens_overlay.py --src <原.pptx> --out <截图版.pptx> \
        --shots <shots.json> [--lib <截图目录>] [--tail "对应本页「解决方案」区"]

shots.json 结构（页码从 1 开始，与 PPT 页序一致）：
    {
      "1": [["L055_5342ce7e.png", "N2.MES 生产监控与预警平台"],
            ["L105_8f718c68.png", "N2.MOC 制造运营中心"],
            ["L105_eb7b14db.png", "运营驾驶舱"]],
      "2": [...]
    }
  每页 2~5 张均可，脚本按张数自动生成「点击 N 次走完循环」的动画分组。

设计要点（详见 references/screen-overlay.md）：
- 正文零改动：截图以不透明白底弹层叠加在方案区之上，点击出现、讲完收回。
- 完整可见：每张按原始宽高比等比 contain 居中，不裁切、不拉伸。
- 互不遮挡：同页所有截图共用同一几何位置，同一时刻只允许 1 张可见。
- 节奏恒定：每页固定「出现 → 换图…（n-1 次）→ 收回」共 n+1 次点击。
"""
import argparse
import io
import json
import os
import shutil

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.util import Inches, Pt

# ---------------- 几何默认值（13.333 × 7.5 in 画布，方案区主内容带） ----------------
BOX_X, BOX_Y, BOX_W, BOX_H = 1.5667, 1.26, 10.20, 5.74
PAD = 0.06
CAP_H = 0.36
CAP_X = BOX_X + 0.10
CAP_W = BOX_W - 0.20
# 说明条放在弹层顶部内侧：y = BOX_Y + 0.04，落在方案区正文区（TOP=1.36）之上。
# 这样 check_layout.py 会把说明条判给弹层自身（合法归属），而不是判给底下的模块卡片。
CAP_Y = BOX_Y + 0.04
IMG_X = BOX_X + PAD
IMG_Y = CAP_Y + CAP_H + 0.06
IMG_W = BOX_W - 2 * PAD
IMG_H = BOX_Y + BOX_H - 0.06 - IMG_Y

ACCENT = RGBColor(0x0B, 0x72, 0x85)
INK = RGBColor(0x33, 0x42, 0x52)
EDGE = RGBColor(0xC8, 0xD2, 0xDE)


# ---------------- 动画 XML ----------------
def _effect(cid, spid, visible, cls, node_type):
    """单个显隐效果。node_type: 组内第一个用 clickEffect，其余用 withEffect。"""
    return (
        '<p:par><p:cTn id="%(a)d" presetID="1" presetClass="%(cls)s" presetSubtype="0"'
        ' fill="hold" nodeType="%(nt)s"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        '<p:childTnLst><p:set><p:cBhvr>'
        '<p:cTn id="%(b)d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="%(spid)d"/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        '</p:cBhvr><p:to><p:strVal val="%(val)s"/></p:to></p:set>'
        '</p:childTnLst></p:cTn></p:par>'
        % {"a": cid, "b": cid + 1, "spid": spid, "val": visible and "visible" or "hidden",
           "cls": cls, "nt": node_type}
    )


def _click_group(cid, effects):
    return (
        '<p:par><p:cTn id="%d" fill="hold">'
        '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="%d" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        '<p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:par>' % (cid, cid + 1, "".join(effects))
    )


def _timing_xml(groups):
    """groups: [[(spid, visible, 'entr'|'exit'), ...], ...] —— 每个内层列表 = 一次点击。

    坑：同一次点击内的多个效果必须是 clickEffect + withEffect，
    若全写成 clickEffect，一次点击只会播一个效果（n 张图会退化成 2n+2 次点击）。
    """
    cid, blocks = 3, []
    for g in groups:
        effs = []
        for i, (spid, vis, cls) in enumerate(g):
            effs.append(_effect(cid, spid, vis, cls,
                                "clickEffect" if i == 0 else "withEffect"))
            cid += 2
        blocks.append(_click_group(cid, effs))
        cid += 2
    return (
        '<p:timing %s><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        '<p:childTnLst>%s</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0">'
        '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0">'
        '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        % (nsdecls("p", "a"), "".join(blocks))
    )


def _fit(im_w, im_h):
    """等比 contain 居中，返回 (w, h, x, y)。"""
    ar = im_w / float(im_h)
    if ar >= IMG_W / float(IMG_H):
        w, h = IMG_W, IMG_W / ar
    else:
        h, w = IMG_H, IMG_H * ar
    return w, h, IMG_X + (IMG_W - w) / 2.0, IMG_Y + (IMG_H - h) / 2.0


def build(src, out, shots, lib, tail, title="系统界面"):
    def _res(fn):
        p = os.path.join(lib, fn)
        if not os.path.exists(p):
            raise SystemExit("截图缺失：%s" % p)
        return p

    for idx, picks in shots.items():
        if not picks:
            raise SystemExit("第 %s 页没有任何截图" % idx)
        for fn, _lb in picks:
            _res(fn)

    shutil.copyfile(src, out)
    prs = Presentation(out)
    report = []

    for idx, slide in enumerate(prs.slides, 1):
        picks = shots.get(str(idx)) or shots.get(idx)
        if not picks:
            report.append("P%02d 未配置截图，跳过" % idx)
            continue
        n = len(picks)

        # 1) 不透明白底弹层面板（同时充当 figure 容器与遮挡层）
        panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       Inches(BOX_X), Inches(BOX_Y),
                                       Inches(BOX_W), Inches(BOX_H))
        panel.fill.solid()
        panel.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        panel.line.color.rgb = EDGE
        panel.line.width = Pt(1.25)
        panel.shadow.inherit = False
        panel.name = "P%02d-overlay-panel" % idx

        # 2) n 张截图（同几何位置叠放）
        pics = []
        for k, (fn, label) in enumerate(picks, 1):
            p = _res(fn)
            im_w, im_h = Image.open(p).size
            w, h, x, y = _fit(im_w, im_h)
            pic = slide.shapes.add_picture(p, Inches(x), Inches(y),
                                          Inches(w), Inches(h))
            pic._element.nvPicPr.cNvPr.set(
                "name", "P%02d-shot%d-%s" % (idx, k, label.replace(" ", "_")[:40]))
            pic._element.nvPicPr.cNvPr.set(
                "descr", "系统截图 %d/%d：%s" % (k, n, label))
            pics.append(pic)

        # 3) n 条说明（同几何位置叠放，放在面板顶部内侧）
        caps = []
        for k, (fn, label) in enumerate(picks, 1):
            tb = slide.shapes.add_textbox(Inches(CAP_X), Inches(CAP_Y),
                                          Inches(CAP_W), Inches(CAP_H))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            par = tf.paragraphs[0]
            par.alignment = PP_ALIGN.CENTER
            r1 = par.add_run()
            r1.text = "%s %d/%d　" % (title, k, n)
            r1.font.size, r1.font.bold = Pt(10.5), True
            r1.font.color.rgb = ACCENT
            r2 = par.add_run()
            r2.text = label + ("　｜　%s" % tail if tail else "")
            r2.font.size = Pt(10.5)
            r2.font.color.rgb = INK
            tb.name = "P%02d-caption%d" % (idx, k)
            caps.append(tb)

        # 4) 点击分组：① 面板+图1+说明1 ② 换图2 … (n+1) 收回
        pid = panel.shape_id
        groups = [[(pid, True, "entr"),
                   (pics[0].shape_id, True, "entr"),
                   (caps[0].shape_id, True, "entr")]]
        for k in range(1, n):
            groups.append([(pics[k - 1].shape_id, False, "exit"),
                           (caps[k - 1].shape_id, False, "exit"),
                           (pics[k].shape_id, True, "entr"),
                           (caps[k].shape_id, True, "entr")])
        groups.append([(pics[-1].shape_id, False, "exit"),
                       (caps[-1].shape_id, False, "exit"),
                       (pid, False, "exit")])

        slide._element.append(parse_xml(_timing_xml(groups)))
        report.append("P%02d 面板id=%d 图=%s 说明=%s 点击组=%d"
                      % (idx, pid, [p.shape_id for p in pics],
                         [c.shape_id for c in caps], len(groups)))

    prs.save(out)
    print("\n".join(report))
    print("\nSAVED %s (%d bytes)" % (out, os.path.getsize(out)))
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="原方案 PPT")
    ap.add_argument("--out", required=True, help="输出的截图版 PPT")
    ap.add_argument("--shots", required=True, help="截图配置 JSON")
    ap.add_argument("--lib", default="", help="截图素材目录（默认取 shots.json 同目录下的 _lib）")
    ap.add_argument("--tail", default="对应本页「解决方案」区",
                    help="说明条尾部文案，置空则不带尾注")
    ap.add_argument("--title", default="系统界面", help="说明条前缀")
    a = ap.parse_args()

    with io.open(a.shots, encoding="utf-8") as f:
        shots = json.load(f)
    lib = a.lib or os.path.join(os.path.dirname(os.path.abspath(a.shots)), "_lib")
    build(a.src, a.out, shots, lib, a.tail, a.title)


if __name__ == "__main__":
    main()
