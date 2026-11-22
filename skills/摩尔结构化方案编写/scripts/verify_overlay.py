# -*- coding: utf-8 -*-
"""校验截图版 PPT：动画引用、点击分组数、几何是否越界。

用法：
    python verify_overlay.py <截图版.pptx>

输出每页一行：面板/截图/说明数量 · 点击组数 · 悬空形状ID · 几何是否 OK。
判定标准：悬空 spid 必须为「无」，几何必须 True，
点击组数必须等于「截图张数 + 1」（出现 1 次 + 每张之间切换 + 收回 1 次）。
"""
import re
import sys
import zipfile

from pptx import Presentation
from pptx.oxml.ns import qn

SW, SH = 13.333, 7.5
SAFE = 0.0  # 允许贴边，不允许越界


def slide_bounds(shapes):
    xs, ys, x2, y2 = [], [], [], []
    for s in shapes:
        try:
            xs.append(s.left / 914400.0)
            ys.append(s.top / 914400.0)
            x2.append((s.left + s.width) / 914400.0)
            y2.append((s.top + s.height) / 914400.0)
        except TypeError:
            continue
    if not xs:
        return True, None
    ok = (min(xs) >= SAFE - 1e-6 and min(ys) >= SAFE - 1e-6
          and max(x2) <= SW + 1e-6 and max(y2) <= SH + 1e-6)
    return ok, (min(xs), min(ys), max(x2), max(y2))


def main(path):
    # 1) XML 可解析性
    bad = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist()
                 if re.match(r"ppt/slides/slide\d+\.xml$", n)]
        for n in names:
            try:
                from lxml import etree
                etree.fromstring(z.read(n))
            except Exception as e:
                bad.append((n, str(e)[:80]))
    print("页数=%d  尺寸=%.3fx%.3f in" % (len(names), SW, SH))
    print("XML 解析失败文件数=%d %s" % (len(bad), bad))

    prs = Presentation(path)
    for idx, slide in enumerate(prs.slides, 1):
        shapes = list(slide.shapes)
        panels = [s for s in shapes if s.name.startswith("P%02d-overlay-panel" % idx)]
        pics = [s for s in shapes if s.name.startswith("P%02d-shot" % idx)]
        caps = [s for s in shapes if s.name.startswith("P%02d-caption" % idx)]

        tim = slide._element.find(qn("p:timing"))
        if tim is not None:
            from lxml import etree
            txt = etree.tostring(tim).decode()
            spids = [int(x) for x in re.findall(r'spid="(\d+)"', txt)]
            nclick = txt.count('nodeType="clickEffect"')
            nentr = txt.count('presetClass="entr"')
            nexit = txt.count('presetClass="exit"')
        else:
            spids, nclick, nentr, nexit = [], 0, 0, 0

        alive = {s.shape_id for s in shapes}
        missing = sorted({x for x in spids if x not in alive})
        geo_ok, _ = slide_bounds(shapes)
        expect = len(pics) + 1
        flag = "" if nclick == expect else "  << 点击组数异常，应为 %d" % expect
        print("P%02d 面板=%d 截图=%d 说明=%d | 点击组=%d entr=%d exit=%d | "
              "悬空spid=%s | 几何OK=%s%s"
              % (idx, len(panels), len(pics), len(caps), nclick, nentr, nexit,
                 missing or "无", geo_ok, flag))


if __name__ == "__main__":
    main(sys.argv[1])
