# -*- coding: utf-8 -*-
"""版式体检（deck_kit 配套）：1) 文本是否溢出自身文本框  2) 内容是否越出所属卡片  3) 是否越出页面"""
import sys
from PIL import ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

path = sys.argv[1] if len(sys.argv) > 1 else 'deck.pptx'
FP = 'C:/Windows/Fonts/msyh.ttc'
E = 914400.0
_c = {}


def font(pt):
    k = int(round(pt * 4))
    if k not in _c:
        _c[k] = ImageFont.truetype(FP, k)
    return _c[k]


def tw(t, pt):
    f = font(pt)
    w = f.getlength(t) if hasattr(f, 'getlength') else f.getsize(t)[0]
    return w / 4.0 / 72.0


def lines(t, pt, w):
    if not t:
        return 1
    ws, buf = [], ''
    for ch in t:
        if ch == ' ':
            if buf:
                ws.append(buf); buf = ''
            ws.append(' ')
        else:
            buf += ch
    if buf:
        ws.append(buf)
    n, cur = 1, 0.0
    for x in ws:
        wd = tw(x, pt)
        if x == ' ':
            cur += wd
            continue
        if cur + wd > w and cur > 0:
            n += 1; cur = wd
        else:
            cur += wd
    return n


def need_h(tf, w):
    h = 0.0
    for p in tf.paragraphs:
        if not p.runs:
            continue
        pt = max(r.font.size.pt if r.font.size else 12 for r in p.runs)
        seg = ''.join(r.text for r in p.runs)
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        h += pt * 1.30 * ls * lines(seg, pt, w) / 72.0
        h += (p.space_after.pt if p.space_after else 0) / 72.0
    return h


prs = Presentation(path)
bad = 0
for si, sl in enumerate(prs.slides, 1):
    cards = []
    for sh in sl.shapes:
        if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        w, h = sh.width / E, sh.height / E
        if w > 2.5 and h > 0.85:
            cards.append((sh.left / E, sh.top / E, w, h))
    for sh in sl.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        x, y = sh.left / E, sh.top / E
        w, h = sh.width / E, sh.height / E
        nh = need_h(sh.text_frame, w)
        # 1) 溢出自身文本框
        if nh > h + 0.015:
            print('  [文本框溢出] P%d y=%.2f h=%.2f need=%.2f %s'
                  % (si, y, h, nh, sh.text_frame.text[:34].replace('\n', '|')))
            bad += 1
        # 2) 越出所属卡片
        for cxx, cyy, cww, chh in cards:
            if cxx - 0.02 <= x and cyy - 0.02 <= y and \
               x < cxx + cww and y < cyy + chh and (cww * chh) < 90:
                if y + nh > cyy + chh - 0.03:
                    print('  [越出卡片] P%d 文本底=%.2f 卡片底=%.2f %s'
                          % (si, y + nh, cyy + chh, sh.text_frame.text[:34].replace('\n', '|')))
                    bad += 1
                if x + w > cxx + cww - 0.04:
                    print('  [越出卡片右] P%d 文本右=%.2f 卡片右=%.2f %s'
                          % (si, x + w, cxx + cww, sh.text_frame.text[:34].replace('\n', '|')))
                    bad += 1
                break
        # 3) 越出页面
        if y + nh > 7.42 or x + w > 13.25:
            print('  [越出页面] P%d right=%.2f bottom=%.2f %s'
                  % (si, x + w, y + nh, sh.text_frame.text[:34].replace('\n', '|')))
            bad += 1

print('版式问题总数:', bad)
