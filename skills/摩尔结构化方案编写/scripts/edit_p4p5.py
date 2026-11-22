# -*- coding: utf-8 -*-
"""对第4、5页做版式与内容调整（仅改这两页，其余不动）：
1) 业务需求绿色圆点、当前痛点红色圆点 与其文字行精确垂直居中对齐
2) 方案价值 KPI 数值/主张字号调大
3) 删除页脚「数据来源」与「摩尔元数·专题方案」标注
输出：同目录 _修订.pptx（规避预览锁文件）
"""
import sys
from pptx import Presentation
from pptx.util import Emu, Pt

SRC = "E:/02跟进项目/青岛-中核控制系统工程有限公司-核燃料核化工-20261016/售前资料/中核控制_生产品质设备专题方案_模板版.pptx"
OUT = "E:/02跟进项目/青岛-中核控制系统工程有限公司-核燃料核化工-20261016/售前资料/中核控制_生产品质设备专题方案_模板版_修订.pptx"

def IN(e): return e/914400.0
def EM(v): return Emu(int(round(v*914400)))

def para_center_rel(tf, idx):
    pos=0.0
    for i,p in enumerate(tf.paragraphs):
        if i>idx: break
        fs=max([(r.font.size.pt if r.font.size else 9) for r in p.runs] or [9])
        ln=p.line_spacing
        if ln is None: lh=fs*1.2
        elif isinstance(ln,float): lh=ln*fs
        else: lh=float(ln)
        sa=(p.space_after.pt if p.space_after is not None else 0.0)
        if i==idx: return pos + (lh/72.0)/2.0
        pos += lh/72.0 + sa/72.0
    return pos

def find(slide,sid):
    for sh in slide.shapes:
        if sh.shape_id==sid: return sh
    return None

def align_bullets(slide, tb_id, dot_ids):
    tb=find(slide,tb_id); dots=[d for d in (find(slide,x) for x in dot_ids) if d]
    dots.sort(key=lambda s:s.top)
    for k,dot in enumerate(dots):
        if k>=len(tb.text_frame.paragraphs): break
        c_rel=para_center_rel(tb.text_frame,k)
        newtop=IN(tb.top)+c_rel-IN(dot.height)/2.0
        dot.top=EM(newtop)

def align_manual(slide, pairs):
    for tb_id,dot_id in pairs:
        tb=find(slide,tb_id); dot=find(slide,dot_id)
        newtop=IN(tb.top)+para_center_rel(tb.text_frame,0)-IN(dot.height)/2.0
        dot.top=EM(newtop)

def enlarge(slide, ids, newpt):
    for sh in slide.shapes:
        if sh.shape_id in ids and sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs: r.font.size=Pt(newpt)

def remove(slide, ids):
    for sh in list(slide.shapes):
        if sh.shape_id in ids:
            sp=sh._element; sp.getparent().remove(sp)

prs=Presentation(SRC)
slides=list(prs.slides)

# ---- PAGE 4 (index 3) ----
s4=slides[3]
align_bullets(s4, 139, [140,141,142,143])   # 业务需求 绿点
align_bullets(s4, 150, [151,152,153,154])   # 当前痛点 红点
enlarge(s4, [239,242,245,248], 18)          # KPI 数值 15->18
enlarge(s4, [236], 12)                      # 主张句 9.5->12
enlarge(s4, [240,243,246,249], 9.5)         # 标签 8.5->9.5
remove(s4, [251,252])                       # 页脚 数据来源 / 品牌

# ---- PAGE 5 (index 4) ----
s5=slides[4]
align_bullets(s5, 235, [236,237,238,239])   # 业务需求 绿点
align_manual(s5, [(247,246),(249,248),(251,250),(253,252)])  # 当前痛点 红点
enlarge(s5, [261,264,267,270], 16)          # KPI 数值 13->16
enlarge(s5, [262,265,268,271], 9)           # 标签 7.5->9
remove(s5, [273,274])                       # 页脚

prs.save(OUT)
print("SAVED", OUT)
