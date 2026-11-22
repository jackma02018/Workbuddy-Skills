# -*- coding: utf-8 -*-
"""build_deck — 摩尔结构化方案 PPT 生成入口

用法（Windows，务必用绝对路径调 python）：
    python build_deck.py content.json
    python build_deck.py content.json --out E:/xx/xxx.pptx --no-check
    python build_deck.py --dump-example          # 导出内容模板 content_template.json

    # V2.0 模板模式：沿用客户上传模板的封面/目录/封底/过渡页/标题/Logo 位置（零改动），
    # 仅把模板页眉组克隆到新增内容页；内容页用 deck_kit 四模块骨架渲染。
    python build_deck.py content.json --template "E:/02跟进项目/PPT模板.pptx" \
          --out E:/xx/方案_模板版.pptx [--dividers dividers.json]

内容 JSON 结构见 references/content-schema.md。
引擎（deck_kit.py）与内容解耦：换客户、换行业只改 JSON，不动代码。
V2.0 默认：圆点精确对齐文字、价值字号加大、页脚数据来源/品牌标注不绘制
（内容 JSON 设 meta.footer_anno=true 并填 source/brand 可恢复）。
"""
import io
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import deck_kit as K                                    # noqa: E402
from pptx import Presentation                           # noqa: E402
from pptx.util import Inches                            # noqa: E402
from pptx.dml.color import RGBColor                     # noqa: E402

BONE_ORDER = ['A', 'B', 'C', 'D', 'E']
HERE = os.path.dirname(os.path.abspath(__file__))


def to_rgb(v):
    """'1F6FEB' / '#1F6FEB' / [31,111,235] -> RGBColor；已是 RGBColor 则原样返回"""
    if isinstance(v, RGBColor):
        return v
    if isinstance(v, (list, tuple)) and len(v) == 3:
        return RGBColor(int(v[0]), int(v[1]), int(v[2]))
    s = str(v).lstrip('#').strip()
    if len(s) != 6:
        raise ValueError('颜色值非法（需 6 位十六进制）：%r' % v)
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def norm_page(p, i):
    """规范化单页内容：补默认骨架、颜色转 RGBColor、元组化，并做结构校验"""
    p = dict(p)
    p['bone'] = p.get('bone') or BONE_ORDER[(i - 1) % len(BONE_ORDER)]
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
    # 卡片支持 (图标, 标题) 或 (图标, 标题, 说明) 两种写法
    cards = []
    for c in alt.get('cards', []):
        c = list(c) + [''] * (3 - len(c))
        cards.append((c[0], c[1], c[2]))
    alt['cards'] = cards
    p['alt'] = alt

    icons = list(p.get('flow_icons') or [])
    # 图标数与流程步数不一致时自动补齐/截断，避免渲染期索引越界
    while len(icons) < len(fl['steps']):
        icons.append('check')
    p['flow_icons'] = icons[:len(fl['steps'])] or ['check']

    # 结构校验：四固定模块 + 双视角方案
    missing = [k for k in ('title', 'sub', 'need', 'pain', 'flow', 'alt', 'kpis')
               if not p.get(k)]
    if missing:
        raise ValueError('第 %d 页缺字段：%s' % (i, missing))
    if len(p['kpis']) not in (3, 4):
        raise ValueError('第 %d 页 kpis 需 3 或 4 项，当前 %d' % (i, len(p['kpis'])))
    return p


def load_content(path):
    with io.open(path, encoding='utf-8') as f:
        cfg = json.load(f)
    meta = cfg.get('meta', {})
    pages = [norm_page(p, i) for i, p in enumerate(cfg['pages'], 1)]
    for k, v in meta.items():
        if k in ('total', 'corner', 'source', 'brand'):
            K.META[k] = v
    K.META['total'] = len(pages)
    return meta, pages


def build(meta, pages, out):
    prs = Presentation()
    prs.slide_width = Inches(K.SW)
    prs.slide_height = Inches(K.SH)
    for i, d in enumerate(pages, 1):
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        K.header(sl, '%02d' % i, d['title'], d['sub'])
        K.BONES[d['bone']][0](sl, d)
        K.footer(sl)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    prs.save(out)
    return len(pages)


def run_check(pptx_path):
    """调用配套版式体检脚本，返回 (问题数, 输出文本)"""
    script = os.path.join(HERE, 'check_layout.py')
    r = subprocess.run([sys.executable, script, pptx_path],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    txt = (r.stdout or '') + (r.stderr or '')
    n = 0
    for line in txt.splitlines():
        if line.startswith('版式问题总数:'):
            try:
                n = int(line.split(':')[-1].strip())
            except ValueError:
                n = -1
    return n, txt.strip()


EXAMPLE = {
    "meta": {
        "corner": "客户名 · MES 建设方案",
        "footer_anno": False,
        "source": "数据来源：现场调研记录 ｜ 指标为行业基准示意口径，正式立项以现场调研数据替换",
        "brand": "摩尔元数 · MES 方案",
        "out": "E:/xx/MES专题方案.pptx"
    },
    "pages": [
        {
            "bone": "A",
            "title": "页面主标题（结论句，不要写“需求分析”这类栏目名）",
            "sub": "一句话副标题：承接什么、驱动什么",
            "need": ["业务需求 1（≤20 字）", "业务需求 2", "业务需求 3", "业务需求 4"],
            "pain": [["痛点短标签", "可归因描述（含数量/频次）"],
                     ["痛点短标签", "可归因描述"],
                     ["痛点短标签", "可归因描述"],
                     ["痛点短标签", "可归因描述"]],
            "flow": {
                "name": "流程型思维",
                "sub": "端到端主流程 · 六步闭环",
                "steps": [["步骤名", "一句话说明"], ["步骤名", "一句话说明"],
                          ["步骤名", "一句话说明"], ["步骤名", "一句话说明"],
                          ["步骤名", "一句话说明"], ["步骤名", "一句话说明"]],
                "note": "关键卡控点：×× · ×× · ××"
            },
            "alt": {
                "name": "第二视角（系统思维 / 数据驱动思维 / 精益思维…）",
                "sub": "该视角的解释框架",
                "color": "5B5FC6",
                "cards": [["gear", "卡片标题", "一句话说明"],
                          ["chip", "卡片标题", "一句话说明"],
                          ["network", "卡片标题", "一句话说明"],
                          ["calc", "卡片标题", "一句话说明"]]
            },
            "flow_icons": ["doc", "split", "check", "people", "scan", "arrow"],
            "kpis": [["95%", "指标名"], ["↓60%", "指标名"], ["3 → 1", "指标名"], ["1 套", "指标名"]],
            "claim": "本页方案结论（一句话，收口到经营收益）"
        }
    ]
}


def main():
    args = [a for a in sys.argv[1:]]
    if '--dump-example' in args:
        p = os.path.join(os.getcwd(), 'content_template.json')
        with io.open(p, 'w', encoding='utf-8') as f:
            f.write(json.dumps(EXAMPLE, ensure_ascii=False, indent=2))
        print('已导出模板:', p)
        return
    if not args:
        print(__doc__)
        return
    # -------- V2.0 模板模式 --------
    if '--template' in args:
        from build_template_deck import build as tpl_build, DEFAULT_TPL
        tpl = args[args.index('--template') + 1] if '--template' in args else DEFAULT_TPL
        src = args[0]
        out = args[args.index('--out') + 1] if '--out' in args else \
            os.path.splitext(os.path.abspath(src))[0] + '_模板版.pptx'
        dividers = None
        if '--dividers' in args:
            dv = args[args.index('--dividers') + 1]
            dividers = json.load(io.open(dv, encoding='utf-8')) if os.path.exists(dv) \
                else json.loads(dv)
        n = tpl_build(tpl, src, out, dividers=dividers)
        print('SAVED %s  content_pages=%d' % (out, n))
        if '--no-check' not in args:
            bad, txt = run_check(out)
            print(txt)
            print('>>> 版式体检：%s' % ('通过（0 问题）' if bad == 0
                  else '发现 %d 处问题（封面/封底若为模板原稿自带旗标属正常）' % bad))
        return
    # -------- 原自包含模式 --------
    src = args[0]
    out = None
    if '--out' in args:
        out = args[args.index('--out') + 1]
    do_check = '--no-check' not in args

    meta, pages = load_content(src)
    out = out or meta.get('out')
    if not out:
        out = os.path.splitext(os.path.abspath(src))[0] + '.pptx'
    if not os.path.isabs(out):
        out = os.path.abspath(out)

    n = build(meta, pages, out)
    print('SAVED %s  pages=%d  size=%d bytes' % (out, n, os.path.getsize(out)))
    print('骨架序列: ' + ' → '.join('%02d:%s' % (i, p['bone']) for i, p in enumerate(pages, 1)))
    if do_check:
        bad, txt = run_check(out)
        print(txt)
        print('>>> 版式体检：%s' % ('通过（0 问题）' if bad == 0 else '发现 %d 处问题，需修 JSON 文案长度或调骨架' % bad))


if __name__ == '__main__':
    main()
