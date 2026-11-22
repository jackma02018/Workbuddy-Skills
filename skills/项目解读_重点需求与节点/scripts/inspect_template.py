# -*- coding: utf-8 -*-
"""
解析任意 .potx / .pptx 的版式结构、占位符与主题色/字体。
用法：python inspect_template.py <模板.potx>
输出：版式清单(名称+占位符索引/类型)、主题色(hex)、默认字体。
换模板时先跑它，确认 build_deck.py 的 get_layout 关键字（标题幻灯片/标题和内容/仅标题）能命中。
"""
import sys, zipfile
from pptx import Presentation
from pptx.util import Emu

def main():
    if len(sys.argv) < 2:
        print("用法: python inspect_template.py <模板.potx>"); return
    path = sys.argv[1]
    # .potx 需先修正 content-type 才能被 python-pptx 打开
    tmp = path + "._insp.pptx"
    with zipfile.ZipFile(path, "r") as zin:
        data = {n: zin.read(n) for n in zin.namelist()}
    ct = data["[Content_Types].xml"].decode("utf-8").replace(
        "presentationml.template.main+xml", "presentationml.presentation.main+xml")
    data["[Content_Types].xml"] = ct.encode("utf-8")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in data.items(): zout.writestr(n, b)

    prs = Presentation(tmp)
    print("=== 幻灯片尺寸 ===")
    print(f"  {Emu(prs.slide_width).inches:.2f} x {Emu(prs.slide_height).inches:.2f} 英寸 "
          f"({'16:9' if abs(Emu(prs.slide_width).inches/Emu(prs.slide_height).inches-1.777)<0.02 else '其他'})")
    print("=== 版式清单 ===")
    for i, lay in enumerate(prs.slide_layouts):
        phs = []
        for ph in lay.placeholders:
            phs.append(f"[{ph.placeholder_format.idx}] {ph.name}({ph.placeholder_format.type})")
        print(f"  L{i}: {lay.name}  ->  " + " | ".join(phs))
    print("=== 主题色 ===")
    try:
        for c in prs.slide_masters[0].element.clrMap:
            pass
    except Exception:
        pass
    # 直接从 theme1.xml 读 accent 颜色
    theme = None
    for part in prs.part.package.iter_parts():
        if part.partname.endswith("theme1.xml"):
            theme = part.blob.decode("utf-8"); break
    if theme:
        import re
        for m in re.finditer(r'<a:(\w+)\s+val="(#[0-9A-Fa-f]{6})"', theme):
            name, val = m.group(1), m.group(2)
            if name.startswith("dk") or name.startswith("lt") or name.startswith("accent"):
                print(f"  {name}: {val}")
    print("=== 默认字体 ===")
    try:
        for part in prs.part.package.iter_parts():
            if part.partname.endswith("theme1.xml"):
                t = part.blob.decode("utf-8")
                import re
                lat = re.search(r'<a:latin typeface="([^"]+)"', t)
                ea = re.search(r'<a:ea typeface="([^"]+)"', t)
                print(f"  latin: {lat.group(1) if lat else '?' } | ea(中): {ea.group(1) if ea else '?'}")
                break
    except Exception as e:
        print("  读取失败:", e)
    try:
        import os; os.remove(tmp)
    except Exception: pass

if __name__ == "__main__":
    main()
