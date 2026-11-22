# -*- coding: utf-8 -*-
"""
对生成后的 pptx 做三项校验：① 包完整性 ② 几何溢出 ③ 内容抽查（编号/优先级/缩写/待确认）。
用法：python verify_deck.py <生成的.pptx> [--check-priority]   # --check-priority 时把含 P0/P1/P2 判为异常
"""
import sys, zipfile
from pptx import Presentation
from pptx.util import Emu

def main():
    if len(sys.argv) < 2:
        print("用法: python verify_deck.py <pptx> [--check-priority]"); return
    path = sys.argv[1]
    check_pri = "--check-priority" in sys.argv

    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
    print("① 包完整性:", "OK" if bad is None else f"BAD: {bad}")

    prs = Presentation(path)
    SW, SH = prs.slide_width, prs.slide_height
    ob = 0
    for si, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            try:
                if sh.left is None: continue
                if sh.left < 0 or sh.top < 0 or (sh.left + sh.width) > SW + 10000 or (sh.top + sh.height) > SH + 10000:
                    ob += 1; print(f"  ⚠ 溢出 S{si}: R={Emu(sh.left+sh.width).inches:.2f} B={Emu(sh.top+sh.height).inches:.2f}")
            except Exception: pass
    print(f"② 几何溢出: {'无' if ob==0 else f'{ob} 处'} (画布 {Emu(SW).inches:.2f}x{Emu(SH).inches:.2f}in)")

    # ③ 内容抽查：重点需求页（通常是第 2 页）
    print("③ 内容抽查（重点需求页）:")
    s2 = prs.slides[1]
    full = ""
    for sh in s2.shapes:
        if sh.has_text_frame:
            full += sh.text_frame.text + "\n"
    if check_pri:
        has_pri = any(x in full for x in ["P0", "P1", "P2"])
        print(f"  - 含 P0/P1/P2: {has_pri} {'⚠应去除' if has_pri else '✓'}")
    # 编号连续性检查（数字开头行）
    lines = [l.strip() for l in full.splitlines() if l.strip()[:1].isdigit()]
    nums = [int(l.split()[0]) for l in lines if l.split()[0].isdigit()]
    cont = nums == list(range(1, len(nums) + 1))
    print(f"  - 需求条数: {len(nums)}，编号连续: {'✓' if cont else '⚠'+str(nums)}")
    print(f"  - 含 ① 原圈码: {'⚠' if '①' in full else '✓无'}")
    print("  - 缩写(SMT/BOM/ERP/FIFO/ECN/Andon/IQC/SPC/CPK):",
          {a: (a in full) for a in ["SMT","BOM","ERP","FIFO","ECN","Andon","IQC","SPC","CPK"]})
    print(f"  - 待确认 标注出现次数: {full.count('待确认')}")
    print("Slide count:", len(prs.slides))

if __name__ == "__main__":
    main()
