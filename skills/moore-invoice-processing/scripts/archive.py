#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把已重命名的票据按文件名里的业务日期归入对应差旅日期文件夹。

前置：文件已按 <YYYYMMDD>[-<YYYYMMDD>]_<类型>_<补充>_<金额>.<ext> 命名。
日期直接取自文件名首位，不再二次解析 PDF。

用法:
    python archive.py --root "D:\\00费用报销\\未报销" --source "D:/.../未报销/_流程/待处理"
    python archive.py --root ... --source ... --execute      # 默认只预览

归档规则:
- 顶层形如 "20260911~20160913 总部售前培训" 的文件夹解析为日期区间 [fs, fe]。
- 文件名里的业务日期区间与之**取交集**，唯一命中即移入。
- 0 命中 -> 留在原地并列出（可能是新行程，需新建文件夹）。
- 多命中 -> 列出待人工判断，不擅自移动。
- 目标已存在同名文件时加序号，绝不覆盖。
"""
import argparse
import os
import re
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

DATE_RANGE = re.compile(r"^(20\d{6})(?:-(20\d{6}))?")
SPLIT = re.compile(r"[~～\-—]")


def norm_date(s):
    d = re.sub(r"\D", "", s or "")
    return d if len(d) == 8 else None


def folder_ranges(root):
    """顶层日期文件夹 -> [(name, path, start, end)]。兼容 '20260911~20110913' 年份笔误。"""
    ranges = []
    for name in sorted(os.listdir(root)):
        full = os.path.join(root, name)
        if not os.path.isdir(full) or name[0] in "_.":
            continue
        dates = [norm_date(p) for p in SPLIT.split(name.split()[0])]
        dates = [d for d in dates if d]
        if not dates:
            continue
        start = dates[0]
        end = dates[1] if len(dates) > 1 else dates[0]
        if len(dates) > 1:
            cand = dates[1]
            # 第二个日期年份写错（如 20110913 应为 20260913）时按起始年份校正
            if not cand.startswith(start[:4]) and cand[4:] >= start[4:]:
                cand = start[:4] + cand[4:]
            end = cand
        ranges.append((name, full, start, end))
    return ranges


def unique_path(folder, name):
    base, ext = os.path.splitext(name)
    cand, n = name, 1
    while os.path.exists(os.path.join(folder, cand)):
        n += 1
        cand = "%s_%d%s" % (base, n, ext)
    return os.path.join(folder, cand)


def main():
    ap = argparse.ArgumentParser(description="按业务日期把票据归入差旅文件夹")
    ap.add_argument("--root", required=True, help="报销根目录（内含日期文件夹）")
    ap.add_argument("--source", required=True, help="待归档文件所在目录")
    ap.add_argument("--execute", action="store_true", help="实际移动，默认只预览")
    args = ap.parse_args()

    ranges = folder_ranges(args.root)
    print("=" * 78)
    print("日期文件夹 %d 个：" % len(ranges))
    for name, _, fs, fe in ranges:
        print("   %-36s [%s ~ %s]" % (name, fs, fe))
    print("=" * 78)

    moved, pending, ambiguous, nodate = [], [], [], []
    for fn in sorted(os.listdir(args.source)):
        src = os.path.join(args.source, fn)
        if not os.path.isfile(src):
            continue
        m = DATE_RANGE.match(fn)
        if not m:
            nodate.append(fn)
            continue
        fs = m.group(1)
        fe = m.group(2) or fs
        hits = [(n, p) for n, p, ds, de in ranges if max(fs, ds) <= min(fe, de)]
        if not hits:
            pending.append((fn, fs, fe))
        elif len(hits) > 1:
            ambiguous.append((fn, [h[0] for h in hits]))
        else:
            dst_dir = hits[0][1]
            dst = unique_path(dst_dir, fn)
            if args.execute:
                shutil.move(src, dst)
            moved.append((fn, hits[0][0]))

    mode = "已移动" if args.execute else "待移动"
    print("\n%s %d 个：" % (mode, len(moved)))
    for fn, folder in moved:
        print("   %-58s -> %s" % (fn, folder))

    if pending:
        print("\n未匹配到日期文件夹 %d 个（建议新建文件夹或人工判断）：" % len(pending))
        for fn, fs, fe in pending:
            print("   - %-56s [%s ~ %s]" % (fn, fs, fe))
    if ambiguous:
        print("\n命中多个文件夹，需人工判断 %d 个：" % len(ambiguous))
        for fn, names in ambiguous:
            print("   ? %-56s -> %s" % (fn, names))
    if nodate:
        print("\n文件名缺少日期前缀（应先用 rename_invoices.py 重命名）%d 个：" % len(nodate))
        for fn in nodate[:20]:
            print("   !", fn)

    print("\n汇总：%s %d / 待确认 %d / 多命中 %d / 无日期 %d"
          % (mode, len(moved), len(pending), len(ambiguous), len(nodate)))
    if not args.execute:
        print("（预览模式，未移动任何文件。加 --execute 执行）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
