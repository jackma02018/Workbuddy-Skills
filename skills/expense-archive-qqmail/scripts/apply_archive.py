# -*- coding: utf-8 -*-
"""按「YYYYMMDD_类别_金额.pdf」重命名，并按业务日期范围归入对应差旅日期文件夹。

用法:
    python apply_archive.py --root <报销根目录> --parsed parsed.json [--today 20260915]
"""
import argparse
import datetime
import json
import os
import re
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

CATEGORY = {
    "train":        "火车发票",
    "ride_invoice": "网约车发票",
    "trip_sheet":   "网约车行程",
}


def norm_date(s):
    d = re.sub(r"\D", "", s or "")
    return d if len(d) == 8 else None


def folder_ranges(root):
    """顶层日期文件夹 -> (name, path, start, end)。兼容 '20260911~20110913' 年份笔误。"""
    ranges = []
    for name in sorted(os.listdir(root)):
        full = os.path.join(root, name)
        if not os.path.isdir(full) or name.startswith("_"):
            continue
        dates = [norm_date(p) for p in re.split(r"[~～\-—]", name.split()[0])]
        dates = [d for d in dates if d]
        if not dates:
            continue
        start = dates[0]
        end = dates[1] if len(dates) > 1 else dates[0]
        if len(dates) > 1:
            cand = dates[1]
            if not cand.startswith(start[:4]) and cand[4:] >= start[4:]:
                cand = start[:4] + cand[4:]      # 如 20110913 -> 20260913
            end = cand
        ranges.append((name, full, start, end))
    return ranges


def pick_folder(start, end, ranges):
    s, e = norm_date(start), norm_date(end or start)
    return [(n, p) for n, p, fs, fe in ranges if max(s, fs) <= min(e, fe)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="报销根目录（内含日期文件夹）")
    ap.add_argument("--parsed", required=True, help="parse_invoices.py 产出的 json")
    ap.add_argument("--today", help="文件名日期 YYYYMMDD，默认取今天")
    ap.add_argument("--dry-run", action="store_true", help="只预览不移动")
    args = ap.parse_args()

    today = args.today or datetime.date.today().strftime("%Y%m%d")
    items = json.load(open(args.parsed, encoding="utf-8"))
    ranges = folder_ranges(args.root)

    print("日期文件夹：")
    for name, _, fs, fe in ranges:
        print("   %s   [%s ~ %s]" % (name, fs, fe))

    plan, unmatched = [], []
    for d in items:
        cat = "退票" if d.get("category") == "退票" else CATEGORY.get(d.get("kind"))
        if cat is None or d.get("amount") is None or not d.get("biz_date"):
            unmatched.append(d)
            continue
        hits = pick_folder(d["biz_date"], d.get("biz_date_end"), ranges)
        newname = "%s_%s_%.2f.pdf" % (today, cat, d["amount"])
        if len(hits) == 1:
            plan.append((d, hits[0], newname))
        elif not hits:
            unmatched.append(d)
        else:
            print("!! 命中多个文件夹，需人工判断：%s -> %s" % (d.get("basename"), [h[0] for h in hits]))

    print("\n执行计划（今日日期 %s）：" % today)
    for d, (fname, fpath), newname in plan:
        dst = os.path.join(fpath, newname)
        flag = "  ** 目标已存在，将被替换" if os.path.exists(dst) else ""
        if not args.dry_run:
            shutil.move(d["path"], dst)
        print("   [%-5s] %-42s -> %s/%s%s" % (d.get("category"), d.get("basename"), fname, newname, flag))

    print("\n移动文件数：%d" % len(plan))
    if unmatched:
        print("未归档 %d 个：" % len(unmatched))
        for d in unmatched:
            print("   -", d.get("basename"), d.get("category"), d.get("biz_date"), d.get("amount"))

    print("\n归档后各文件夹存量：")
    for name, full, _, _ in ranges:
        n = len([f for f in os.listdir(full) if f.lower().endswith(".pdf")])
        print("   %-34s %2d 个 PDF%s" % (name, n, "   <-- 仍为空，注意是否漏开票" if n == 0 else ""))


if __name__ == "__main__":
    main()
