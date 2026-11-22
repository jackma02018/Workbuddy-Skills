# -*- coding: utf-8 -*-
"""从票面 PDF 提取：类别 / 业务日期范围 / 金额。

支持三类票面：12306 火车票发票（含退票）、网约车电子发票、网约车行程报销单。
必须按坐标还原版式，否则「价税合计」会被 get_text("text") 拆散而取错金额。

用法:
    python parse_invoices.py --train <火车票pdf目录> --ride <网约车pdf目录> --json parsed.json
    python parse_invoices.py --dirs a/ b/ c/ --json parsed.json
"""
import argparse
import fitz
import glob
import json
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

DATE_CN = re.compile(r"(\d{4})年(\d{2})月(\d{2})日")
DATE_ISO = re.compile(r"(\d{4})-(\d{2})-(\d{2})")


def lines_of(page):
    """按 (y, x) 还原阅读顺序，返回 [(y, [(x0, text), ...]), ...]"""
    words = page.get_text("words")
    words.sort(key=lambda w: (round(w[1], 1), round(w[0], 1)))
    lines, cur_y, buf = [], None, []
    for w in words:
        y = round(w[1], 1)
        if cur_y is None or abs(y - cur_y) > 3:
            if buf:
                lines.append((cur_y, buf))
            buf, cur_y = [], y
        buf.append((round(w[0]), w[4]))
    if buf:
        lines.append((cur_y, buf))
    return lines


def join(line):
    return "".join(t for _, t in line)


def num_val(tok):
    """兼容 '¥60.42' / '￥60.42' / '1,234.00' 等写法。"""
    t = tok.replace("¥", "").replace("￥", "").replace(",", "").strip()
    return float(t) if re.fullmatch(r"-?\d+\.\d{2}", t) else None


def find_date_of_year(line):
    m = DATE_CN.search(join(line))
    return "%s-%s-%s" % m.groups() if m else None


def parse_train(path):
    doc = fitz.open(path)
    lines = lines_of(doc[0])
    doc.close()
    text = "\n".join(join(l) for _, l in lines)
    info = {"kind": "train"}

    m = re.search(r"发票号码:(\d+)", text)
    info["invoice_no"] = m.group(1) if m else None

    ride = None
    for _, ln in lines:                       # 跳过「开票日期」行，首个年月日即乘车日期
        if "开票日期" in join(ln):
            continue
        d = find_date_of_year(ln)
        if d:
            ride = d
            break
    info["biz_date"] = ride

    info["category"] = "退票" if "退票费" in text else ("火车发票" if "票价" in text else "未知")

    m = re.search(r"￥\s*(-?\d+\.\d{2})", text)
    info["amount"] = float(m.group(1)) if m else None

    stations, after_no = [], False
    for _, ln in lines:
        s = join(ln)
        if after_no and re.fullmatch(r"[\u4e00-\u9fa5]+站", s):
            stations.append(s)
        if s.startswith("发票号码"):
            after_no = True
    info["from_station"] = stations[0] if len(stations) > 0 else None
    info["to_station"] = stations[1] if len(stations) > 1 else None

    m = re.search(r"^([A-Z]\d+)$", text, re.M)
    info["train_no"] = m.group(1) if m else None
    return info


def parse_ride_invoice(path):
    doc = fitz.open(path)
    lines = lines_of(doc[0])
    doc.close()
    text = "\n".join(join(l) for _, l in lines)
    info = {"kind": "ride_invoice"}

    for _, ln in lines:
        s = join(ln)
        if "发票号码" in s:
            m = re.search(r"(\d{15,})", s)
            if m:
                info["invoice_no"] = m.group(1)
                break

    seller = None                             # 销售方在右列，取 x 更大的「名称：」
    for _, ln in lines:
        if "名称：" in join(ln):
            cand = [(x, t) for x, t in ln if t.startswith("名称：")]
            if cand:
                _, t = max(cand, key=lambda c: c[0])
                seg = t.split("名称：", 1)[1]
                if seg:
                    seller = seg
    info["seller"] = seller

    all_dates = sorted(set(m.group(0) for _, ln in lines for m in [DATE_ISO.search(join(ln))] if m))
    body_dates, in_table = [], False
    for _, ln in lines:
        s = join(ln)
        if "出行人" in s and "出行日期" in s:
            in_table = True
            continue
        if in_table:
            m = DATE_ISO.search(s)
            if m:
                body_dates.append(m.group(0))
    info["travel_dates"] = sorted(set(body_dates))
    info["biz_date"] = info["travel_dates"][0] if info["travel_dates"] else None
    info["biz_date_end"] = info["travel_dates"][-1] if info["travel_dates"] else None
    info["all_dates"] = all_dates

    total = None
    for _, ln in lines:                       # 「（小写）」行 x 最大的数值即价税合计
        if "（小写）" in join(ln):
            vals = [(x, num_val(t)) for x, t in ln]
            vals = [(x, v) for x, v in vals if v is not None]
            if vals:
                total = max(vals, key=lambda c: c[0])[1]
    info["amount"] = total

    seller = seller or ""
    if "滴滴" in seller:
        info["category"] = "网约车发票"
    elif "曹操" in seller or "吉利优行" in seller:
        info["category"] = "网约车发票"
    else:
        info["category"] = "网约车发票"
    return info


def parse_trip_sheet(path):
    doc = fitz.open(path)
    lines = lines_of(doc[0])
    doc.close()
    text = "\n".join(join(l) for _, l in lines)
    info = {"kind": "trip_sheet"}

    m = re.search(r"行程起止日期：(\d{4}-\d{2}-\d{2})\s*至\s*(\d{4}-\d{2}-\d{2})", text)
    if m:
        info["biz_date"], info["biz_date_end"] = m.group(1), m.group(2)
    else:
        ds = sorted(set(DATE_ISO.search(join(l)).group(0) for _, l in lines if DATE_ISO.search(join(l))))
        info["biz_date"], info["biz_date_end"] = (ds[0], ds[-1]) if ds else (None, None)

    m = re.search(r"合计\s*([\d.]+)\s*元", text)
    info["amount"] = float(m.group(1)) if m else None

    m = re.search(r"共\s*(\d+)\s*笔行程", text)
    info["trip_count"] = int(m.group(1)) if m else None
    info["category"] = "网约车行程"
    return info


def classify_and_parse(path):
    doc = fitz.open(path)
    first = doc[0].get_text("text")
    doc.close()
    if "电子客票号" in first:
        return parse_train(path)
    if "行程起止日期" in first or "行程单" in first or "行程报销单" in first:
        return parse_trip_sheet(path)
    return parse_ride_invoice(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--train", action="append", default=[], help="12306 PDF 目录（可重复）")
    ap.add_argument("--ride", action="append", default=[], help="网约车 PDF 目录（可重复）")
    ap.add_argument("--dirs", nargs="*", default=[], help="任意 PDF 目录")
    ap.add_argument("--json", required=True, help="输出 parsed.json")
    args = ap.parse_args()

    dirs = args.train + args.ride + args.dirs
    files = sorted(f for d in dirs for f in glob.glob(os.path.join(d, "**", "*.pdf"), recursive=True))
    if not files:
        sys.exit("没有找到 PDF，检查 --train / --ride / --dirs 参数")

    out = []
    for f in files:
        try:
            d = classify_and_parse(f)
        except Exception as e:
            d = {"kind": "error", "error": repr(e)}
        d["path"] = f
        d["basename"] = os.path.basename(f)
        out.append(d)

    # 发票票面无「出行日期」时，用等额行程单反推行程日期
    sheets = [d for d in out if d.get("kind") == "trip_sheet"]
    for d in out:
        if d.get("kind") == "ride_invoice" and not d.get("biz_date") and d.get("amount") is not None:
            for s in sheets:
                if s.get("amount") is not None and abs(s["amount"] - d["amount"]) < 0.005:
                    d["biz_date"] = s.get("biz_date")
                    d["biz_date_end"] = s.get("biz_date_end")
                    d["date_from"] = "按等额行程单 %s 反推" % s.get("basename")
                    break

    json.dump(out, open(args.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    print("| 文件 | 类别 | 起 | 止 | 金额 | 备注 |")
    print("|" + "---|" * 6)
    for d in out:
        print("| %s | %s | %s | %s | %s | %s |" % (
            d.get("basename"), d.get("category"), d.get("biz_date"),
            d.get("biz_date_end") or "-", d.get("amount"),
            (d.get("train_no") or d.get("seller") or d.get("invoice_no") or d.get("error", ""))))
    print("\n文件数: %d    金额合计: %.2f" % (len(out), sum(d.get("amount") or 0 for d in out)))


if __name__ == "__main__":
    main()
