# -*- coding: utf-8 -*-
"""按分类映射把原图复制重命名到输出目录，并生成分类对照清单。
原图始终保留不动（复制而非移动/删除）。

mapping JSON 格式（二选一）：
  1) {"原文件名": "类别"}           或  {"原文件名": ""}（保留原名）
  2) {"原文件名": {"category": "类别", "desc": "内容说明"}}   （推荐，可带说明）

用法:
  python apply_classification.py --src <原图文件夹> --dst <输出文件夹> --mapping <mapping.json>
"""
import argparse, json, os, csv, shutil


def parse_val(v):
    if isinstance(v, dict):
        return (v.get("category") or ""), (v.get("desc") or "")
    return (v or ""), ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="原始图片文件夹")
    ap.add_argument("--dst", required=True, help="分类输出文件夹")
    ap.add_argument("--mapping", required=True, help="分类映射 JSON 文件路径")
    args = ap.parse_args()

    with open(args.mapping, encoding="utf-8") as f:
        mapping = json.load(f)

    os.makedirs(args.dst, exist_ok=True)
    counters = {}
    rows = []
    for orig, val in mapping.items():
        cat, desc = parse_val(val)
        sp = os.path.join(args.src, orig)
        if not os.path.exists(sp):
            print("缺失原图，跳过:", orig)
            continue
        ext = os.path.splitext(orig)[1]
        if cat:
            counters[cat] = counters.get(cat, 0) + 1
            new_name = f"{cat}_{counters[cat]}{ext}"
        else:
            new_name = orig  # 保留原名
        shutil.copy2(sp, os.path.join(args.dst, new_name))
        rows.append((new_name, cat if cat else "(保留原名)", orig, desc))

    manifest = os.path.join(args.dst, "分类对照清单.csv")
    with open(manifest, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["新文件名", "类别", "原文件名", "图片内容说明"])
        for r in rows:
            w.writerow(r)
    print(f"完成：{len(rows)} 张 -> {args.dst}")
    print(f"清单：{manifest}")


if __name__ == "__main__":
    main()
