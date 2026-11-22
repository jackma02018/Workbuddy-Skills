# -*- coding: utf-8 -*-
"""对 分类对照清单.csv 中带名表单做规范化，输出更改前后对照清单。
内置 MOM/ISO 制造行业默认规则，可用 --rules / --semantic JSON 覆盖或扩充。

用法:
  python normalize_forms.py --manifest <分类对照清单.csv> --out <输出.csv> [--rules rules.json] [--semantic sem.json]
"""
import argparse, json, os, csv

# 内置默认规则（制造/MOM 行业）；可被 --rules JSON 覆盖或扩充
DEFAULT_RULES = {
    "首件检验记录表": "首件检验记录表[FAI]",
    "产品点检记录(电性测试)": "产品点检记录表(电性测试)",
    "ECR变更追踪表(IQC源料状态)": "ECN变更追踪表(IQC原料状态)",
    "红胶储存/使用管理制表": "红胶储存/使用管理表",
}
# 通用后缀/错别字修正（命中即替换）
SUFFIX_FIXES = [
    ("管理制表", "管理表"),
]
# 语义定位：首件检验 / NCR / ECN 三项；可被 --semantic JSON 覆盖或扩充
DEFAULT_SEMANTIC = {
    "首件检验记录表": "首件检验",
    "ECR变更追踪表(IQC源料状态)": "ECN",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--manifest", required=True, help="分类对照清单.csv")
    ap.add_argument("--out", required=True, help="输出对照清单.csv")
    ap.add_argument("--rules", default="", help="可选 JSON：{原表单名: 新表单名}")
    ap.add_argument("--semantic", default="", help="可选 JSON：{原表单名: 语义类别}")
    args = ap.parse_args()

    rules = dict(DEFAULT_RULES)
    semantic = dict(DEFAULT_SEMANTIC)
    if args.rules and os.path.exists(args.rules):
        rules.update(json.load(open(args.rules, encoding="utf-8")))
    if args.semantic and os.path.exists(args.semantic):
        semantic.update(json.load(open(args.semantic, encoding="utf-8")))

    rows = []
    with open(args.manifest, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            name = row.get("图片内容说明") or ""
            new = rules.get(name)
            if new is None:
                for a, b in SUFFIX_FIXES:
                    if a in name:
                        new = name.replace(a, b)
                        break
            if new is None:
                new = name
            changed = "是" if new != name else "否"
            sem = semantic.get(name, "")
            rows.append((row.get("类别", ""), name, new, changed, sem, row.get("原文件名", "")))

    with open(args.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "类别", "原表单名称", "新表单名称", "是否更改",
                    "语义类别(首件检验/ECN/NCR)", "原文件名"])
        for i, r in enumerate(rows, 1):
            w.writerow([i, *r])
    print(f"表单记录: {len(rows)}，已写入 {args.out}")


if __name__ == "__main__":
    main()
