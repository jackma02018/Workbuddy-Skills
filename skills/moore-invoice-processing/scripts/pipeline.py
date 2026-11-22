#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""摩尔发票处理 —— 一键流水线：筛选邮件 → 下载附件 → 解压 → 重命名 → 归档 → 汇总。

每一步都可单独跑（对应同名脚本），本文件只负责编排与降级决策。

用法:
    # 全程预览，不做任何改动
    python pipeline.py --dest "D:\\00费用报销\\未报销" --after 2026-09-15 --template railway

    # 实际执行（重命名 + 归档）
    python pipeline.py --dest "D:\\00费用报销\\未报销" --after 2026-09-15 --template railway --execute

    # 只跑后半段（附件已经在本地了）
    python pipeline.py --dest ... --steps rename,archive --execute

步骤: fetch, download, extract, rename, archive
默认全跑。任一步失败即中止，已下载的中间产物保留在 <dest>/_流程 下，可断点续跑。
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile
import time

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
ALL_STEPS = ["fetch", "download", "extract", "rename", "archive"]


def run_step(step, script, args, workdir):
    print("\n" + "=" * 78)
    print("▶ 步骤 %s : %s" % (step, script))
    print("=" * 78)
    proc = subprocess.run([sys.executable, os.path.join(HERE, script)] + args,
                          cwd=workdir, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    if proc.stdout:
        print(proc.stdout.rstrip())
    if proc.stderr and proc.returncode != 0:
        print("--- stderr ---")
        print(proc.stderr.rstrip()[-2000:])
    return proc.returncode


def main():
    ap = argparse.ArgumentParser(description="发票处理流水线")
    ap.add_argument("--dest", required=True, help="报销根目录（内含日期文件夹）")
    ap.add_argument("--workdir", help="中间产物目录，默认 <dest>/_流程")
    ap.add_argument("--days", type=int, help="只处理最近 N 天（与 --after 二选一）")
    ap.add_argument("--after", help="起始时间，如 2026-09-15（服务端过滤）")
    ap.add_argument("--before", help="截止时间")
    ap.add_argument("--sender", help="发件人关键词/正则")
    ap.add_argument("--subject", default="发票|行程|电子客票", help="主题关键词，默认匹配常见发票邮件")
    ap.add_argument("--keyword", help="全文关键词/正则")
    ap.add_argument("--ext", default="zip,pdf", help="附件扩展名过滤，默认 zip,pdf")
    ap.add_argument("--skip-ofd", action="store_true", default=True, help="丢弃 .ofd（默认开启）")
    ap.add_argument("--alias", help="QQ 邮箱 alias_id，省略则自动探测")
    ap.add_argument("--template", default="railway", choices=["standard", "railway", "legacy12306"])
    ap.add_argument("--steps", default=",".join(ALL_STEPS))
    ap.add_argument("--execute", action="store_true", help="实际重命名并归档，默认只预览")
    ap.add_argument("--report", help="汇总 JSON 输出路径")
    args = ap.parse_args()

    workdir = args.workdir or os.path.join(args.dest, "_流程")
    raw_dir = os.path.join(workdir, "原始附件")
    ticket_dir = os.path.join(workdir, "待处理")
    manifest = os.path.join(workdir, "manifest.json")
    plan = os.path.join(workdir, "rename_plan.json")
    token = os.path.join(workdir, ".qqmail_token")
    for d in (workdir, raw_dir, ticket_dir):
        os.makedirs(d, exist_ok=True)

    steps = [s.strip() for s in args.steps.split(",") if s.strip()]
    unknown = [s for s in steps if s not in ALL_STEPS]
    if unknown:
        sys.exit("未知步骤: %s（可选 %s）" % (unknown, "/".join(ALL_STEPS)))

    if args.after and args.days:
        sys.exit("--after 与 --days 不能同时给")
    after = args.after
    if args.days:
        after = time.strftime("%Y-%m-%d", time.localtime(time.time() - args.days * 86400))

    print("工作目录: %s" % workdir)
    print("时间范围: %s ~ %s" % (after or "(不限)", args.before or "(不限)"))
    print("命名口径: %s   执行模式: %s" % (args.template, "实际改动" if args.execute else "预览"))

    codes = {}
    # ---------------- fetch
    if "fetch" in steps:
        cmd = ["--manifest", manifest, "--skip-ofd", "--subject", args.subject,
               "--ext", args.ext, "--token", token]
        for k, v in (("--after", after), ("--before", args.before),
                     ("--sender", args.sender), ("--keyword", args.keyword)):
            if v:
                cmd += [k, v]
        codes["fetch"] = run_step("fetch", "decrypt_token.py", ["--out", token], workdir)
        if codes["fetch"] != 0:
            sys.exit("取不到邮箱 token —— 若报 401，先调一次 MCP GetMe 触发续期再重试")
        codes["fetch"] = run_step("fetch", "fetch_mail.py", cmd, workdir)
        if codes["fetch"] != 0:
            sys.exit("邮件筛选失败，中止")

    # ---------------- download
    if "download" in steps:
        cmd = ["--manifest", manifest, "--out", raw_dir, "--token", token]
        if args.alias:
            cmd += ["--alias", args.alias]
        codes["download"] = run_step("download", "fetch_attachments.py", cmd, workdir)
        if os.path.exists(token):
            os.remove(token)                      # token 用完即删
            print("已清理 token: %s" % token)
        if codes["download"] != 0:
            print("!! 部分附件下载失败，剩余条目可用同一命令重跑补全")

    # ---------------- extract
    if "extract" in steps:
        cmd = ["--src", raw_dir, "--dst", ticket_dir]
        codes["extract"] = run_step("extract", "extract_zip.py", cmd, workdir)

    # ---------------- rename
    if "rename" in steps:
        cmd = ["--root", ticket_dir, "--template", args.template, "--json", plan]
        if args.execute:
            cmd.append("--execute")
        codes["rename"] = run_step("rename", "rename_invoices.py", cmd, workdir)

    # ---------------- archive
    if "archive" in steps:
        cmd = ["--root", args.dest, "--source", ticket_dir]
        if args.execute:
            cmd.append("--execute")
        codes["archive"] = run_step("archive", "archive.py", cmd, workdir)

    # ---------------- report
    summary = {"steps": codes, "template": args.template,
               "execute": args.execute, "workdir": workdir}
    if os.path.exists(plan):
        items = json.load(open(plan, encoding="utf-8"))
        ok = [i for i in items if i.get("new")]
        skipped = [i for i in items if not i.get("new")]
        total = 0.0
        for i in ok:
            try:
                total += float(i["new"].rsplit("_", 1)[1].rsplit(".", 1)[0])
            except Exception:
                pass
        summary.update({"ok": len(ok), "skipped": len(skipped),
                        "amount_total": round(total, 2),
                        "skipped_detail": [{"file": s["old"], "reason": s["skip_reason"]}
                                           for s in skipped]})
        print("\n" + "=" * 78)
        print("汇总：成功 %d / 跳过 %d / 金额合计 %.2f 元" % (len(ok), len(skipped), total))
        if skipped:
            print("跳过原因明细：")
            for s in skipped:
                print("   - %-52s %s" % (s["old"], s["skip_reason"]))
        print("=" * 78)

    if args.report:
        json.dump(summary, open(args.report, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
        print("汇总已写出: %s" % args.report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
