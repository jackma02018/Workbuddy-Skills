#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""解压邮件附件压缩包，只保留票证文件（PDF / 图片 / 可选 OFD），摊平到目标目录。

用法:
    python extract_zip.py --src ./原始附件 --dst ./待处理 [--keep-ofd]
                          [--flat/--no-flat] [--pattern zip]

要点:
- zip 内的中文文件名常用 cp437 伪装，需还原成 GBK（失败再退 UTF-8）。
- 同名文件自动加序号 _2，绝不覆盖。
- 非压缩文件（如直接的 PDF 附件）原样复制过去，保证输入可以是混合目录。
"""
import argparse
import os
import shutil
import sys
import zipfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

TICKET_EXT = {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}


def fix_name(fn):
    """还原 zip 里被 cp437 编码的中文文件名。"""
    try:
        fn.encode("cp437")
    except UnicodeEncodeError:
        return fn
    for enc in ("gbk", "utf-8"):
        try:
            return fn.encode("cp437").decode(enc)
        except Exception:
            continue
    return fn


def unique_path(dst_dir, name):
    base, ext = os.path.splitext(name)
    cand, n = name, 1
    while os.path.exists(os.path.join(dst_dir, cand)):
        n += 1
        cand = "%s_%d%s" % (base, n, ext)
    return os.path.join(dst_dir, cand)


def main():
    ap = argparse.ArgumentParser(description="解压附件并摊平票证文件")
    ap.add_argument("--src", required=True, help="原始附件目录（含 zip / 直接附件）")
    ap.add_argument("--dst", required=True, help="输出目录")
    ap.add_argument("--keep-ofd", action="store_true", help="保留 .ofd（默认丢弃）")
    ap.add_argument("--no-flat", action="store_true", help="保留 zip 内的子目录结构")
    ap.add_argument("--pattern", default="zip", help="压缩包扩展名，逗号分隔，默认 zip")
    args = ap.parse_args()

    keep = set(TICKET_EXT)
    if args.keep_ofd:
        keep.add(".ofd")
    zips = {e.lower().lstrip(".") for e in args.pattern.split(",") if e.strip()}

    os.makedirs(args.dst, exist_ok=True)
    kept, skipped, failed = [], [], []

    for name in sorted(os.listdir(args.src)):
        src = os.path.join(args.src, name)
        if os.path.isdir(src):
            continue
        ext = os.path.splitext(name)[1].lower().lstrip(".")

        if ext in zips:
            try:
                with zipfile.ZipFile(src) as zf:
                    for info in zf.infolist():
                        inner = fix_name(info.filename)
                        if inner.endswith("/"):
                            continue
                        iext = os.path.splitext(inner)[1].lower()
                        if iext not in keep:
                            skipped.append("%s!%s" % (name, inner))
                            continue
                        rel = inner if args.no_flat else os.path.basename(inner)
                        target = unique_path(args.dst, rel) if args.no_flat else None
                        if args.no_flat:
                            os.makedirs(os.path.dirname(target), exist_ok=True)
                        else:
                            target = unique_path(args.dst, os.path.basename(inner))
                        with zf.open(info) as s, open(target, "wb") as d:
                            d.write(s.read())
                        kept.append(os.path.basename(target))
            except Exception as exc:
                failed.append("%s: %s" % (name, exc))
            continue

        if "." + ext in keep:
            shutil.copy2(src, unique_path(args.dst, name))
            kept.append(name)
        else:
            skipped.append(name)

    print("=" * 72)
    print("解压完成：保留 %d / 丢弃 %d / 失败 %d" % (len(kept), len(skipped), len(failed)))
    print("输出目录: %s" % args.dst)
    print("=" * 72)
    if skipped:
        print("已丢弃（非票证格式）:")
        for s in skipped[:30]:
            print("   -", s)
    if failed:
        print("解压失败（需人工处理）:")
        for f in failed:
            print("   !!", f)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
