# -*- coding: utf-8 -*-
"""原图 -> 缩略图（长边缩放），用于降低逐张视觉识别的上下文开销。不改动原图。
用法:
  python downscale.py --src <原图文件夹> --dst <缩略图文件夹> [--max-edge 1000]
"""
import argparse, os
from PIL import Image

EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="源图片文件夹")
    ap.add_argument("--dst", required=True, help="缩略图输出文件夹")
    ap.add_argument("--max-edge", type=int, default=1000, help="长边最大像素，默认 1000")
    args = ap.parse_args()

    os.makedirs(args.dst, exist_ok=True)
    count = 0
    for f in sorted(os.listdir(args.src)):
        if f.lower().endswith(EXTS):
            sp = os.path.join(args.src, f)
            try:
                im = Image.open(sp)
                im.thumbnail((args.max_edge, args.max_edge))
                im.save(os.path.join(args.dst, f))
                count += 1
            except Exception as e:
                print("跳过:", f, e)
    print(f"缩略图生成完成: {count} 张 -> {args.dst}")


if __name__ == "__main__":
    main()
