#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_qa.py — MES 调研报告质检工具

能力：
  1) extract : 提取 docx 全文（段落 + 表格单元格），便于通读质检
  2) fix     : 跨 <w:t> run 边界的安全替换。docx 文本常被拆成多个 run，
               普通字符串替换会匹配不到；本工具用「每字符间允许
               </w:t>.*?<w:t>」的正则解决，并保留其余格式。

用法：
  python docx_qa.py extract <file.docx>
  python docx_qa.py fix --in <in.docx> --out <out.docx> --rules rules.json
  python docx_qa.py fix --in <in.docx> --out <out.docx> --pair "旧词||新词" [--pair ...]

rules.json 格式： [["旧词1","新词1"], ["旧词2","新词2"]]
"""
import sys
import os
import re
import json
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ----------------------------- 提取 -----------------------------
def get_document_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def _get_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def extract_all(path):
    import xml.etree.ElementTree as ET
    root = ET.fromstring(get_document_xml(path))
    body = root.find(W + "body")
    out = []
    for el in body:
        tag = el.tag.split("}")[-1]
        if tag == "p":
            out.append(_get_text(el))
        elif tag == "tbl":
            for row in el.findall(W + "tr"):
                cells = []
                for tc in row.findall(W + "tc"):
                    cells.append(" ".join(_get_text(p) for p in tc.findall(W + "p")))
                out.append(" | ".join(cells))
    return "\n".join(out)


# ----------------------------- 跨 run 替换 -----------------------------
def _escape_per_char(old):
    """为 old 构造允许 run 边界中断的正则：每个字符后可选 </w:t>.*?<w:t>。"""
    return "".join(re.escape(c) + r"(?:</w:t>.*?<w:t[^>]*>)?" for c in old)


def raw_replace(xml, old, new):
    pat = _escape_per_char(old)
    return re.sub(pat, new, xml, flags=re.DOTALL)


def count_match(xml, old):
    pat = _escape_per_char(old)
    return len(re.findall(pat, xml, flags=re.DOTALL))


def fix_docx(in_path, out_path, rules):
    with zipfile.ZipFile(in_path) as z:
        data = z.read("word/document.xml").decode("utf-8")
        for old, new in rules:
            cnt = count_match(data, old)
            data = raw_replace(data, old, new)
            print(f"  replace [{old}] -> [{new}] : {cnt} 处")
        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in z.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, data)
                else:
                    zout.writestr(item, z.read(item.filename))
    try:
        os.replace(tmp, out_path)
        print("SAVED", out_path)
    except PermissionError:
        alt = out_path + ".unlocked.docx"
        os.replace(tmp, alt)
        print(f"[WARN] 原文件被锁定，已另存到: {alt}")


# ----------------------------- CLI -----------------------------
def main(argv):
    if not argv:
        print(__doc__)
        return
    cmd = argv[0]
    if cmd == "extract":
        if len(argv) < 2:
            print("用法: python docx_qa.py extract <file.docx>")
            return
        print(extract_all(argv[1]))
        return

    if cmd == "fix":
        in_path = out_path = None
        rules = []
        i = 1
        while i < len(argv):
            a = argv[i]
            if a == "--in":
                in_path = argv[i + 1]; i += 2
            elif a == "--out":
                out_path = argv[i + 1]; i += 2
            elif a == "--rules":
                with open(argv[i + 1], "r", encoding="utf-8") as f:
                    rules = json.load(f)
                i += 2
            elif a == "--pair":
                old, new = argv[i + 1].split("||", 1)
                rules.append([old, new]); i += 2
            else:
                i += 1
        if not in_path or not out_path:
            print("用法: python docx_qa.py fix --in <in.docx> --out <out.docx> "
                  "--rules rules.json | --pair \"旧||新\"")
            return
        fix_docx(in_path, out_path, rules)
        return

    print("未知子命令:", cmd)
    print(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
