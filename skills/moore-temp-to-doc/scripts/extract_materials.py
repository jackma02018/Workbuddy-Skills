#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_materials.py — 多源临时素材摘要器（零依赖）

把四类临时素材（办公文档 / 富文本 / 媒体转写 / 通讯文本）转成结构化 markdown 摘要，
供后续装配 content.json 用。**零第三方依赖**，仅用 Python 标准库。

支持的素材类型：

| 类型 | 提取方式 | 备注 |
|---|---|---|
| `.docx` | 复用 template_docx_report.py 的 docx_text 逻辑 | 调用 extract 子命令的等价实现 |
| `.xlsx` | 解析 xl/sharedStrings.xml + xl/worksheets/sheetN.xml | 含每张 sheet 的行列内容 |
| `.pptx` | 解析 ppt/slides/slide*.xml 抽每页标题+正文 | 适合提纲式汇报材料 |
| `.pdf`  | 简单文本流提取（解压/读取），扫描版提示用户 | 复杂版用 pdftotext 更佳 |
| `.html` | re.sub 去标签，保留段落结构 | 适合客户资料/竞品官网存档 |
| `.md`/`.txt` | 直接读 | 录音转写 / 笔记 / 邮件正文 |
| `.eml` | 解析邮件头 + 正文（处理 multipart） | 适合 Outlook/QQ 邮箱导出 |

用法：
  # 自动识别（按扩展名）
  python extract_materials.py meeting.docx
  python extract_materials.py notes.md
  python extract_materials.py 聊天记录.txt

  # 批量摘要整个目录（递归）
  python extract_materials.py <目录> --all

  # 指定输出文件
  python extract_materials.py xxx.pdf --out <输出.md>

  # 指定类型（当扩展名不规范时）
  python extract_materials.py xxx --type pdf

输出：markdown 到 stdout 或 --out 指定文件。
"""
import sys
import os
import re
import zipfile
import html as html_lib
import json
import argparse
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS


# ============================ 各类型提取器 ============================
def extract_docx(path):
    """docx → 段落+表格纯文本（复用 template_docx_report 的 docx_text 逻辑）。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    xml = xml.replace("</w:p>", "\n").replace("</w:tr>", "\n").replace("</w:tc>", "\t")
    xml = re.sub(r"<w:tab[^>]*/>", "\t", xml)
    xml = re.sub(r"<w:br[^>]*/>", "\n", xml)
    txt = re.sub(r"<[^>]+>", "", xml)
    for a, b in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                 ("&quot;", '"'), ("&apos;", "'")):
        txt = txt.replace(a, b)
    out = []
    for line in txt.split("\n"):
        s = line.strip()
        if s and s != "\t":
            out.append(re.sub(r"(\s*\t\s*)+", " | ", s))
    return "\n".join(out)


def extract_xlsx(path):
    """xlsx → 每 sheet 的二维表（读 sharedStrings 拼接单元格引用）。"""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        # 1) 读共享字符串表
        sst = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.iter():
                if si.tag.endswith("}si"):
                    txt = "".join(t.text or "" for t in si.iter() if t.tag.endswith("}t"))
                    sst.append(txt)
        # 2) 找所有 sheet，按 sheet1.xml, sheet2.xml ... 顺序
        sheets = sorted(
            [n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)],
            key=lambda n: int(re.search(r"sheet(\d+)\.xml", n).group(1))
        )
        out = []
        for s in sheets:
            root = ET.fromstring(z.read(s))
            sheet_data = root.find("{%s}sheetData" % W_NS)
            out.append(f"## Sheet: {os.path.basename(s)}")
            if sheet_data is None:
                out.append("  (空)")
                continue
            for row in sheet_data.findall("{%s}row" % W_NS):
                cells = []
                for c in row.findall("{%s}c" % W_NS):
                    t = c.get("t", "n")
                    v = c.find("{%s}v" % W_NS)
                    if v is None or v.text is None:
                        cells.append("")
                        continue
                    if t == "s":
                        idx = int(v.text)
                        cells.append(sst[idx] if idx < len(sst) else "")
                    elif t == "inlineStr":
                        is_node = c.find("{%s}is" % W_NS)
                        cells.append("".join(t2.text or "" for t2 in is_node.iter() if t2.tag.endswith("}t")) if is_node is not None else "")
                    else:
                        cells.append(v.text)
                out.append("  " + " | ".join(cells))
        return "\n".join(out)


def extract_pptx(path):
    """pptx → 每页标题+正文纲要。"""
    with zipfile.ZipFile(path) as z:
        names = sorted(
            [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=lambda n: int(re.search(r"slide(\d+)\.xml", n).group(1))
        )
        out = []
        for n in names:
            num = re.search(r"slide(\d+)\.xml", n).group(1)
            out.append(f"## Slide {num}")
            try:
                root = ET.fromstring(z.read(n))
            except ET.ParseError as e:
                out.append(f"  (解析失败: {e})")
                continue
            texts = []
            for sp in root.iter():
                if sp.tag.endswith("}sp"):
                    lines = []
                    for t in sp.iter():
                        if t.tag.endswith("}t") and t.text:
                            lines.append(t.text)
                    if lines:
                        texts.append(" ".join(lines))
            for t in texts:
                out.append(f"  - {t}")
            if not texts:
                out.append("  (空)")
        return "\n".join(out)


def extract_pdf(path):
    """
    PDF → 尝试提取文本流（限于简单文本型 PDF）。
    扫描版 PDF 提示用 pdftotext 或 Read 工具。
    """
    # 1) 试着用 pdfplumber/pypdf 之类——没有，零依赖
    # 2) 退化：作为二进制读，提取 BT...ET 文本操作符之间的字符串
    # 这对简单文本 PDF 能拿到 80%+，扫描版基本拿不到
    try:
        data = open(path, "rb").read()
    except Exception as e:
        return f"(读取失败: {e})"
    # 解压 zlib 流（很多 PDF 文本用 FlateDecode 压缩）
    import zlib
    chunks = []
    for m in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", data, re.S):
        try:
            decompressed = zlib.decompress(m.group(1))
            chunks.append(decompressed)
        except Exception:
            chunks.append(m.group(1))
    text = b"\n".join(chunks).decode("utf-8", "ignore")
    # 抽取 (...) 形式的字符串（PDF 文本对象常用）
    lines = re.findall(r"\((.*?)\)\s*T[jJ]", text)
    if not lines:
        return ("(PDF 未提取到文本流——可能是扫描版或加密 PDF。\n"
                "建议：① 安装 poppler 后用 `pdftotext -layout xxx.pdf -`；\n"
                "      ② 用 Read 工具多模态识别整页；\n"
                "      ③ 若 PDF 由 Word 导出，可让客户重发 .docx 源文件。)")
    return "\n".join(lines)


def extract_html(path):
    """HTML → 文本（保留段落结构）。"""
    src = open(path, encoding="utf-8", errors="ignore").read()
    # 去 script/style
    src = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", src, flags=re.S | re.I)
    # 把块级标签换成换行
    src = re.sub(r"</?(p|br|div|li|h[1-6]|tr|td|th|section|article)[^>]*>", "\n",
                 src, flags=re.I)
    # 去所有剩余标签
    txt = re.sub(r"<[^>]+>", "", src)
    # HTML 实体还原
    txt = html_lib.unescape(txt)
    # 合并多余空行
    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip()


def extract_md_txt(path):
    """md / txt 直接读。"""
    return open(path, encoding="utf-8", errors="ignore").read().strip()


def extract_eml(path):
    """eml → 邮件头 + 正文（处理 multipart）。"""
    import email
    from email.policy import default
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=default)
    headers = []
    for k in ("From", "To", "Cc", "Subject", "Date"):
        v = msg.get(k)
        if v:
            headers.append(f"{k}: {v}")
    body_parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                body_parts.append(part.get_content())
            elif ct == "text/html":
                body_parts.append("[HTML 正文]\n" + extract_html_from_str(part.get_content()))
    else:
        body_parts.append(msg.get_content())
    return "\n".join(headers) + "\n\n" + "\n\n---\n\n".join(p for p in body_parts if p)


def extract_html_from_str(s):
    """复用 html 提取逻辑作用于字符串。"""
    src = s
    src = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", src, flags=re.S | re.I)
    src = re.sub(r"</?(p|br|div|li|h[1-6]|tr|td|th)[^>]*>", "\n", src, flags=re.I)
    txt = re.sub(r"<[^>]+>", "", src)
    return re.sub(r"\n{3,}", "\n\n", html_lib.unescape(txt)).strip()


# ============================ 类型路由 ============================
DISPATCH = {
    ".docx": extract_docx,
    ".xlsx": extract_xlsx,
    ".pptx": extract_pptx,
    ".pdf": extract_pdf,
    ".html": extract_html,
    ".htm": extract_html,
    ".md": extract_md_txt,
    ".txt": extract_md_txt,
    ".eml": extract_eml,
}


def extract_one(path, type_override=None):
    """提取单个文件为 markdown 块（含元信息头）。"""
    if not os.path.exists(path):
        return f"# {os.path.basename(path)}\n\n(文件不存在: {path})"
    if os.path.isdir(path):
        return extract_dir(path)
    ext = os.path.splitext(path)[1].lower()
    tp = type_override or ext.lstrip(".")
    fn = DISPATCH.get(ext) if not type_override else DISPATCH.get("." + type_override.lower())
    if fn is None:
        return (f"# {os.path.basename(path)}\n\n"
                f"(不支持的扩展名 `{ext}`。支持: {', '.join(sorted(DISPATCH.keys()))})")
    sz = os.path.getsize(path)
    header = (f"# {os.path.basename(path)}\n\n"
              f"- 路径: `{path}`\n"
              f"- 大小: {sz} bytes\n"
              f"- 类型: {tp}\n\n")
    try:
        body = fn(path)
    except Exception as e:
        body = f"(提取失败: {type(e).__name__}: {e})"
    return header + body


def extract_dir(dir_path):
    """递归摘要整个目录。"""
    out = []
    for root, dirs, files in os.walk(dir_path):
        # 跳过常见干扰目录
        dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ("__pycache__",)]
        for f in sorted(files):
            full = os.path.join(root, f)
            out.append(extract_one(full))
            out.append("\n\n---\n\n")
    return "".join(out)


# ============================ CLI ============================
def main(argv):
    p = argparse.ArgumentParser(
        description="多源临时素材摘要器（零依赖）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("path", help="素材路径（文件或目录）")
    p.add_argument("--type", help="强制指定类型（覆盖扩展名）")
    p.add_argument("--out", "-o", help="输出文件（默认 stdout）")
    p.add_argument("--all", action="store_true", help="目录递归摘要")
    args = p.parse_args(argv)

    target = args.path
    if os.path.isdir(target) and not args.all:
        # 单目录默认也走递归摘要
        target_path = target
    else:
        target_path = target

    result = extract_one(target_path, type_override=args.type)

    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(result)
        print("OK -> %s  (%d chars, %d bytes)" % (
            args.out, len(result), os.path.getsize(args.out)))
    else:
        print(result)


if __name__ == "__main__":
    main(sys.argv[1:])