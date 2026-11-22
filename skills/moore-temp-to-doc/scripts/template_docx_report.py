#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
template_docx_report.py — 零依赖「模板保格式」Word 报告生成器

核心能力：以一份已有 .docx 报告（如《MES客户调研报告模板_20260820.docx》）为模板，
复用其 styles.xml / theme / fontTable，仅替换 word/document.xml 正文，
生成格式与模板 100% 一致的新报告。**不需要 python-docx**，仅用标准库。

用法：
  # 1) 抽取任何 docx 的全文（段落 + 表格），用于读模板骨架 / 质检
  python template_docx_report.py extract <file.docx>

  # 2) 导出模板骨架（章节标题 + 样式ID + 表格字段名），用于确定报告章节
  python template_docx_report.py outline <template.docx>

  # 3) 从内容 JSON 生成报告
  python template_docx_report.py build --template <template.docx> --content <content.json> --out <out.docx>

  # 4) 校验生成的 docx（ZIP 完整性 + XML 良构 + 结构统计）
  python template_docx_report.py verify <out.docx>

内容 JSON 格式见同目录 examples/content_example.json，或读 references/content_schema.md。
"""
import sys
import os
import re
import json
import zipfile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

FONT_XML = ('<w:rFonts w:hint="eastAsia" w:ascii="{f}" w:hAnsi="{f}" '
            'w:eastAsia="{f}" w:cs="{f}"/>')
DEFAULT_FONT = "微软雅黑"

DOC_NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData"'
)


# ============================ 基础工具 ============================
def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def runs(text, bold=False, color=None, size=None, font=DEFAULT_FONT):
    """文本 → w:r 序列，支持 **粗体** 行内标记。"""
    out = []
    for p in re.split(r"(\*\*.+?\*\*)", str(text)):
        if not p:
            continue
        b = bold
        if p.startswith("**") and p.endswith("**") and len(p) > 4:
            p, b = p[2:-2], True
        rpr = "<w:rPr>" + FONT_XML.format(f=font)
        if b:
            rpr += "<w:b/>"
        if color:
            rpr += '<w:color w:val="%s"/>' % color
        if size:
            rpr += '<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (size, size)
        rpr += "</w:rPr>"
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(p)))
    return "".join(out)


def para(text="", align=None, bold=False, color=None, size=None,
         bullet=False, style=None, spacing=None, indent=None, font=DEFAULT_FONT):
    ppr = "<w:pPr>"
    if style:
        ppr += '<w:pStyle w:val="%s"/>' % style
    if spacing:
        ppr += "<w:spacing %s/>" % spacing
    if indent:
        ppr += "<w:ind %s/>" % indent
    if align:
        ppr += '<w:jc w:val="%s"/>' % align
    ppr += "<w:rPr>" + FONT_XML.format(f=font) + ("<w:b/>" if bold else "")
    if color:
        ppr += '<w:color w:val="%s"/>' % color
    if size:
        ppr += '<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (size, size)
    ppr += "</w:rPr></w:pPr>"
    body = ""
    if text:
        body = runs(("• " + str(text)) if bullet else text,
                    bold=bold, color=color, size=size, font=font)
    return "<w:p>%s%s</w:p>" % (ppr, body)


def make_table(rows, widths, header=True, shade="D9E2F3", font=DEFAULT_FONT):
    grid = "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
    borders = ("<w:tblBorders>"
               + "".join('<w:%s w:val="single" w:color="auto" w:sz="4" w:space="0"/>' % s
                         for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
               + "</w:tblBorders>")
    tblpr = ('<w:tblPr><w:tblStyle w:val="33"/><w:tblW w:w="0" w:type="auto"/>'
             '<w:tblInd w:w="0" w:type="dxa"/>' + borders +
             '<w:tblLayout w:type="fixed"/>'
             '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>'
             '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar>'
             "</w:tblPr><w:tblGrid>" + grid + "</w:tblGrid>")

    trs = ""
    for i, row in enumerate(rows):
        is_h = header and i == 0
        tcs = ""
        for j, c in enumerate(row):
            tcpr = '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>' % widths[j]
            if is_h and shade:
                tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade
            tcpr += '<w:vAlign w:val="center"/></w:tcPr>'
            paras = "".join(para(line, bold=is_h, size="19", font=font)
                            for line in str(c).split("\n"))
            tcs += "<w:tc>%s%s</w:tc>" % (tcpr, paras)
        trs += ('<w:tr><w:trPr>' + ("<w:tblHeader/>" if is_h else "")
                + "</w:trPr>" + tcs + "</w:tr>")
    return "<w:tbl>" + tblpr + trs + "</w:tbl>"


# ============================ 抽取 / 骨架 ============================
def docx_text(path):
    """提取 docx 全文（段落 + 表格单元格），保留换行。"""
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


def docx_outline(path):
    """导出模板骨架：章节标题（含样式ID）+ 表格字段名。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")

    lines = []
    lines.append("# 模板骨架: %s" % os.path.basename(path))
    lines.append("")
    lines.append("## 章节标题")
    for m in re.finditer(r"<w:p [^>]*>(.*?)</w:p>", xml, re.S):
        body = m.group(1)
        st = re.search(r'<w:pStyle w:val="([^"]+)"', body)
        if st:
            txt = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", body, re.S))
            if txt.strip():
                lines.append("  [style=%s] %s" % (st.group(1), txt.strip()))

    lines.append("")
    lines.append("## 表格字段名")
    # 整体解析文档，避免片段缺少命名空间声明导致解析失败
    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring(xml)
    except Exception:
        root = None
    if root is not None:
        for i, tbl in enumerate(root.iter(W + "tbl"), 1):
            rows = tbl.findall(W + "tr")
            if not rows:
                continue
            hdr = []
            for tc in rows[0].findall(W + "tc"):
                hdr.append("".join(t.text or "" for t in tc.iter(W + "t")).strip())
            lines.append("  表%d (%d行 x %d列): %s"
                         % (i, len(rows), len(hdr), " | ".join(hdr)))

    lines.append("")
    lines.append("## 页面设置")
    m = re.search(r"<w:sectPr.*?</w:sectPr>", xml, re.S)
    lines.append("  " + (m.group(0) if m else "(未找到)"))
    return "\n".join(lines)


# ============================ 生成 ============================
def content_to_body(content):
    """内容 JSON → document.xml 的 body XML。"""
    blocks = []
    for blk in content.get("blocks", []):
        t = blk.get("type")
        if t == "title":
            blocks.append(para(blk["text"], align="center", bold=True,
                               size=blk.get("size", "28")))
        elif t == "subtitle":
            blocks.append(para(blk["text"], align="center",
                               color=blk.get("color", "666666"),
                               size=blk.get("size", "18")))
        elif t == "h1":
            blocks.append(para(blk["text"], style=blk.get("style", "3"), bold=True,
                               size=blk.get("size", "30")))
        elif t == "h2":
            blocks.append(para(blk["text"], style=blk.get("style", "4"), bold=True))
        elif t == "h3":
            blocks.append(para(blk["text"], style=blk.get("style", "5"), bold=True))
        elif t == "p":
            blocks.append(para(blk.get("text", ""), align=blk.get("align"),
                               bold=blk.get("bold", False), color=blk.get("color"),
                               size=blk.get("size"), bullet=blk.get("bullet", False)))
        elif t == "table":
            blocks.append(make_table(blk["rows"], blk["widths"],
                                     header=blk.get("header", True),
                                     shade=blk.get("shade", "D9E2F3")))
        elif t == "blank":
            blocks.append(para(""))
    return "".join(blocks)


def build(template, content_path, out_path):
    with open(content_path, "r", encoding="utf-8") as f:
        content = json.load(f)

    body = content_to_body(content)
    sect = content.get(
        "sectPr",
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720" w:num="1"/>'
        '<w:docGrid w:linePitch="360" w:charSpace="0"/></w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<w:document ' + DOC_NS + ' mc:Ignorable="w14 w15 wp14"><w:body>'
                + body + sect + "</w:body></w:document>")

    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(template) as src:
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.namelist():
                if item == "word/document.xml":
                    continue
                dst.writestr(item, src.read(item))
            dst.writestr("word/document.xml", document.encode("utf-8"))
    print("BUILT -> %s  (%d bytes)" % (out_path, os.path.getsize(out_path)))
    return out_path


# ============================ 校验 ============================
def verify(path):
    from xml.dom.minidom import parseString

    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        print("ZIP 完整性:", "OK" if bad is None else "损坏: %s" % bad)
        ok = True
        for n in z.namelist():
            if n.endswith((".xml", ".rels")):
                try:
                    parseString(z.read(n))
                except Exception as e:
                    ok = False
                    print("  XML 错误 %s -> %s" % (n, e))
        print("XML 良构:", "全部通过" if ok else "存在错误")
        xml = z.read("word/document.xml").decode("utf-8")

    print("段落数:", xml.count("<w:p>"))
    print("表格数:", xml.count("<w:tbl>"))
    print("行数  :", xml.count("<w:tr>"))
    styles = re.findall(r'<w:pStyle w:val="([^"]+)"/>', xml)
    print("标题样式分布:", {s: styles.count(s) for s in sorted(set(styles))})
    return ok and bad is None


# ============================ CLI ============================
def main(argv):
    if not argv:
        print(__doc__)
        return
    cmd = argv[0]
    if cmd == "extract" and len(argv) >= 2:
        print(docx_text(argv[1]))
    elif cmd == "outline" and len(argv) >= 2:
        print(docx_outline(argv[1]))
    elif cmd == "build":
        kw = {}
        i = 1
        while i < len(argv) - 1:
            if argv[i] in ("--template", "--content", "--out"):
                kw[argv[i][2:]] = argv[i + 1]
                i += 2
            else:
                i += 1
        if not all(k in kw for k in ("template", "content", "out")):
            print("用法: build --template <t.docx> --content <c.json> --out <o.docx>")
            return
        build(kw["template"], kw["content"], kw["out"])
    elif cmd == "verify" and len(argv) >= 2:
        verify(argv[1])
    else:
        print(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
