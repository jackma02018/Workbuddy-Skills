# -*- coding: utf-8 -*-
"""
电子发票预处理脚本：抽取票面信息 → 重命名 PDF → 生成配套 JPG
用于 ERP 系统上传前的格式准备。

用法：
  python sf_invoice_process.py <pdf_file1> [pdf_file2 ...]

产物：
  - 同目录下输出重命名后的 PDF（¥金额+销方名称+发票号码+物流单号.pdf）
  - 同目录下输出同名 JPG 图片（200 DPI）
"""

import fitz  # pymupdf
import pdfplumber
import re
import os
import sys


def parse_invoice_fields(pdf_path):
    """从电子发票 PDF 中抽取关键字段。返回 dict 或 None。"""
    try:
        with pdfplumber.open(pdf_path) as pdf_doc:
            raw_text = pdf_doc.pages[0].extract_text()
    except Exception as err:
        print(f"  ERROR: 读取 PDF 文本失败: {err}")
        return None

    if not raw_text:
        print("  ERROR: PDF 中未检测到文本内容")
        return None

    # 发票类型校验 — 仅处理增值税专用发票
    if "增值税专用发票" not in raw_text:
        print("  ERROR: 非增值税专用发票，跳过该文件")
        return None

    fields = {}

    # 发票号码
    match_no = re.search(r"发票号码[：:](\d+)", raw_text)
    if match_no:
        fields["invoice_no"] = match_no.group(1)
    else:
        print("  ERROR: 未找到发票号码")
        return None

    # 销方名称 — 取最后一个「名称：XXX有限公司」，销方在购方之后
    seller_matches = re.findall(r"名称[：:](.*?有限公司)", raw_text)
    if seller_matches:
        fields["seller"] = seller_matches[-1].strip()
    else:
        print("  ERROR: 未找到销方名称")
        return None

    # 价税合计（小写）
    match_amt = re.search(r"（小写）([¥￥])([\d.]+)", raw_text)
    if match_amt:
        fields["total_amount"] = match_amt.group(2)
    else:
        print("  ERROR: 未找到价税合计金额")
        return None

    # 物流单号（备注区）— 必填项，SF 编号可能在备注前后
    match_rmk = re.search(r"(SF\d+)\s*备\s*注|备\s*注\s*(SF\d+)", raw_text, re.DOTALL)
    if match_rmk:
        fields["tracking_no"] = match_rmk.group(1) or match_rmk.group(2)
    else:
        print("  ERROR: 未找到物流单号")
        return None

    fields["raw_text"] = raw_text
    return fields


def handle_single_pdf(pdf_path):
    """处理单个发票 PDF：重命名并生成 JPG。返回新 PDF 路径。"""
    print(f"\n处理文件: {os.path.basename(pdf_path)}")

    fields = parse_invoice_fields(pdf_path)
    if not fields:
        print("  SKIP: 字段抽取失败")
        return None

    print(f"  发票号码: {fields['invoice_no']}")
    print(f"  销方名称: {fields['seller']}")
    print(f"  价税合计: ¥{fields['total_amount']}")
    print(f"  物流单号: {fields['tracking_no']}")

    # 组装新文件名（物流单号为必填项，始终包含）
    output_stem = f"¥{fields['total_amount']}+{fields['seller']}+{fields['invoice_no']}+{fields['tracking_no']}"

    parent_dir = os.path.dirname(os.path.abspath(pdf_path))
    output_pdf = os.path.join(parent_dir, output_stem + ".pdf")
    output_jpg = os.path.join(parent_dir, output_stem + ".jpg")

    # 重命名 PDF
    if os.path.abspath(pdf_path) != os.path.abspath(output_pdf):
        if os.path.exists(output_pdf):
            print(f"  WARN: 目标 PDF 已存在，跳过重命名: {output_stem}.pdf")
            output_pdf = pdf_path
        else:
            os.rename(pdf_path, output_pdf)
            print(f"  已重命名: {os.path.basename(output_pdf)}")

    # 生成 JPG 图片
    if os.path.exists(output_jpg):
        print(f"  WARN: JPG 已存在，跳过生成: {output_stem}.jpg")
    else:
        try:
            pdf_doc = fitz.open(output_pdf)
            first_page = pdf_doc[0]
            scale_matrix = fitz.Matrix(200 / 72, 200 / 72)
            pixmap = first_page.get_pixmap(matrix=scale_matrix)
            pixmap.save(output_jpg)
            pdf_doc.close()
            print(f"  已生成 JPG: {os.path.basename(output_jpg)} ({pixmap.width}x{pixmap.height})")
        except Exception as err:
            print(f"  ERROR: JPG 生成失败: {err}")

    return output_pdf


def main():
    if len(sys.argv) < 2:
        print("用法: python sf_invoice_process.py <pdf_file1> [pdf_file2 ...]")
        sys.exit(1)

    file_list = sys.argv[1:]
    ok_count = 0
    fail_count = 0

    for file_path in file_list:
        if not os.path.exists(file_path):
            print(f"\nSKIP: 文件不存在: {file_path}")
            fail_count += 1
            continue
        if not file_path.lower().endswith(".pdf"):
            print(f"\nSKIP: 非 PDF 文件: {file_path}")
            fail_count += 1
            continue

        result_path = handle_single_pdf(file_path)
        if result_path:
            ok_count += 1
        else:
            fail_count += 1

    print(f"\n{'='*50}")
    print(f"处理完成: 成功 {ok_count}, 失败 {fail_count}")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
