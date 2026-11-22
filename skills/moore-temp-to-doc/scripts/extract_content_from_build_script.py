#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_content_from_build_script.py — 把「项目内硬编码的 build_report.py」转成
通用 content.json，供 template_docx_report.py build 使用。

背景：首份报告（江阴旺达）是用一个把客户名/路径/内容全部硬编码的 build_report.py 生成的。
换客户时与其改那个大脚本，不如把它的 h1/h2/h3/para/table 调用提取成 content.json
—— 内容与引擎分离，后续只需编辑 JSON。

用法：
  python extract_content_from_build_script.py <build_report.py> [--out <content.json>]

识别范围：调用形态为 h1("...") / h2("...") / h3("...") / para("...", 关键字参数) /
table([[...]], [宽度])。若原脚本结构差异较大，本工具会打印 WARN，
此时请人工核对生成的 json 或直接手写 json。
"""
import sys
import os
import re
import json
import ast

TARGETS = ("h1", "h2", "h3", "para", "table")


def scan_calls(src):
    """一次遍历，返回 [(position, fname, args, kwargs)]，保证位置与参数一一对应。
    跳过 def/class 定义行与被注释的调用。"""
    out = []
    pat = re.compile(r"(?<![\w.])(" + "|".join(TARGETS) + r")\s*\(", re.S)
    for m in pat.finditer(src):
        fname = m.group(1)
        # 跳过函数/类定义（def table(...) / class para(...)）
        line_start = src.rfind("\n", 0, m.start()) + 1
        prefix = src[line_start:m.start()]
        if re.match(r"\s*(def|class)\s*$", prefix):
            continue
        # 跳过整行注释
        if prefix.lstrip().startswith("#"):
            continue
        i = m.end() - 1          # 指向 '('
        depth = 0
        in_s = None
        j = i
        while j < len(src):
            ch = src[j]
            if in_s:
                if ch == "\\":
                    j += 2
                    continue
                if ch == in_s:
                    in_s = None
            elif ch in "\"'":
                in_s = ch
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    break
            j += 1
        body = src[i + 1:j]
        try:
            node = ast.parse("f(" + body + ")", mode="eval").body
        except SyntaxError:
            out.append((m.start(), fname, None, None))
            continue
        out.append((m.start(), fname, node.args, node.keywords))
    out.sort(key=lambda x: x[0])
    return out


def lit(node, env=None):
    """求值字面量；env 提供已解析的模块级常量（如 CW = 8640），支持简单算术。"""
    try:
        return ast.literal_eval(node)
    except Exception:
        pass
    if env is None:
        return None
    # 支持 CW、CW-1500、8640 这类表达式
    try:
        code = compile(ast.Expression(node), "<lit>", "eval")
        return eval(code, {"__builtins__": {}}, dict(env))
    except Exception:
        return None


def collect_consts(tree):
    """收集模块级简单赋值常量：CW = 8640 / NAME = "xxx"。"""
    env = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            t = node.targets[0]
            if isinstance(t, ast.Name):
                try:
                    env[t.id] = ast.literal_eval(node.value)
                except Exception:
                    try:
                        env[t.id] = eval(compile(ast.Expression(node.value),
                                                 "<c>", "eval"),
                                         {"__builtins__": {}}, dict(env))
                    except Exception:
                        pass
    return env


def convert(path):
    src = open(path, encoding="utf-8").read()
    blocks, warns = [], []

    # 收集模块级常量，用于解析 CW、CW-1500 这类列宽表达式
    try:
        env = collect_consts(ast.parse(src))
    except SyntaxError:
        env = {}

    for _, fname, args, kwargs in scan_calls(src):
        if args is None:
            warns.append("跳过无法解析的 %s(...) 调用" % fname)
            continue

        if fname in ("h1", "h2", "h3"):
            txt = lit(args[0], env) if args else None
            if isinstance(txt, str):
                blocks.append({"type": fname, "text": txt})
            continue

        if fname == "para":
            txt = lit(args[0], env) if args else ""
            if not isinstance(txt, str):
                txt = ""
            if not txt.strip():
                blocks.append({"type": "blank"})
                continue
            b = {"type": "p", "text": txt}
            for k in (kwargs or []):
                v = lit(k.value, env)
                if k.arg in ("bullet", "bold", "align", "color", "size") and v is not None:
                    b[k.arg] = v
            blocks.append(b)
            continue

        if fname == "table":
            rows = lit(args[0], env) if args else None
            widths = lit(args[1], env) if len(args) > 1 else None
            if isinstance(rows, list) and isinstance(widths, list):
                widths = [int(w) for w in widths]
                if sum(widths) != 8640:
                    warns.append("表格列宽之和 = %d（应为 8640）：%s" % (sum(widths), widths))
                blocks.append({"type": "table",
                               "rows": [[str(c) for c in r] for r in rows],
                               "widths": widths})
            else:
                warns.append("跳过无法解析的 table(...) 调用"
                             "（rows/widths 非字面量，可能需要人工转换）")
            continue

    return blocks, warns


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return
    src = argv[0]
    out = (argv[argv.index("--out") + 1]
           if "--out" in argv and argv.index("--out") + 1 < len(argv)
           else os.path.splitext(src)[0] + ".content.json")

    blocks, warns = convert(src)
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"blocks": blocks}, f, ensure_ascii=False, indent=2)

    types = {}
    for b in blocks:
        types[b["type"]] = types.get(b["type"], 0) + 1
    print("已提取 %d 个内容块 -> %s" % (len(blocks), out))
    print("类型分布:", types)
    for w in warns:
        print("[WARN]", w)
    if warns:
        print("\n提示：有 WARN 时请人工核对生成的 json 是否完整。")


if __name__ == "__main__":
    main(sys.argv[1:])
