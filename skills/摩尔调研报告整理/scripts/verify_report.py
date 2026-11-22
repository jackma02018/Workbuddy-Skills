#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""摩尔调研报告整理 —— 交付验收脚本（零第三方依赖）

用法:
    python verify_report.py <报告.docx> [--fingerprint]

检查 10 类：结构 / 三段式 / 表格 / 措辞红线 / 元说明 / 编号与附录 /
页脚 / 术语 / 包完整性 / 文件规范。
退出码 0 = 无 FAIL（允许 WARN），1 = 存在 FAIL。
"""
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

W = r"<w:t(?:\s[^>]*)?>(.*?)</w:t>"          # 严格匹配 w:t，避免误吃 w:tcPr / w:tcW
BLOCK = re.compile(r"(<w:tbl>.*?</w:tbl>|<w:p[ >].*?</w:p>)", re.S)
PARA = re.compile(r"<w:p[ >].*?</w:p>", re.S)
ROW = re.compile(r"<w:tr[ >].*?</w:tr>", re.S)
CELL = re.compile(r"<w:tc>.*?</w:tc>", re.S)

CHAPTERS = ["一、项目基本信息", "二、项目背景", "三、项目调研",
            "四、实施范围与边界建议", "五、后续行动计划"]
SUB2 = ["2.1", "2.2", "2.3"]
SUB3 = ["3.%d" % i for i in range(1, 12)]
SUB4 = ["4.1", "4.2"]

HEADERS = [
    ["项目字段", "内容"],
    ["车间数量", "线体/区域名称", "线体数量", "位置", "设备组成"],
    ["产品类别", "代表产品", "应用领域"],
    ["系统 / 设备名称", "供应商", "主要功能", "应用现状"],
    ["仓库名称", "数量", "面积（㎡）", "位置", "存放物料"],
    ["车间/产线", "产线数量", "位置", "设备组成"],
    ["待办事项", "责任人", "预计完成时间"],
    ["术语", "中文全称", "本报告用法"],
]

# 交付文档中 0 命中的禁用词（宽泛词只用资料来源类，业务词不误伤）
BANNED = [
    "[E-", "证据索引", "附：证据索引",
    "录音", "转写", "听记", "纪要", "源文件", "时间码", "发言人",
    "资料基础", "口述", "复述", "钉钉 AI", "两段录音", "4 份源文件",
    "按客户要求", "按要求", "本次修订", "上一版", "根据意见",
    "供应商方案与行业案例演示", "依据：", "证据编号均连续无跳号",
    "XXXXXX", "XX有限公司", "TBD", "待填写",
]

# 本次逐条删过的元说明精确字符串
META_SENTENCES = [
    "参会人员姓名本次资料未载明，本报告不推断具体人名。",
    "按项目资料夹命名；客户工商全称本次资料未载明",
    "园区政府提供整层空间，本期范围为其半层；具体门牌地址本次资料未载明",
    "摩尔元数侧项目负责人本次资料未载明。",
    "本报告不据此推算产能。",
    "同一事实只在一处陈述。",
    "15:38—16:54", "16:54—17:03",
]

TERMS = ["MES", "MOM", "WMS", "QMS", "IOT", "ECN", "SOP", "Andon", "SPC", "OEE"]
FULL_NAMES = ["制造执行系统", "制造运营管理", "仓储管理系统", "质量管理系统",
              "物联网", "工程变更通知", "标准作业指导书", "统计过程控制",
              "设备综合效率"]


def text_of(xml_fragment):
    return "".join(re.findall(W, xml_fragment, re.S))


def load(path):
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf-8")
    body = doc[doc.find("<w:body>"):]
    top_paras, tables, seq = [], [], []
    for m in BLOCK.finditer(body):
        x = m.group(0)
        if x.startswith("<w:tbl"):
            rows = []
            for r in ROW.findall(x):
                rows.append([text_of(c).strip() for c in CELL.findall(r)])
            tables.append(rows)
            seq.append(("T", len(tables) - 1))
        else:
            t = text_of(x).strip()
            if t:
                top_paras.append(t)
                seq.append(("P", t))
    all_paras = [text_of(x) for x in PARA.findall(doc)]
    footers = []
    for n in z.namelist():
        if "footer" in n and n.endswith(".xml"):
            footers.append((n, z.read(n).decode("utf-8")))
    return z, doc, top_paras, tables, all_paras, footers


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, name, detail):
        self.items.append((level, name, detail))

    def check(self, ok, name, ok_msg, fail_msg):
        self.add("PASS" if ok else "FAIL", name, ok_msg if ok else fail_msg)

    def warn(self, name, detail):
        self.add("WARN", name, detail)

    def summary(self):
        n = {"PASS": 0, "WARN": 0, "FAIL": 0}
        for lv, name, detail in self.items:
            n[lv] += 1
            print("[%-4s] %s —— %s" % (lv, name, detail))
        print("\n== 汇总: PASS=%d  WARN=%d  FAIL=%d ==" % (n["PASS"], n["WARN"], n["FAIL"]))
        return 1 if n["FAIL"] else 0


def section_span(paras, start_prefix, ends):
    """返回从 start_prefix 段落之后到下一个同级标题之前的段落列表"""
    try:
        i = next(k for k, p in enumerate(paras) if p.startswith(start_prefix))
    except StopIteration:
        return None
    j = len(paras)
    for k in range(i + 1, len(paras)):
        if any(paras[k].startswith(e) for e in ends):
            j = k
            break
    return paras[i + 1:j]


def run(path, fingerprint=False):
    r = Report()
    if not os.path.isfile(path):
        print("[FAIL] 文件不存在: %s" % path)
        return 1
    try:
        z, doc, paras, tables, all_paras, footers = load(path)
    except Exception as e:
        print("[FAIL] 无法读取 docx: %s" % e)
        return 1

    full = "".join(all_paras) + "".join(text_of(f[1]) for f in footers)

    # 09 包完整性
    bad = []
    for n in z.namelist():
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                bad.append("%s(%s)" % (n, str(e)[:40]))
    r.check(not bad, "09 包完整性", "全部 XML/rels 可解析，共 %d 个部件" % len(z.namelist()),
            "XML 解析失败: %s" % "; ".join(bad[:3]))

    # 01 结构
    missing = [c for c in CHAPTERS if not any(p.startswith(c) for p in paras)]
    r.check(not missing, "01a 五章齐全", "一至五章齐全且顺序固定",
            "缺章: %s" % ", ".join(missing))
    idx = [next((k for k, p in enumerate(paras) if p.startswith(c)), -1) for c in CHAPTERS]
    r.check(idx == sorted(idx) and -1 not in idx, "01b 章序连续", "章号顺序正确", "章号顺序异常")

    for label, subs in [("2.x", SUB2), ("3.x", SUB3), ("4.x", SUB4)]:
        miss = [s for s in subs if not any(p.startswith(s + " ") for p in paras)]
        r.check(not miss, "01c %s 子节连续" % label,
                "%s 共 %d 个子节，无跳号" % (label, len(subs)),
                "缺失: %s" % ", ".join(miss))

    abs_span = section_span(paras, "调研结论摘要", [CHAPTERS[0]])
    n_abs = 0 if abs_span is None else len([p for p in abs_span if re.match(r"^[1-9]\. ", p)])
    r.check(n_abs == 5, "01d 结论摘要 5 条", "调研结论摘要 5 条", "实际 %d 条" % n_abs)

    concl = 0
    for c in CHAPTERS:
        sp = section_span(paras, c, CHAPTERS + ["附：术语约定", "文末自检清单"])
        if sp and sp[0].startswith("本章结论："):
            concl += 1
    r.check(concl == 5, "01e 每章结论句", "5 个章首均有「本章结论：」", "仅 %d 章有章首结论" % concl)

    # 02 三段式
    ends = SUB3[1:] + ["四、实施范围与边界建议"]
    trio_ok = 0
    for s in SUB3:
        sp = section_span(paras, s + " ", ends)
        if sp and all(any(p.startswith(k) for p in sp) for k in ("现状：", "需求：", "痛点：")):
            trio_ok += 1
    r.check(trio_ok == 11, "02 三段式", "11 个子节均含 现状/需求/痛点 段首结论",
            "仅 %d/11 个子节完整" % trio_ok)
    for k in ("现状：", "需求：", "痛点："):
        r.check(full.count(k) >= 11, "02 %s 段数" % k[:-1], "全文 %d 处" % full.count(k),
                "%s 仅 %d 处" % (k, full.count(k)))

    # 03 表格
    r.check(len(tables) == 8, "03a 表格总数", "8 张表（7 主表 + 1 术语表）",
            "实际 %d 张" % len(tables))
    if len(tables) >= 8:
        for i, expect in enumerate(HEADERS):
            head = " | ".join(tables[i][0]) if tables[i] else ""
            miss = [h for h in expect if h not in head]
            r.check(not miss, "03b 表%d 表头" % (i + 1 if i < 7 else 8),
                    "字段沿用范本", "缺字段: %s" % ", ".join(miss))
        t7 = tables[6]
        empty = [k for k, row in enumerate(t7[1:], 2)
                 if len(row) < 3 or not row[1].strip() or not row[2].strip()]
        r.check(not empty, "03c 表7 责任人与时间", "%d 条待办两列均非空" % (len(t7) - 1),
                "第 %s 行责任人或时间为空" % empty)

    # 04 措辞红线
    hits = {k: full.count(k) for k in BANNED if full.count(k)}
    bare = len(re.findall(r"E-\d{3}", full))
    if bare:
        hits["E-nnn 裸编号"] = bare
    r.check(not hits, "04a 禁用词 0 命中", "资料来源字眼/证据标注/占位符均为 0",
            "命中: %s" % ", ".join("%s×%d" % (k, v) for k, v in list(hits.items())[:8]))

    meta_hits = [s for s in META_SENTENCES if s in full]
    r.check(not meta_hits, "04b 元说明 0 命中", "7 类已删元说明句均未复现",
            "复现: %s" % " / ".join(m[:24] for m in meta_hits))

    # 05 编号与附录
    r.check("附：术语约定" in full, "05a 术语约定附录", "附：术语约定（不编章号）存在", "缺术语约定附录")
    self_chk = len(re.findall(r"（[1-7]）", full))
    r.check(self_chk >= 7, "05b 自检清单 7 条", "文末自检清单 7 条齐全",
            "仅 %d 条" % self_chk)
    r.check("文末自检清单" in full, "05c 自检清单标题", "文末自检清单存在", "缺文末自检清单")
    if "其它管理（总）" in full:
        r.warn("05d 3.9 称谓消歧", "检出「其它管理（总）」——范本遗留称谓，应统一为「车间分组细化」")
    else:
        r.add("PASS", "05d 3.9 称谓消歧", "3.9 标题与引用处一致")

    # 06 页脚
    ftext = "".join(text_of(f[1]) for f in footers)
    fxml = "".join(f[1] for f in footers)
    need = ["编制：", "编制日期：", "版本：", "第"]
    miss = [k for k in need if k not in ftext]
    r.check(footers and not miss, "06a 页脚四要素", "编制/编制日期/版本/页码齐全",
            "页脚缺失或字段不全: %s" % (", ".join(miss) if miss else "无 footer 部件"))
    r.check("PAGE" in fxml, "06b 页码字段", "PAGE 域存在", "页码为硬编码或缺失")
    r.check("依据" not in ftext, "06c 页脚无资料来源", "页脚无「依据」字段", "页脚仍含「依据」")

    # 07 术语
    term_tbl = " ".join(" ".join(row) for row in tables[7]) if len(tables) >= 8 else ""
    miss_t = [t for t in TERMS if t not in term_tbl]
    r.check(not miss_t, "07a 术语表覆盖", "MES/MOM/WMS/QMS/IOT/ECN/SOP/Andon/SPC/OEE 齐备",
            "缺: %s" % ", ".join(miss_t))
    miss_f = [t for t in FULL_NAMES if t not in full]
    r.check(not miss_f, "07b 中文全称", "关键术语中文全称齐备", "缺: %s" % ", ".join(miss_f))
    # 术语约定表「资料中出现的 ECN/ECM/DCN 等不同写法统一为 ECN」属于规范说明，不算正文异写
    prose = full.replace("ECN/ECM/DCN", "").replace("ECM/DCN", "")
    r.check(not re.search(r"\b(ECM|DCN)\b", prose), "07c 术语异写", "ECN 写法统一（术语表说明除外）",
            "正文检出 ECN 异写（ECM/DCN）")

    # 08 文件规范
    fn = os.path.basename(path)
    ok_fn = bool(re.search(r"MES 系统建设现场调研报告_\d{8}(_修订版)?\.docx$", fn))
    r.check(ok_fn, "08a 文件命名", "符合 <客户简称> MES 系统建设现场调研报告_YYYYMMDD.docx",
            "命名不规范: %s" % fn)
    d = os.path.dirname(path) or "."
    tmp = [f for f in os.listdir(d) if re.match(r"^_.*\.(py|txt)$", f) or f == "__pycache__"]
    if tmp:
        r.warn("08b 目录整洁", "检出中间文件: %s" % ", ".join(tmp[:5]))
    else:
        r.add("PASS", "08b 目录整洁", "交付目录无中间文件")

    if fingerprint:
        print("\n-- 指纹 --")
        print("size=%d bytes | parts=%d | tables=%d | top_paras=%d | cells=%d | chars=%d" %
              (os.path.getsize(path), len(z.namelist()), len(tables), len(paras),
               sum(len(row) for t in tables for row in t), len(full)))
        print("footer=%s" % repr(ftext))
        print("baseline_sha1=%s" % __import__("hashlib").sha1(open(path, "rb").read()).hexdigest()[:16])
    return r.summary()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    sys.exit(run(args[0], fingerprint="--fingerprint" in sys.argv))
