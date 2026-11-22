---
name: mes-customer-research-report
description: "Generate a customer research or survey report (especially MES / WMS / smart-manufacturing pre-sales) that strictly follows a given Word template (.docx) and is filled from a folder of source materials (images, screenshots, PPT, Excel, audio). Also covers proofreading the generated report for typos and cross-section inconsistencies, and producing a ~200-character project summary with a fixed set of elements. Trigger phrases: 按模板生成调研报告, 基于模板输出报告, MES客户调研报告, 调研报告质检, 检查错别字, 输出项目总结, or when the user attaches a template .docx plus a materials folder."
agent_created: true
---

# MES 客户调研报告生成与质检

## 目的
把「一个 Word 模板 + 一个素材文件夹」转成结构一致的客户调研报告（docx），并对成品做
错别字 / 上下文一致性质检，最后输出一段约 200 字、含固定要素的项目总结。

## 与 `mes-template-docx-report` 的分工
本技能偏向**素材理解 + 事后质检**；若「格式与模板严格一致」是硬要求、且素材以 docx/Excel 文本为主，
优先用 **`mes-template-docx-report`**（复用模板 styles.xml、零依赖、格式逐像素一致）：

| | **本技能** | mes-template-docx-report |
|---|---|---|
| 侧重 | 多模态素材理解、质检、项目总结 | 模板格式保真复制 + 结构化内容装配 |
| 生成 | editor_sdk / python-docx | 复用模板 styles.xml，标准库替换正文 XML |
| 依赖 | 可能需要 python-docx | 零第三方依赖 |
| 适用 | 素材含图片/PPT/音频需深度解析 | 素材以 docx/Excel 为主，格式一致性优先 |

串联用法：`mes-template-docx-report` 产出报告 → 本技能做质检与总结。

## 何时使用
- 用户给出模板 `.docx` 并要求「按模板 / 基于模板输出 / 生成调研报告」。
- 素材在文件夹里（截图、工艺流程图、PPT、Excel 台账、录音等）。
- 用户要求「检查错别字 / 一致性」「输出报告总结」。

## 工作流

### 1. 解析模板骨架
用 Python `zipfile` 读 `word/document.xml`，按出现顺序提取段落（`p`）与表格（`tbl`）：
- 段落文本取自所有 `w:t`；样式名取自 `w:pPr/w:pStyle/@w:val`。
- 表格逐行取各 `w:tc` 单元格文本。
这样得到模板的「章节标题 + 表格字段名」骨架，作为生成报告的章节依据。

### 2. 分析素材（并行读取）
- **图片 / 截图 / 流程图 PNG**：用 Read 工具直接看（多模态）。
- **PPT（.pptx）**：zipfile 读 `ppt/slides/slideN.xml`，取 `a:t` 命名空间文本。
- **Excel（.xlsx）**：先读 `xl/sharedStrings.xml`，再逐 `xl/worksheets/sheetN.xml`
  按 `c/@t=s` 查共享字符串，拼成行。
- **音频（.m4a/.mp3）**：本地若无 STT/转写工具，**不要臆造内容**，在报告里标注
  「录音未转写，内容基于可读图文整理」即可。

### 3. 生成报告（docx）
优先用 `tencent-local-office-edit` 技能的 editor_sdk：
`create_doc` → `doc_insert_markdown`（可指向本地 `.md` 文件）→ `save_file`。
若 SDK 不可用，退回 `python-docx` 直接生成。
正文中所有「待确认 / 不确定」项显式标注，不编造。

### 4. 质检（错别字 + 上下文不一致）
通读全文，重点查：
- 错别字、冗余重复（如 `MES / MES 与 WMS`）。
- 字段命名前后不一致（如《流转卡》vs《流程卡》、同一车间两种叫法）。
- 日期 / 地点 / 人物冲突（报告正文 vs 用户后来给的总结口径）。
- 格式不统一（`8/21` vs `8月21日`、多余空格）。
修正统一用 `scripts/docx_qa.py`（见下），不要依赖 editor_sdk 的内容编辑接口。

### 5. 输出总结
约 200 字，固定包含下列要素（顺序可调）：
调研日期 / 调研人 / 地点 → 客户参与部门 → 主营产品 → 生产工艺 → 项目背景 →
核心需求与痛点 → 客户已上线系统 → 下一步工作计划。
模板与样例见 `references/summary_template.md`。

## 关键技术坑（务必先看）
0. **本机 shell 有缺陷，且 pip 装不上包**：
   - `dirname` / `mkdir` / `rm` / `head` / `tail` 等基础命令不可用，`cd` 报 `null directory`。
     一切命令用 **Python 解释器绝对路径**执行；文件增删改一律走 Python（`os.remove` / `os.makedirs`），
     不要用 `rm` / `mkdir`。临时脚本落到工作区内，不要用 `/tmp`。
   - `pip install python-docx` 实测因网络超时**装不上**。不要依赖 python-docx，
     直接用标准库 `zipfile` 读写 `word/document.xml`（本技能脚本已是如此）。
   - Python 解释器路径：`C:/Users/majunpeng/.workbuddy/binaries/python/versions/3.13.12/python.exe`
1. **editor_sdk `open_file` 是异步流式打开**：同一 bash 命令里紧接着调
   `doc_find_and_replace` 会报 `DocEditor: document is not open`。正文编辑不要走 SDK，
   直接用 `scripts/docx_qa.py` 读写 `word/document.xml`。
2. **文件被锁定 → `PermissionError [WinError 5]`**：原文件被预览面板
   （`present_files` 打开的查看器）或编辑器独占。此时 `os.replace` / 写原路径会失败。
   解决：修正结果**另存为新文件名**（如 `原名_修正.docx`）交付；确需覆盖原文件时，
   等用户关闭预览后再用脚本写入。
3. **docx 文本被拆成多个 `<w:t>` run**：字符串常在 run 边界被切断，普通 `in` 查找或
   字符串替换会匹配不到。`scripts/docx_qa.py` 的跨 run 正则替换可处理（每字符间允许
   `</w:t>.*?<w:t>`），且保留其余格式。
4. **模板提取 / 全文提取**：均用 `zipfile` 读 `word/document.xml`，迭代 `body` 的
   `p` / `tbl`；命名空间用 `w:`（正文）与 `a:`（PPT 绘图）。

## 附：docx_qa.py 用法
```bash
# 提取全文（段落+表格），用于通读质检
python docx_qa.py extract <报告.docx>

# 批量修正：rules.json = [["旧词1","新词1"], ["旧词2","新词2"]]
python docx_qa.py fix --in <in.docx> --out <out_修正.docx> --rules rules.json

# 或内联多组替换
python docx_qa.py fix --in <in.docx> --out <out_修正.docx> --pair "旧词||新词" --pair "旧词2||新词2"
```
脚本对每处替换打印命中数；跨 run 文本也能正确替换；输出另存，避免锁文件问题。
