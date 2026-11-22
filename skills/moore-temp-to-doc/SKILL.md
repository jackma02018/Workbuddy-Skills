---
name: moore-temp-to-doc
description: "摩尔售前场景下，把任何临时文件（智能纪要 docx、调研计划 xlsx、现场照片、销售 PDF/HTML 资料、PPT 草稿、腾讯会议录音转写、客户微信/企业微信对话、邮件正文等）转成与给定 Word 模板格式完全一致的正式 .docx 报告。覆盖调研报告、POC 反馈、方案汇报、竞品分析、项目周报、客户汇报等场景。复用模板 styles.xml 只替换正文，零第三方依赖，格式与模板逐像素一致。Trigger phrases: 临时文件转正式报告, 调研纪要整理成报告, 录音转写转文档, 聊天记录出报告, 销售资料编报告, 客户汇报稿, POC 反馈报告, 方案汇报稿, 临时草稿转 docx, 或用户给出一份模板 .docx 加一批散落素材时。"
agent_created: true
---

# 摩尔临时文件转 docx（保格式 · 零依赖）

## 结论

把「**一份 Word 模板** + **任意形态的临时素材**」变成**一份格式与模板完全一致的正式报告 .docx**，可直接用于客户演示、内部评审、归档留档。

引擎仍是「**模板保格式**」：复用模板 zip 里的 styles.xml / theme / fontTable，仅替换正文 XML → 字体、标题配色、表格样式、页面设置与模板 100% 一致；**零第三方依赖**，仅用 Python 标准库。

## 适用场景

| 场景 | 典型素材 | 输出文件名建议 |
|---|---|---|
| **客户调研报告** | 智能纪要 docx + 调研计划 xlsx + 现场照片 | `<客户名称>MES调研报告<YYYYMMDD>.docx` |
| **POC 反馈报告** | POC 日志 docx + 客户打分 xlsx + 现场照片 | `<客户名称>POC反馈报告<YYYYMMDD>.docx` |
| **方案汇报稿** | 方案 PPT + 客户提问笔记 + 录音转写 | `<客户名称>MES方案汇报<YYYYMMDD>.docx` |
| **竞品分析报告** | 竞品官网 HTML/PDF + 对比表 xlsx + 调研笔记 | `<竞品类型>竞品分析报告<YYYYMMDD>.docx` |
| **项目周报 / 月报** | 飞书/企微日报截图 + 项目台账 xlsx | `<项目名称>周报<YYYYWxx>.docx` |
| **会议纪要正式留档** | 腾讯会议转写 txt + 个人速记 md | `<会议主题>会议纪要<YYYYMMDD>.docx` |
| **客户走访记录** | 微信聊天截图 + 录音转写 + 现场照片 | `<客户名称>走访记录<YYYYMMDD>.docx` |
| **招标应答 / 投标资料** | 招标文件 PDF + 应答笔记 md + 资质 docx | `<项目名称>投标应答<YYYYMMDD>.docx` |
| **任意正式留档** | 任意临时文件 | 按客户/项目命名 |

## 何时使用

**触发信号**
- 用户给出一份 Word 模板（`.docx`）＋ 一批散落素材，要求"按这个模板出报告"。
- 素材以**任意形态**存在：docx / xlsx / pdf / html / pptx / jpg / png / mp4 转写 / 聊天截图 / 邮件正文 / md 笔记。
- 用户说"临时文件整理"、"草稿转正式"、"汇报稿"、"调研纪要出报告"、"录音整理成文"。
- 用户要求"格式跟模板一致"、"客户名字+日期"命名。

**不适用**
- 从零创作、**无模板可依** → 用 `tencent-docx`。
- 素材以**图片/录音**为主且需多模态深度解析 → 先走 `mes-customer-research-report`。
- 只是局部改写已生成报告的个别字词 → 用 `mes-customer-research-report/scripts/docx_qa.py`。
- 需要输出**Excel 工作表或 PPT** → 分别用 `tencent-docs-sheet-generation` / `tencent-pptx`。

## 临时文件类型（覆盖四类）

| 类别 | 格式 | 处理方式 | 信息密度 |
|---|---|---|---|
| **A. 办公文档** | `.docx` 智能纪要 / Word 文档 | `template_docx_report.py extract` 抽全文 | ⭐⭐⭐⭐⭐ |
| | `.xlsx` 调研计划 / 台账 / 打分表 | 读 `xl/sharedStrings.xml` + `xl/worksheets/sheetN.xml` | ⭐⭐⭐⭐⭐ |
| | `.pptx` 草稿 / 汇报材料 | 读 `ppt/slides/slide*.xml` 抽纲要 | ⭐⭐⭐ |
| **B. 富文本资料** | `.pdf` 客户资料 / 招标文件 / 竞品手册 | pdftotext 抽文本，扫描版用 Read 工具多模态 | ⭐⭐⭐⭐ |
| | `.html` 网页 / 临时笔记 / 官网存档 | Read 工具或简单正则去标签 | ⭐⭐⭐⭐ |
| | `.md` / `.txt` 笔记 / 录音转写 | 直接 Read | ⭐⭐⭐⭐⭐ |
| **C. 媒体** | `.jpg` / `.png` 现场照片 / 聊天截图 | Read 工具多模态识别 | ⭐⭐⭐ |
| | `.mp3` / `.m4a` / `.mp4` 录音录像 | 先转写为 `.txt`，再用 md 笔记处理 | ⭐⭐⭐⭐（转写后） |
| **D. 通讯** | 微信 / 企业微信聊天记录（截图或文本） | Read 工具或正则解析 | ⭐⭐⭐ |
| | 邮件正文（QQ 邮箱 / 腾讯企业邮 / Outlook） | 直接 Read `.eml` / `.msg` 或粘文本 | ⭐⭐⭐⭐ |

**素材冲突优先级**：Excel 客户书面回复 > 录音转写原文 > 智能纪要正文 > 销售手记 > 本人推断；
推断项必须显式标注「待验证」，**不臆造数据**。

## 输入

| 输入 | 必需 | 说明 |
|---|---|---|
| 模板 `.docx` | ✅ | 提供章节骨架 + 排版样式。**应是已成型的成品报告**，例如上一家客户的同类报告或公司标准模板 |
| 临时素材 | ✅ | 4 类全支持：办公文档 / 富文本 / 媒体 / 通讯 |
| 客户 / 项目名称 | ✅ | 用于正文标题与文件命名 |
| 报告日期 | ✅ | 通常取素材中最新事件日期 |
| 报告类型 | ✅ | 决定章节模板与文件名（见上表） |
| 输出目录 | ❌ | 默认与素材同目录 |

## 输出

- **主交付物**：`<客户/项目名称><报告类型><YYYYMMDD>.docx`，格式与模板完全一致。
- **中间产物**（建议沉淀在 `<workspace>/.workbuddy/`）：
  - `build_report.py` — 本次的内容装配脚本（下次可直接改 JSON 复用）
  - `materials/` — 处理后的素材 markdown 摘要（可选）
  - `memory/YYYY-MM-DD.md` — 当日工作日志
  - `memory/MEMORY.md` — 项目长期事实

## 执行步骤

### Step 0 — 环境自检（必做）

**本机 shell 存在缺陷**：`dirname` / `mkdir` / `rm` / `head` / `tail` 等基础命令不可用，`cd` 会报 `null directory`。
因此：

- 一切命令**用 Python 解释器的绝对路径**执行，不要依赖 shell 内建工具。
- **文件增删改一律走 Python**（`os.remove` / `os.makedirs` / `shutil`），不要用 `rm` / `mkdir`。
- 写临时脚本时，落到目标工作区内（如 `<workspace>/.workbuddy/xxx.py`），不要用 `/tmp`。

```bash
# Python 解释器路径（本机）
C:/Users/majunpeng/.workbuddy/binaries/python/versions/3.13.12/python.exe
```

### Step 1 — 抽取模板骨架（确定章节契约）

```bash
python template_docx_report.py outline <模板.docx>
```

输出：章节标题（含 `pStyle` 样式 ID）、各表格字段名与尺寸、`sectPr` 页面设置。
**这份骨架就是后续报告必须对齐的章节契约**，先看它再动笔。

同时用 `extract` 读模板正文，学习其**行文口吻与字段填写粒度**（例如「现状 / 痛点 / 需求」三分法）。

```bash
python template_docx_report.py extract <模板.docx>
```

**关键：若素材对应的报告类型与模板章节不匹配**（如用调研模板写 POC 报告），先与用户确认是「套模板结构 + 替换章节」还是「完全重排章节」。

### Step 2 — 多源素材并行精读（按类型分别处理）

**A. 办公文档**

- **`.docx` 智能纪要 / Word 文档** → `template_docx_report.py extract` 子命令。AI 纪要通常带「总结 / 智能章节 / 关键决策 / 金句时刻」结构，**正文可读性优于逐段流水**，优先精读「总结」与「关键决策」。
- **`.xlsx` 调研计划 / 打分表 / 台账** → 读 `xl/sharedStrings.xml` 建共享字符串表，再解析各 `xl/worksheets/sheetN.xml` 中 `t="s"` 的单元格索引。**这类文件常含客户对每个问题的逐条书面回复，是需求的第一手依据。**
- **`.pptx` 草稿 / 汇报材料** → 读 `ppt/slides/slide*.xml` 抽每页标题与正文纲要，作为章节线索。

**B. 富文本资料**

- **`.pdf`** → 优先用 `pdftotext`（poppler 工具）抽文本；若为扫描版（无文本层），用 Read 工具多模态识别。
- **`.html`** → Read 工具直接读，或用 `re.sub` 简单去标签抽文本（注意保留段落结构）。
- **`.md` / `.txt`** → 直接 Read。

**C. 媒体**

- **`.jpg` / `.png`** → Read 工具多模态识别，提取手写痛点清单、屏幕截图中的关键数据、聊天记录中的客户原话。
- **录音录像** → 优先调用腾讯会议/通义听悟等已带转写的来源；否则用 `ffmpeg` 抽音频 + `whisper` 转写为 `.txt`，再按 md 笔记处理。**转写文件务必保留来源文件名 + 时戳，便于交叉验证。**

**D. 通讯**

- **微信/企微聊天截图** → Read 工具识别对话，抽取**客户原话**与**承诺事项**（带方括号标注时间戳）。
- **邮件正文** → 直接 Read `.eml` 或粘正文文本。注意邮件链 `>` 引用行的去除，避免正文与"经过"链接混淆。

**素材冲突处理优先级**：Excel 客户书面回复 > 录音转写原文 > 智能纪要正文 > 销售手记 > 本人推断；
推断项必须显式标注「待验证」，**不臆造数据**。

**辅助脚本** `scripts/extract_materials.py` 提供 PDF/PPT/MD/TXT/HTML 的快速摘要入口（见 CLI 速查）。

### Step 3 — 装配内容 JSON

按 `examples/content_example.json` 的格式，把素材转成 `blocks` 数组：

| block type | 用途 | 关键字段 |
|---|---|---|
| `title` / `subtitle` | 封面标题与副信息 | `text`, `size`, `color` |
| `h1` / `h2` / `h3` | 各级标题（自动套用模板样式 3/4/5） | `text` |
| `p` | 正文段落 | `text`, `bullet`, `bold`, `align` |
| `table` | 表格 | `rows`（二维数组）, `widths`（twips，**必须等于内容区宽 8640**） |
| `blank` | 空行 | — |

**行内加粗**：正文里用 `**文字**` 标记，生成时转成真正的 `w:b` 粗体。
**多行单元格**：在单元格字符串里写 `\n` 即可。

### Step 4 — 构建

```bash
python template_docx_report.py build \
  --template <模板.docx> \
  --content  <内容.json> \
  --out      <客户名称><报告类型><YYYYMMDD>.docx
```

原理：把模板 zip 中除 `word/document.xml` 外的**所有部件原样复制**（styles.xml / theme / fontTable / settings），
只把新生成的正文 XML 写进去 → 段落样式、字体、标题配色、页面设置**与模板 100% 一致**。

### Step 5 — 校验（不可跳过）

```bash
python template_docx_report.py verify <输出.docx>
```

必须全部通过：ZIP 完整性 OK、XML 全部良构、标题样式分布与预期一致。
再 `extract` 一遍通读，核对章节完整性、表格行列数、有无错别字。

### Step 6 — 沉淀与交付

1. **首份报告**：把内容装配脚本留在 `<workspace>/.workbuddy/build_report.py`。
2. **换客户/换报告类型（复用）**：用 `extract_content_from_build_script.py` 把上一份脚本转成 `content.json`，
   **只改 JSON、不改引擎**——然后在 JSON 上编辑客户信息与章节，重新 `build` 即可。

   ```bash
   python extract_content_from_build_script.py build_report.py --out wanda.content.json
   # 编辑 wanda.content.json 里的客户名/产能/痛点…，另存为新客户的 content.json
   python template_docx_report.py build --template <模板> --content <新content.json> --out <新报告.docx>
   ```

   该提取器支持解析 `CW` 这类模块级常量（如 `CW-1500` 形式的列宽），并自动跳过 `def table(...)`
   这类函数定义与注释，实测可从 479 行的硬编码脚本完整还原 121 个内容块（8 表格 / 7 标题1 / 17 标题2 / 84 段落），
   零告警，重建结果与原报告**文本完全一致**。

3. 追加当日工作日志与 `MEMORY.md`（客户背景、产能、系统现状、痛点、待办交期）。
4. 用 `present_files` 呈现报告。

## 关键约束

1. **格式一致性优先于实现便利**。不要用 `python-docx` 从空白 Document 重建——那会丢失模板的
   styles.xml、主题色与页面设置，字体和标题颜色都会走样。**必须复用模板 zip 结构。**
2. **表格 `widths` 之和必须等于内容区宽度 8640 twips**（A4 纵向，左右页边距 1800）。
   否则 Word 会拉伸表格或触发自动重排，破坏版式。
3. **零依赖是硬要求**。本机 pip 下载常超时（`python-docx` 实测装不上）。
   不要为生成 docx 而尝试联网安装库，标准库足够。
4. **不要用 `rm` / `mkdir` / `head` 等 shell 命令**，本机 shell 缺陷会导致失败或静默无输出。
5. **文件被预览面板占用会 `PermissionError`**。若目标文件正被打开，
   先输出到新文件名交付，提示用户关闭预览后再覆盖。
6. **不臆造**。需求、指标、流程均须有素材出处；推断项标注「待验证」；缺失字段留「待补」而非编造。
7. **落款与责任人**：调研人员、销售等若素材中未明确，写「待补」，不要猜测姓名。
8. **文件名规范**：`<客户/项目名称><报告类型><YYYYMMDD>.docx`，例如：
   - `江阴旺达电子有限公司MES调研报告20260916.docx`
   - `比亚迪二厂POC反馈报告20261008.docx`
   - `宁德时代周报2026W37.docx`
9. **录音/录像素材**：必须先转写为文本再进入装配；转写文件归档到 `materials/` 备查。
10. **隐私与脱敏**：客户原话、联系方式、内部数据须按用户合规要求脱敏（默认不脱敏，遇敏感字段询问用户）。

## 附：CLI 速查

```bash
PY=C:/Users/majunpeng/.workbuddy/binaries/python/versions/3.13.12/python.exe
S=<skill>/scripts/template_docx_report.py
E=<skill>/scripts/extract_content_from_build_script.py
M=<skill>/scripts/extract_materials.py   # 新增：多源素材摘要

$PY $S extract  <file.docx>                  # 抽取任何 docx 的全文（段落+表格）
$PY $S outline  <template.docx>              # 导出骨架（章节+样式ID+表格字段）
$PY $S build    --template t.docx --content c.json --out o.docx
$PY $S verify   <out.docx>                   # ZIP + XML + 结构校验
$PY $E <build_report.py> [--out c.json]      # 把项目内硬编码脚本转为可复用 content.json
$PY $M <file> [--type pdf|pptx|md|txt|html]  # 把任意素材摘要为 markdown
```

**完整复用链路**（第二家客户起）

```
outline 模板 →  确认章节契约
extract 素材 →  按 4 类临时文件类型分别处理（extract_materials.py / Read / 智能解析）
   ↓ 提炼：客户事实 / 痛点 / 需求 / 待办 / 竞品 / 风险
extract_content_from_build_script（可选）→ 得到基线 content.json
   ↓ 编辑 JSON：换客户名 / 报告类型 / 章节内容
build    →  生成新报告 docx
verify   →  ZIP/XML/结构校验，对照 outline 查漏章节
extract  →  通读终稿，核对错别字与字段一致性
```

## 附：报告章节基线速查

### MES 客户调研报告（默认 7 章）

```
一、项目基本信息                    （表：项目字段 / 内容）
二、项目背景
    2.1 公司背景                   （基础事实 + 产能规模）
    2.2 主要产品                   （表：产品类别 / 代表产品 / 应用领域）
    2.3 主要工艺流程               （主工艺流向 + 关键工序分条）
三、项目调研                        （按部门展开，每部门「现状/痛点/需求」三分法）
    3.1 高层对系统的期望
    3.2 IT 系统现状、网络架构与车间 AP 覆盖   （表：系统名称/供应商/主要功能/应用现状）
    3.3 计划部门      3.4 生产部门      3.5 工艺部门
    3.6 品质部门      3.7 设备部门      3.8 仓库部门
    3.9 生产部门（SMT / DIP / ASSEMBLY / PACK 分组）
    3.10 项目范围与产能确认            3.11 IT 部门
四、核心痛点与诉求汇总              （表：部门 / 核心痛点 / 优先级 / 对应 MES 能力）
五、实施范围与边界建议
六、竞品与风险
七、后续行动计划                    （表：待办事项 / 责任人 / 预计完成时间）
```

### POC 反馈报告（5 章）

```
一、POC 概述                        （表：客户 / POC 范围 / 周期 / 目标）
二、POC 范围与场景                  （表：场景 / 验证目标 / 关键指标）
三、验证结果                        （按场景分小节，含数据截图 + 客户评价）
四、客户反馈与打分汇总              （表：维度 / 客户评分 / 备注）
五、POC 结论与下一步建议
```

### 方案汇报稿（5 章）

```
一、汇报背景与目标
二、客户现状与核心痛点              （从调研报告中提炼）
三、方案蓝图                        （含功能模块图描述 + 实施路径）
四、关键能力亮点                    （按客户关注点分小节）
五、投资预算与实施周期
六、答疑记录                        （汇报现场客户提问 + 答复要点）
```

### 竞品分析报告（6 章）

```
一、分析目的与范围
二、竞品画像                        （表：竞品 / 厂商 / 定位 / 主要客户）
三、核心维度对比                    （表：维度 / 摩尔元数 / 竞品A / 竞品B）
四、优劣分析                        （按维度逐条展开）
五、客户场景适配度                  （按客户行业/规模分小节）
六、应对策略建议
```

### 周报 / 月报（4 章）

```
一、本期关键进展                    （表：事项 / 进度 / 状态）
二、问题与风险
三、下期计划
四、需领导协调事项
```

### 会议纪要正式留档（3 章）

```
一、会议基本信息                    （表：时间 / 地点 / 主持人 / 参会人 / 主题）
二、议题与决议                      （按议题分小节，含决议与责任人）
三、待办与跟进                      （表：待办 / 责任人 / 截止日期）
```

**样式 ID 对照**（模板 `styles.xml`）：标题1=`pStyle 3`、标题2=`pStyle 4`、标题3=`pStyle 5`、
表格=`tblStyle 33`、项目符号=`pStyle 16`；正文微软雅黑，标题主题色 `376092` / `4F81BD`。