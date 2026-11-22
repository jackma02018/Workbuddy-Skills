---
name: batch-invoice-processor
description: "批量提取电子发票结构化字段（代码、号码、日期、购销方、金额税额等），自动归类 14 种费用类型、核验金额一致性、检测重复与 7 类异常，按人员维度统计并输出 Excel 汇总表、CSV 备用文件与 HTML 审计报告。另支持运费增值税专用发票的 ERP 预处理（PDF 重命名与 JPG 生成）。适用于发票整理、报销汇总、进项票据统计、发票对账、运费发票 ERP 上传等场景。"
agent_created: true
version: 1.0.0
display_name: "批量发票处理"
display_name_en: "Batch Invoice Processor"
description_zh: "批量提取电子发票结构化字段，自动归类费用、核验金额、检测重复与异常，输出Excel汇总、CSV与HTML审计报告，支持运费专票ERP预处理。"
description_en: "Batch-extract structured e-invoice fields, auto-categorize expenses, verify amounts, detect duplicates/anomalies, and output Excel/CSV/HTML audit reports."
visibility: "public"
---

# 批量发票处理（batch-invoice-processor）

面向企业财务报销场景的发票批处理技能：从多种格式的电子发票中抽取结构化字段，自动归类费用、核验金额一致性、检测重复与异常，最终输出 **Excel 汇总表（3 个工作簿）+ CSV 备用文件 + HTML 审计报告**。额外提供运费增值税专用发票的 ERP 预处理能力（PDF 重命名 + JPG 生成）。

## 适用场景

满足以下任一条件时使用本技能：

- 用户提供一批发票文件（PDF/OFD/XML/图片）或整个文件夹，要求「整理」「汇总」「对账」「出报销清单」。
- 用户已通过其他工具做了初步提取，但缺少费用归类、金额核验或台账产出环节。
- 用户需要对运费增值税专用发票进行 ERP 系统上传前的格式准备（重命名 + 转 JPG）。

## 前置依赖

### 通用批处理管道（Node.js）

- **Node.js**：脚本运行环境。WorkBuddy 内置受管 Node（v22），无需用户额外安装。
- **npm 包依赖**：脚本优先加载本地 `node_modules`，找不到时自动回退到 WorkBuddy 受管 Node 工作区。
  首次使用可在 `scripts/` 目录下执行：
  ```bash
  npm install fast-xml-parser jszip
  ```
  - `fast-xml-parser`：解析增值税电子发票 XML 结构化数据；缺失时 XML 发票降级为需视觉读取。
  - `jszip`：解压 OFD 发票取文本层，以及输出 Excel（.xlsx）；缺失时 OFD 降级为需视觉读取，Excel 退回 CSV。
  - 图片/PDF 文本通道零外部依赖（仅用 Node 内置 `fs`/`path`/`os`/`zlib`）。

### ERP 预处理模式（Python，可选）

仅在使用运费发票 ERP 预处理功能时需要：

```bash
pip install pdfplumber pymupdf
```

- `pdfplumber` — PDF 文本内容提取
- `pymupdf` (fitz) — PDF 页面渲染为图片

## 处理流水线（通用批处理）

整体按「输入分流 → 文本提取 → 视觉补全 → 汇总校验 → 输出呈现」五段式流水线执行。

### 输入分流

- 确认用户提供的发票文件路径（单文件或目录）。
- 按扩展名区分通道：
  - `.pdf` / `.ofd` / `.xml` / `.txt` → 进入文本提取通道（脚本处理）。
  - `.png` / `.jpg` / `.jpeg` / `.bmp` / `.webp` → 直接进入视觉读取通道。
- 若输入为目录，递归收集上述类型文件，**跳过**以「结果/汇总/提取」开头的文件名（防止重复处理已输出的产物）。
- 可提示用户遵循附件规范（见 `references/classification.md` 第六节）：每张发票单独一个文件、优先提供含文本层的原文件、文件名可含报销人-费用-金额-日期。

### 文本提取（脚本通道）

对每个文件运行提取脚本，输出统一字段 JSON：

```bash
node <技能目录>/scripts/extract.js "<目标文件或目录>" --json 提取结果.json
```

脚本内部处理路径：
- XML：结构化解析增值税电子发票；
- OFD：解压 ZIP 包取内嵌 XML 文本；
- PDF：零依赖抽取内容流文本 + ToUnicode CMap 中文还原；
- 纯文本：正则提取。

每条记录输出 `confidence`（高/中/低/无）与 `needs_vision` 标记。当 `needs_vision=true`（图片型 PDF 无文本层、纯图片文件、或 OFD 缺 jsZip 依赖）时，字段留空，转下一阶段用 Read 工具补全。

> 脚本以逐文件 try/catch 做崩溃隔离，单个文件异常不影响整批。

### 视觉补全（替代外部 OCR）

对上一阶段标记 `needs_vision=true` 的文件，使用本平台 **Read 工具**读取图片或 PDF（PDF 按页呈现可视内容）：

1. 用 Read 打开文件，获取可视内容。
2. 参照 `references/classification.md` 第一节的字段清单，从图中提取：
   `invoice_code, invoice_number, invoice_date, invoice_type, buyer_name, buyer_tax_id, seller_name, seller_tax_id, amount_ex_tax, tax_amount, total_amount, tax_rate, check_code, remarks`。
3. 能识别报销人时填写 `reimbursement_person`；无法识别则留空（后续人工补全）。
4. 将提取结果**写回** `提取结果.json` 中对应记录，清除 `needs_vision` 标记、设置 `confidence`（如「读图」）、`page`（PDF 多页时填页码）。
   - 无法识别的字段留空字符串，并在 `remarks` 注明「未识别」。
   - 若某页不是发票（如支付凭证、银行通知），保留记录并设 `needs_vision=false`、`invoice_type` 为空、`remarks` 写「非发票页」，供处理日志记录。

> 若用户已连接腾讯文档 OCR（`mcp__tencent-docs__ocr.*`）且偏好 OCR 方式，可改用其提取图片/PDF 文本，再映射到同一字段结构。

### 汇总校验（脚本通道）

将补全后的 `提取结果.json` 交给汇总脚本：

```bash
node <技能目录>/scripts/organize.js --input 提取结果.json --out 发票汇总_YYYY-MM-DD.xlsx
```

脚本自动完成以下处理：
- 费用归类（`expense_category`，规则见 `references/classification.md` 第二节）；
- 金额逻辑核验：`amount_ex_tax + tax_amount = total_amount`（铁路/公路客票标记「免校验」）；
- 查重（全电发票按 20 位号码）；
- 异常检测（金额不一致/重复/税率为 0/金额异常大/日期异常/连号/同一销售方 ≥ 5 张）；
- 按人员维度统计；
- 生成 **Excel（发票明细 / 按人员统计 / 处理日志 3 个工作簿）** 与 **CSV 备用文件**，并输出统计摘要（用于生成 HTML 报告）。

输出文件名建议带日期后缀：`发票汇总_YYYY-MM-DD.xlsx`。

### 输出呈现

- 将 Excel / CSV 呈现给用户（使用 present_files）。
- 将脚本输出的统计摘要整理为**自包含 HTML 审计报告**（内联 CSS，使用 `references/invoice_report_template.html` 作为样式参考），保存为 `发票汇总_YYYY-MM-DD.html` 并呈现。
- 回复中给出：
  - 总计：**N 张发票，价税合计 ¥X,XXX.XX**；
  - 按人员统计表；
  - 处理情况：重复张数与明细、异常张数与类型、仍需视觉补全的文件（如有）。
- 若发现重复/异常/未补全，明确列出并建议人工确认；必要时根据用户指示补录或对接报销系统。

## ERP 预处理模式 — 运费发票上传准备

适用于运费类增值税专用发票的 ERP 系统上传前格式准备。此为独立于通用批处理管道的专用工具，通过 Python 脚本直接处理。

### 执行方式

```bash
python scripts/sf_invoice_process.py "file1.pdf" "file2.pdf" ...
```

脚本自动完成信息抽取、PDF 重命名、JPG 生成，最后输出成功/失败计数。产物均落在各 PDF 所在目录。

### 处理流程

**阶段一：环境自检**
```bash
pip install pdfplumber pymupdf
```

**阶段二：信息抽取**（pdfplumber 打开 PDF 首页，提取文本后按规则解析）
- 发票类型校验 — 文本中必须含「增值税专用发票」，否则跳过
- 发票号码 — 正则 `发票号码[：:](\d+)`
- 销方名称 — 匹配所有「名称：XXX有限公司」，取最后一条（销方位于购方之后）
- 价税合计 — 正则 `（小写）[¥￥]([\d.]+)`，取含税总额（非不含税金额）
- 物流单号 — 正则 `(SF\d+)\s*备\s*注|备\s*注\s*(SF\d+)`，找不到则报错跳过

**阶段三：文件产出**
- PDF 重命名：`¥{价税合计}+{销方名称}+{发票号码}+{物流单号}.pdf`，目标名已存在则告警跳过
- JPG 生成：pymupdf 将首页渲染为 200 DPI（`fitz.Matrix(200/72, 200/72)`），已存在则告警跳过

> 注意：ERP 预处理模式使用 pdfplumber 提取文本（非 Node.js 管道），命名格式和产出与通用批处理模式不同。两种模式的产物不可互换——ERP 模式专门适配运费发票的上传格式要求，通用批处理模式用于财务归档与审计。

### 运费发票票面格式参考

增值税专用发票电子版的典型文本布局（示例数据已泛化）：

```
电子发票（增值税专用发票）
发票号码：12345678901234567890
开票日期：20XX年XX月XX日
购 名称：XX贸易有限公司
...
销 名称：XX物流有限公司
...
价税合计（大写） X仟X佰X拾X圆整 （小写）¥XXX.XX
...
SF1234567890123    ← 物流单号可能在备注之前
备
注
```

或物流单号在备注之后：

```
价税合计（大写） X仟X佰X拾X圆整 （小写）¥XXX.XX
备
注
SF1234567890123    ← 物流单号可能在备注之后
```

## 大批量处理策略

- 文本通道支持整目录一次性提取；视觉通道按页进行，图片较多时分批（每批 ≤ 8 张图片）以保持稳定。
- 所有提取结果先汇总到单个 `提取结果.json`，再统一跑 `organize.js`（跨文件去重/校验在脚本内一次性完成）。

## 已知局限（使用前须知）

- PDF 文本通道支持未压缩或 FlateDecode 压缩的内容流；部分 ASCII85/LZW/PDF1.5+ 对象流压缩的发票可能漏字段（置信度标「低」）。
- 纯图片 PDF、扫描件、手机截图：**无文本层**，必须走视觉读取阶段（或 OCR）；数字字段（发票号、小写金额）若图中使用特殊字体仍可能提取失败，需人工补录。
- 未做发票真伪查验（可对接增值税发票查验平台作为后续扩展）。
- Excel 输出依赖 jszip；缺失时退回 CSV。

## 参考文件

- `references/classification.md`：字段定义、费用归类规则、异常判定标准、发票类型识别、Excel 列顺序、附件规范、特殊场景处理说明。
- `references/invoice_report_template.html`：HTML 审计报告模板（Wise 风格绿色主题，内联 CSS），使用 `{{占位符}}` 标记需替换的变量。

## 文件清单

- `scripts/extract.js` — 文本层提取：PDF/OFD/XML/图片 → 统一字段 JSON
- `scripts/organize.js` — 汇总校验：费用归类、金额核验、去重、异常检测、Excel+CSV 输出
- `scripts/sf_invoice_process.py` — ERP 预处理：运费专票 PDF → 重命名 + JPG 生成（Python，独立工具）
- `references/classification.md` — 字段规范、归类规则、异常标准
- `references/invoice_report_template.html` — HTML 审计报告模板

---

> **免责声明**：本工具仅提供发票数据的本地提取、整理和格式化输出能力，所有数据处理均在本机完成。使用本工具提取的票据数据仅供参考，具体财税处理请以税务机关和会计准则为准。
