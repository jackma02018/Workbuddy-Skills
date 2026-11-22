---
name: railway-invoice-mail
description: "从邮箱（QQ 邮箱）抓取 12306 铁路电子发票邮件，批量下载附件 zip，解压出 PDF，按「出发站-到达站-乘车日期-金额.pdf」重命名，并生成带 =SUM() 合计的 Excel 汇总。触发场景：整理邮箱里的火车票/12306 发票、报销交通费、下载车票发票附件、把电子客票改名归档。本技能自带绕过上下文爆炸的附件下载方案（直接调 REST API 落盘）。"
description_zh: "从 QQ 邮箱批量抓取 12306 铁路电子发票附件，解压 PDF 并按「出发站-到达站-乘车日期-金额」重命名，生成带合计公式的 Excel 汇总，用于交通费报销归档。"
description_en: "Fetch 12306 railway e-invoice attachments from QQ Mail in bulk, extract PDFs, rename them as origin-destination-date-amount, and generate an Excel summary with SUM totals for travel-expense reimbursement."
version: 1.0.0
display_name: "12306 邮箱发票助手"
display_name_en: "12306 Email Invoice Helper"
disable-model-invocation: false
agent_created: true
---

# 12306 电子发票：邮箱抓取 → 重命名 → Excel 汇总

把「邮箱里的一堆 12306 发票邮件」变成「一文件夹整齐命名的 PDF + 一张汇总 Excel」。
解析与重命名复用 `railway-invoice` 技能；本技能补上**从邮箱取附件**这一段。

## 何时使用

- "整理邮箱里今天的 12306 发票" / "把这批火车票发票下载下来改名" / "统计这个月交通费"
- 前提：发票邮件在 QQ 邮箱里（Agent 邮箱若未开通则不可用）

## 关键约束（务必先读）

**不要用 MCP 的 `DownloadAttachment` 批量下载。** 它把附件以 base64 塞进工具返回值——
单个 42KB 附件约 1.4 万 token，20 个就是 ~28 万 token，足以打爆上下文。
12306 一封发票邮件常带 10 个附件，两封就是 20 个。改用本技能的 REST 直连脚本落盘。

**限流 10 次/分钟。** `fetch_attachments.py` 已内置 6.5s 最小间隔 + 429 退避重试，
且会跳过已存在的文件，可安全重复运行补漏。

## 流程

### 1. 找到发票邮件

```
SearchMessages: q="12306", after=<今天零点 UTC-8h>, has_attachments=true
```

12306 发票邮件主题固定为「网上购票系统-电子发票通知」，一封邮件带多个附件
（每个附件一个 zip，内含 `.pdf` + `.ofd`，zip 名即发票号码）。
注意别把「用户支付通知 / 用户改签通知」也算进来——它们没附件。

用 `ListAttachments` 拿到 `attachment_id` 与 `filename`，写成 manifest JSON：

```json
[
  {"message_id": "msg_xxx", "attachment_id": "att_xxx", "filename": "26329116804010105185.zip", "size": 42327}
]
```

### 2. 取出邮箱 token（解密本地凭证，token 不进上下文）

```bash
python3 <SKILL_DIR>/scripts/decrypt_token.py --out /tmp/.qqmail_token
```

凭证在 `~/.workbuddy/connectors/<userId>/.credentials.v3.json`，加密参数：
AES-256-GCM，HKDF-SHA256，**PRK = HMAC(key=salt, msg=ikm)**，
ikm = masterKey(32B，同目录 `.master.key`) || userId，salt 取自 header，
info = `workbuddy-oauth-credentials-v1`；AAD = `{userId}|{serverKey}|{fieldName}`
（serverKey 形如 `qq-mail|xxxxxxxx`）。

> 坑：HKDF extract 的顺序是 HMAC(salt, ikm)，写反会一直报 InvalidTag。

### 3. 下载附件

```bash
python3 <SKILL_DIR>/scripts/fetch_attachments.py \
    --manifest attachments.json --out ./原始附件 --token /tmp/.qqmail_token
```

用完清理 token：`rm -f /tmp/.qqmail_token`

### 4. 从邮件正文提取清单（可选但强烈建议，用于交叉校验）

```bash
python3 <SKILL_DIR>/scripts/verify_from_mail.py msg_xxx msg_yyy
```

正文是 HTML，脚本先剥标签再匹配。输出张数与合计金额，用作第 5 步的基准。

### 5. 解析 PDF + 重命名 + 生成 Excel

```bash
python3 ~/.workbuddy/skills/railway-invoice/scripts/parse_railway_invoices.py \
    "原始附件" "重命名PDF" --summary "交通费总额.xlsx"
```

- 取**乘车日期**（与发车时间同一行的日期），不是开票日期
- 出发站 = 版面最左的站，到达站 = 最右的站
- 文件名格式 `出发站-到达站-乘车日期-金额.pdf`（默认分隔符 `-`，脚本 `--sep` 可改），重名自动加 `_2`
- 示例：`南京南站-镇江南站-2026年08月03日-32.00.pdf`
- 汇总表 `交通费汇总`：序号/出发站/到达站/乘车日期/金额(元)，合计用 `=SUM()`

### 6. 重算公式并校验

```bash
python3 <XLSX_SKILL>/scripts/recalc.py "<汇总文件>.xlsx"   # status 必须为 success，0 errors
```

### 7. 收尾

- 把汇总 xlsx 从 `原始附件/` 移到顶层并改成带日期的名字，例如 `12306电子发票汇总_2026-09-02.xlsx`
- 删掉解压临时目录 `原始附件/.railway_tmp`
- 汇报：张数、总额（取 Excel 合计单元格）、文件位置；`[skip] 无法解析` 的需人工核对

## 排错（实战踩过的坑）

### 1. `HTTP 401 Unauthorized` —— token 已过期（最常见）

本地凭证里的 `accessToken` 是有有效期的（`mcpOAuth.<serverKey>.expiresAt`，毫秒时间戳），
过期后 REST 直连一律 401，**但 MCP 工具调用仍然正常**（连接器会在调用时自动续期）。

判定与修复：

```bash
python -c "import json,time;c=json.load(open(r'<conn-dir>/.credentials.v3.json'));\
v=c['mcpOAuth']['qq-mail|<hash>'];print(v['expiresAt']-int(time.time()*1000))"
# 结果为负 = 已过期
```

修复方法：**先随便调一次 MCP 工具**（如 `mcp__qq-mail__GetMe`）触发连接器续期，
再重新执行 `decrypt_token.py`。续期后校验再跑 `fetch_attachments.py` 即可。
这一步不需要用户重新授权，也不需要手动刷新 OAuth。

### 2. Windows 下 Python 脚本路径必须写成 Windows 风格

托管解释器是原生 Windows 程序，不认 Git Bash 的盘符映射：

- 错（报 `can't open file 'd:\c\Users\...'`）：`python /c/Users/.../decrypt_token.py`
- 对：`python "C:/Users/<用户名>/.workbuddy/skills/railway-invoice-mail/scripts/decrypt_token.py"`

输出目录同理用 `D:/...`。另外 Git Bash 里命令前需先
`export PATH="/usr/bin:/bin:/usr/local/bin:$PATH"`，否则 `ls`/`dirname` 会报 command not found。

### 3. 显式传 `--alias` 省一次请求

`alias_id` 可从 MCP 的 `GetMe` 结果里直接拿到（`aliases[].alias_id`），
传给 `fetch_attachments.py --alias` 可跳过 `/v1/me` 探测。

### 4. 依赖技能 `railway-invoice` 可能没装

第 5 步的 `parse_railway_invoices.py` 属于**另一个技能** `railway-invoice`。
若该目录不存在，下载这半段仍然可用，自己写解析器即可 —— 见下方「自建解析器要点」。

## 自建解析器要点（pymupdf）

- **必须按坐标还原版式**：`page.get_text("words")` 按 `(round(y), round(x))` 排序分组。
  `get_text("text")` 会把「价税合计（大写）/（小写）/ 陆拾圆零肆角贰分 / ¥60.42」
  拆到四处，导致取错金额。
- **12306**：`￥(\d+\.\d{2})` 取金额；「开票日期」那一行要排除掉，剩下的
  `YYYY年MM月DD日` 才是**乘车日期**；含 `退票费` 即退票票，`票价` 为普通车票。
- **ZIP 内含 `.pdf` + `.ofd`**，只取 PDF（用户通常只要 PDF）。
- **网约车发票**：`(小写)` 那一行取 x 最大的数值设为价税合计；注意数值可能被渲染成
  **单个 token `¥60.42`**（滴滴是拆开的 `¥` + `60.42`，曹操是合并的），需先剥掉 `¥/￥/,`。
- **数电发票可能没有「出行日期」明细**（受票方信息缺失），此时用**等额行程单反推**
  行程日期：找金额完全一致的行程报销单，取其「行程起止日期」。

## 命名口径

非必须以 `出发站-到达站-日期-金额` 命名时（例如用户要求 `年月日_类别_金额`），
自行写 `apply.py`：

- 类别字典建议：`train→火车发票`、`ride_invoice→网约车发票`、`trip_sheet→网约车行程`、
  含「退票费」的→`退票`。**滴滴/曹操统一叫「网约车」**可避免第三方开票
  （曹操实由「苏州市吉利优行电子科技有限公司」开票，却走滴滴开票入口）带来的歧义。
- 日期文件夹**决定归档落位**：解析文件夹名中的日期区间
  `20260911~20160913`，与文件业务日期区间取交集。注意文件夹名可能有年份笔误
  （如 `20260911~20110913`），需按起始年份校正后再比对，否则会漏匹配。

## 依赖

`cryptography`、`pdfplumber`、`openpyxl`（解析器若自建则改用 `pymupdf`）。建议装进托管 venv：

```bash
/Users/xhz/.workbuddy/binaries/python/envs/default/bin/pip install cryptography pdfplumber openpyxl
```

## 验收清单

- [ ] PDF 张数 == 邮件正文清单条数
- [ ] PDF 解析总额 == 邮件正文合计金额
- [ ] 抽样几个文件名里的日期是**车票日期**，不是今天/开票日期
- [ ] recalc 返回 `{"status":"success","total_errors":0}`
