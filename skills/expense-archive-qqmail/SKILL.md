---
name: expense-archive-qqmail
description: "当天邮件里的所有报销票据（12306 火车票/退票、滴滴与曹操网约车发票及行程单）→ 下载 → 只取 PDF → 按「年月日_类别_金额.pdf」重命名 → 按文件自身业务日期归入对应差旅日期文件夹。用 REST 直连下载避免上下文爆炸。触发场景：整理今天的报销邮件、归档交通费发票、下载并归类差旅票据。"
version: 1.0.0
display_name: 差旅票据邮件归档
disable-model-invocation: false
agent_created: true
---

# 差旅票据：当天邮件 → 重命名 → 按日期归档

把「邮箱里当天的一堆报销邮件」变成「按差旅行程分好文件夹、整齐命名的 PDF」。
覆盖 12306 火车票发票/退票 + 滴滴 & 曹操网约车发票与行程单。

## 何时使用

- "把今天邮箱里的发票下载归档" / "整理今天的报销票据" / "按日期归到差旅文件夹"
- 目标目录是一堆 `20260911~20260913 总部售前培训` 这类**日期区间文件夹**时

## 铁律

**不要用 MCP 的 `DownloadAttachment` 批量下载。** 附件以 base64 进工具返回值，
单个 42KB ≈ 1.4 万 token，十几个就打爆上下文。改用 `railway-invoice-mail` 技能的
REST 直连脚本落盘。若确实派了子代理去下载，主代理仍应优先走 REST。

**先用 `ListMessages` 拿到当天邮件数与 ListAttachments 清单做交叉核对**，确保无遗漏。

## 流程

### 1. 列出当天带附件邮件

```
ListMessages: dir=inbox, after=<今天零点 ISO8601 带时区>, has_attachments=true, limit=50
```

核对 `has_more=false`；再**去掉 `has_attachments` 重跑一次**，确认当天邮件总数一致
（防止有附件却未被标记的邮件）。

### 2. ListAttachments 取清单，写成 manifest JSON

每项字段：`{"message_id","attachment_id","filename","size"}`。
建议按来源分两个 manifest（12306 / 网约车），输出目录也分开。

### 3. 解密 token —— 注意过期

```bash
python <railway-invoice-mail>/scripts/decrypt_token.py --out <临时目录>/.qqmail_token
```

> **坑 1：token 可能已过期**（`expiresAt` 是毫秒时间戳）。过期时 REST 直接 401，
> 但 MCP 工具仍正常。修复：**先随便调一次 `mcp__qq-mail__GetMe` 触发连接器续期**，
> 再重跑 decrypt_token.py。
>
> **坑 2：Windows 托管 Python 不认 Git Bash 盘符**。脚本路径必须写
> `C:/Users/<用户名>/.workbuddy/skills/...`，不能用 `/c/Users/...`。
> 命令前还需 `export PATH="/usr/bin:/bin:/usr/local/bin:$PATH"`。

### 4. 下载

```bash
python <railway-invoice-mail>/scripts/fetch_attachments.py \
    --manifest m.json --out ./原始附件/12306 --token <token> --alias <alias_id>
```

`--alias` 从 GetMe 结果取，可省一次请求。脚本内置 6.5s 间隔 + 429 退避，可重复跑补漏。
**下载完立刻删除 token 文件。**

### 5. 解压只保留 PDF

12306 的 zip 内含 `.pdf` + `.ofd`，只读 PDF。注意中文文件名可能是 cp437 编码伪装的 GBK。

### 6. 解析：必须按坐标还原版式

```bash
python scripts/parse_invoices.py --train <pdf目录> --ride <pdf目录> --json parsed.json
```

`get_text("text")` 会把「价税合计（大写）/（小写）/ 陆拾圆零肆角贰分 / ¥60.42」
拆到四处，**必须**用 `get_text("words")` 按 `(round(y), round(x))` 排序分行再取。

票面规则：

| 票种 | 识别标志 | 金额 | 日期 |
|---|---|---|---|
| 12306 火车票 | 含 `电子客票号` | `￥(\d+\.\d{2})` | 排除「开票日期」行后的 `YYYY年MM月DD日` = **乘车日期** |
| 12306 退票 | 同上但含 `退票费` | 同上 | 同上 |
| 网约车发票 | 含 `电子发票（普通发票）` | 「（小写）」行 x 最大的数值 | 「出行日期」表体行的日期 |
| 行程报销单 | 含 `行程起止日期` | `合计([\d.]+)元` | `行程起止日期：A 至 B` |

- 数值 token 可能被渲染成**单个 `¥60.42`**（曹操是合并的，滴滴是拆开的），要先剥 `¥/￥/,`。
- **数电发票可能没有「出行日期」明细**：用**等额行程单反推**日期。

### 7. 重命名 + 归档

```bash
python scripts/apply_archive.py --root <报销根目录> --parsed parsed.json
```

- 文件名：`YYYYMMDD_类别_金额.pdf`（`YYYYMMDD` 取**当天**；金额两位小数）
- 类别字典：`train→火车发票`、`ride_invoice→网约车发票`、`trip_sheet→网约车行程`、
  含「退票费」→`退票`
- **曹操出行**：虽然开票方是「苏州市吉利优行电子科技有限公司」，但走滴滴开票入口。
  与滴滴统一叫「网约车发票/网约车行程」，避免名不符实（用户已确认此口径）。
- 归档：解析文件夹名里的日期区间，与文件业务日期区间**取交集**即归入。
  注意文件夹名可能有**年份笔误**（`20260911~20110913` 应为 `20260913`），
  需按起始年份校正后再比对，否则漏匹配。

### 8. 交叉校验

- 用邮件摘要里的金额/日期与 PDF 解析结果对账
- 行程单：明细行金额之和 == 「合计 X 元」
- 汇报时**必须点出仍是空文件夹**（可能意味着该趟行程的票据还没开票）

## 依赖

托管 venv 装 `cryptography`、`pymupdf`、`openpyxl`（`railway-invoice-mail` 的下载脚本依赖 cryptography）。
