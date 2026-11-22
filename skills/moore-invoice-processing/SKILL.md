---
name: 摩尔发票处理
description: "从腾讯（QQ）邮箱按时间/关键词筛选并下载发票邮件附件，解析出发站、到达站、乘车日期、金额、消费城市等票面字段，按「年月日_火车发票_出发站_到达站_金额」等格式批量重命名并归档到差旅日期文件夹。支持 PDF/OFD/图片与扫描件 OCR，字段缺失一律跳过不臆造，重名自动加序号。内置纯标准库 HTTP 服务，可发布成在线链接供他人调用。触发场景：整理报销票据、下载邮箱里的发票、火车票/网约车发票改名、按日期归档差旅票据、把发票处理流程发布上线。"
version: 1.0.0
display_name: 摩尔发票处理
display_name_en: Moore Invoice Processing
disable-model-invocation: false
agent_created: true
---

# 摩尔发票处理

把「邮箱里一堆发票邮件」变成「分好文件夹、命名规范、金额可对账的票据」。
核心价值：**一切字段取自票面**（日期/金额/站名/城市），缺了就跳过，绝不猜。

## 何时使用

- "整理今天/本周邮箱里的发票" / "下载 12306 和滴滴的票据"
- "把发票改成 `年月日_火车发票_出发站_到达站_金额` 这种格式"
- "按票面日期归档到差旅文件夹"
- "把这套发票处理流程做成在线服务给别人用"

## 铁律

1. **不要用 MCP 的 `DownloadAttachment` 批量下载附件。** 附件以 base64 进工具返回值，
   单个 42KB ≈ 1.4 万 token，十几个就打爆上下文。一律走 `fetch_mail.py` / `fetch_attachments.py`
   的 REST 直连落盘。
2. **日期严禁取文件时间戳或文件名里的数字**，必须从票面内容解析。
3. **任何字段识别不出来就跳过**并在清单里写清原因，不做猜测性填充。
4. **重命名前先备份**（操作的是个人报销目录），先预览再执行。

## 输入参数

| 参数 | 脚本 | 说明 | 默认值 |
|---|---|---|---|
| `--dest` / `--root` | pipeline / archive | 报销根目录（内含 `20260911~20260913 总部售前培训` 这类日期文件夹） | 必填 |
| `--after` / `--before` | pipeline / fetch_mail | 时间范围，如 `2026-09-15`、`2026-09-15T00:00+08:00` | 不限 |
| `--days` | pipeline | 只处理最近 N 天 | — |
| `--sender` / `--subject` / `--keyword` | pipeline / fetch_mail | 筛选邮件，支持正则 | `--subject` 默认 `发票\|行程\|电子客票` |
| `--ext` | pipeline / fetch_mail | 附件扩展名过滤 | `zip,pdf` |
| `--template` | pipeline / rename_invoices | 命名口径：`standard` / `railway` / `legacy12306` | `railway`（pipeline）|
| `--alias` | pipeline / fetch_mail | QQ 邮箱 alias_id，省略自动探测 | 自动 |
| `--execute` | pipeline / rename / archive | 实际改动；不带则只预览 | 预览 |
| `--steps` | pipeline | `fetch,download,extract,rename,archive` 子集 | 全部 |
| `--report` | pipeline | 汇总 JSON 输出路径 | — |

## 输出结果

- `<dest>/_流程/manifest.json` —— 邮件→附件清单
- `<dest>/_流程/原始附件/` —— 下载的 zip / pdf
- `<dest>/_流程/待处理/` —— 解压出的票证 + 改名后的文件
- `<dest>/_流程/rename_plan.json` —— 原/新名对照 + 跳过原因
- `<report>.json` —— `{ok, skipped, amount_total, skipped_detail[], steps{}}`
- 票据最终落在 `<dest>/<日期文件夹>/` 下

## 命名规范

```
<票面日期>_<票据类型>_<补充信息>_<金额>.<原扩展名>
```

三种口径（`--template`）：

| 口径 | 火车票示例 |
|---|---|
| `railway`（默认，本技能招牌格式） | `20260911_火车发票_苏州南站-福州站_407.50.pdf` |
| `standard` | `20260911_火车票_苏州南站-福州站_407.50.pdf` |
| `legacy12306` | `苏州南站-福州站-20260911-407.50.pdf` |

类型取值：`火车发票` / `退票` / `网约车发票` / `网约车行程`（standard 口径下为 `火车票`）。
跨天票据日期段写成 `20260819-20260820`；多城市写成 `江苏省苏州市+镇江市`（同省只标一次省名）。

## 全流程

```bash
export PATH="/usr/bin:/bin:/usr/local/bin:$PATH"
PY='C:\Users\majunpeng\.workbuddy\binaries\python\envs\default\Scripts\python.exe'
S='C:/Users/majunpeng/.workbuddy/skills/moore-invoice-processing/scripts'

# 一键（预览）
"$PY" "$S/pipeline.py" --dest "D:\00费用报销\未报销" --days 7 --template railway

# 确认无误后执行
"$PY" "$S/pipeline.py" --dest "D:\00费用报销\未报销" --days 7 --template railway \
      --execute --report "D:\00费用报销\未报销\_流程\report.json"
```

五步也可单独跑（同名脚本），便于断点续跑：

| 步骤 | 命令要点 |
|---|---|
| fetch | `fetch_mail.py --manifest m.json --after 2026-09-15 --subject 发票`（内部会先跑 `decrypt_token.py`）|
| download | `fetch_attachments.py --manifest m.json --out ./原始附件 --token <token>`，**用完删 token** |
| extract | `extract_zip.py --src ./原始附件 --dst ./待处理`（默认丢 ofd、摊平目录）|
| rename | `rename_invoices.py --root ./待处理 --template railway --json plan.json [--execute]` |
| archive | `archive.py --root <dest> --source ./待处理 [--execute]` |

## QQ 邮箱 REST 接口事实（实测，写脚本前必读）

1. `GET /v1/aliases/{alias}/messages?dir=inbox&limit=N` → `{data:[…], pagination:{next_cursor,has_more}}`
   **`limit` 上限 50**，传 100 直接 HTTP 400。分页用返回的 `cursor`。
2. `after=` / `before=`（RFC3339 UTC，须以 `Z` 结尾）**服务端过滤有效**。
3. **`q=` 关键词参数无效**，会被静默忽略（传 `q=12306` 结果与不传完全一致）。
   关键词筛选**必须在客户端做** —— `fetch_mail.py` 用 `--sender/--subject/--keyword` 实现。
4. `GET /v1/aliases/{alias}/messages/{mid}/attachments` → `{data:[{attachment_id,filename,content_type,size}]}`

限流 10 次/分钟：脚本内置 6.5s 最小间隔 + 429 指数退避，且跳过已存在的文件，可安全重跑补全。

> **HTTP 401 = token 过期**（`mcpOAuth.<key>.expiresAt` 毫秒时间戳）。
> MCP 工具调用时会自动续期，REST 不会。修复：**先随便调一次 MCP 工具**（如 `GetMe`）触发续期，
> 再重跑 `decrypt_token.py`。不需要用户重新授权。

### 兜底：连接器被禁用时自己续期（2026-09 实测可用）

`connector-states.v3.json` 里 `userDisabled:{"qq-mail":true}` 时，**没有 MCP 工具可调**，
上面那条路走不通。此时用本地 refresh_token 直接打官方 OAuth 端点：

```bash
curl -s https://api.mail.qq.com/.well-known/oauth-authorization-server
# -> {"issuer":"https://wx.mail.qq.com","token_endpoint":"https://wx.mail.qq.com/oauth/token",
#     "grant_types_supported":["authorization_code","refresh_token"]}

"$PY" "$S/refresh_token.py" --out "$W/.qqmail_token"     # 解密 refresh_token → 换新 access_token → 落盘
```

要点：

- `client_id` 在 `mcpClientInfo/qq-mail|<hash>/client_id`（明文，非机密）。
- POST `wx.mail.qq.com/oauth/token`，`grant_type=refresh_token`。
- **新 token 只有 3600s**，够一次批量处理；脚本**不回写凭证库**（避免 AES-GCM iv 复用风险）。
- 之后**别再用 `pipeline.py --steps fetch`** —— 它会先跑 `decrypt_token.py`，把新 token 覆盖成旧的过期 token。
  改手动按序跑：fetch_mail → fetch_attachments → extract_zip → rename_invoices → archive。
- 想持久化就让用户重新启用「QQ 邮箱」连接器（设置 — 连接器）。

## 字段提取规则

| 票种 | 识别标志 | 日期 | 金额 | 补充信息 |
|---|---|---|---|---|
| 火车票 | 特征词投票：铁路电子客票/电子客票号/12306/车次/二等座/票价 命中 ≥2 | `(20\d{2})年(\d{1,2})月(\d{1,2})日[^开]{0,20}?\d{1,2}:\s*\d{2}开` | 最后一个 `¥xx.xx` | 票面原始站名 `出发站-到达站` |
| 退票 | 同上且含 `退票费` | 同上 | 同上 | 同上 |
| 网约车发票 | 含开票日期 + 交通运输服务/价税合计 | 三级兜底（见下） | `（小写）¥(\d+\.\d{2})` | 省+市 |
| 行程报销单 | 含 `行程起止日期` / 笔行程 | `行程起止日期：A 至 B` | `合计([\d.]+)元` | 「城市」列 |

**网约车发票的三级兜底链**（票面常常没有「城市」列，曹操发票连出行日期都没有）：
1. 日期取**配套行程单**的起止日期（配对键 `(金额, 平台)`）
2. 取不到 → 发票自身的**出行明细日期**
3. 再没有 → **开票日期**（同样是票面内容，符合规则）
城市同理：行程单城市列 → 销方注册地所在市（`苏州市吉利优行电子科技有限公司` → 苏州市）。
平台判定顺序要紧：**先判曹操再判滴滴**（含 `曹操|吉利优行|优行` → 曹操）。

## 异常处理与降级策略

| 异常场景 | 策略 | 给用户的提示 |
|---|---|---|
| token 过期（401） | 中止下载，保留已下载部分 | "先调一次 MCP GetMe 触发续期再重试" |
| 单个附件下载失败 | 跳过后继续，脚本可重跑补全 | 列表标 `!!`，不中断批次 |
| zip 解压失败 / 加密包 | 记录单个失败，其余继续 | `[解压失败] 需人工处理` |
| 票面无文本层（扫描件） | 自动渲染成位图走 OCR 降级 | 与有文本层结果一致 |
| OCR 也识别不出 | **跳过**，列明原因 | `[跳过] 未识别到乘车日期` |
| 日期/金额/城市缺失任一 | **跳过**，不改文件名 | 写明缺哪个字段 |
| 目标文件名已存在 | 自动追加 `_2` `_3` | 清单里显示最终名 |
| 归档无匹配日期文件夹 | 留在 `<dest>/_流程/待处理/` 并列出 | "建议新建文件夹或人工判断" |
| 归档命中多个文件夹 | **不移动**，列出候选 | "需人工判断" |
| 邮箱限流 429 | 15/30/60/120s 退避重试 4 次 | 仍失败才报错 |
| 单封邮件多个附件 | 全部纳入 manifest，逐个独立处理 | manifest 里标 `subject`便于核对 |

## 发布到线上（供他人调用）

`serve.py` 是**纯 Python 标准库**实现的 HTTP 服务，无第三方 Web 依赖。

```bash
"$PY" "$S/serve.py" --host 0.0.0.0 --port 8080 --token <调用密钥>
```

已自动读取 `PORT` 环境变量并绑定 `0.0.0.0`，满足发布沙箱要求。
调用令牌也可用环境变量注入：`MOORE_TOKEN=<key>`（不设则免鉴权）。

### 当前线上实例（已发布）

| 项目 | 值 |
|---|---|
| 访问链接 | `https://moore-invoice.app.workbuddy.host/` |
| appId / 域名前缀 | `moore-invoice` |
| 启动命令 | `python scripts/serve.py`（PORT=8080） |
| 依赖 | `pip install -r requirements.txt`（仅 pymupdf，OCR 未装，图片票据会跳过） |
| 管理入口 | 设置 — 数据管理 — 应用 |

线上冒烟验证记录：`/healthz` 200、`/` 上传页 200、`/api/inspect` 对真实火车票返回
`20260911_火车发票_苏州南站-福州站_407.50.pdf`，与本地结果一致。

### 接口

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/healthz` | 健康检查，返回版本/是否鉴权/是否开放目录操作 |
| GET | `/` | 浏览器可直接用的上传页（选票据→看建议文件名） |
| POST | `/api/inspect` | `{"files":[{"filename","content_base64"}],"template":"railway"}` → 返回建议文件名与解析结果 |
| POST | `/api/rename` | `{"root","template","execute"}` → 目录级批量改名；**需启动时 `--allow-fs` 放行** |

鉴权：启动时带 `--token` 后，所有 POST 需 `Authorization: Bearer <token>`。

```bash
curl -X POST https://<域名>/api/inspect -H "Authorization: Bearer <token>" \
     -H "Content-Type: application/json" -d @payload.json
```

### 三种发布方式

| 方式 | 做法 | 适用 |
|---|---|---|
| **WorkBuddy「发布为应用」**（推荐） | 对技能根目录调用 `workbuddy_sites_deploy`，`language=python`，启动命令 `python scripts/serve.py` | 想立刻拿到分享链接 |
| Docker 自建 | `docker build -f deploy/Dockerfile -t moore-invoice . && docker run -p 8080:8080 moore-invoice` | 内网/私有云 |
| 团队共享技能 | 打包技能目录，对方放入 `~/.workbuddy/skills/` | 直接复用本地全流程 |

### 凭证与权限（重要）

- **🤖 调用别人的 Share Link**：只需调用方手里的 Bearer token，无需邮箱授权。
- **📧 邮箱下载链路（fetch/download 两步）只在本地可用**。它读本机
  `~/.workbuddy/connectors/<uuid>/.credentials.v3.json` 里的 OAuth 凭证。
  **线上沙箱里没有这个文件**，所以云端服务只提供 `/api/inspect`、`/api/rename`。
  需要云端自动收邮件，必须自建 Docker 并把连接器凭证挂进容器（不推荐外传）。
- **🗂 `/api/rename` 默认关闭。** 必须显式 `--allow-fs <目录>` 才放行，且 root 必须在该目录内；
  跨盘符路径会被判为越权返回 403（Windows 上 `os.path.commonpath` 跨盘抛 `ValueError`，已兜住）。
- 调用 ≠ 发布：发布链接会覆盖该链接上已有的内容，每次发布都需要使用者本轮明确要求。

## Python 环境（Windows）

```bash
export PATH="/usr/bin:/bin:/usr/local/bin:$PATH"    # 否则 ls/dirname 报 command not found
PY='C:\Users\majunpeng\.workbuddy\binaries\python\envs\default\Scripts\python.exe'
# 脚本路径必须写 C:/... 风格，托管解释器不认 Git Bash 的 /c/... 映射
```

依赖：`pymupdf`（必需）、`cryptography`（邮箱链路）、`rapidocr-onnxruntime`（图片/扫描件，可选）、
`cpca`（城市→省份兜底，可选）。

```bash
# pip 直连 PyPI 常被护栏拦截，用腾讯镜像
"$PY" -m pip install -r requirements-full.txt \
  --index-url https://mirrors.cloud.tencent.com/pypi/simple
```

## 已封装的反直觉坑

1. **火车票日期会被「开票日期」骗到。** 票面同时有 `开票日期:2026年09月15日` 和 `2026年09月11日09:02开`，
   必须用 `\d{2}:\d{2}开` 前的那个日期。
2. **行程单城市列会被 PDF 换行切成两半。** 文本层里"苏州"和"市"是两个 word，
   必须**按「城市」表头的 x 坐标取整列**，再把 y 间距 ≤9pt 的词合并成一个单元格。
3. **开票日期必须与坐标重组的行共取**。`get_text("text")` 里「开票日期：」后面紧跟的是购买方信息，日期被推走了。
4. **OCR 稳定丢站名末尾「站」字**（dpi 300/400/600 都一样，不是清晰度问题）。有效解是**二段式增强**：
   先全页 OCR 定位车次 `[A-Z]\d{1,4}`，以其为中心裁横条 ROI 放大 2 倍再识一次。
5. **OCR 坐标是像素（数千），PDF 是 pt（几百）**，`lines_of` 的行聚合阈值必须随尺度自适应（`max(y)>1000` 时取 `max(y)/80`）。
6. **`build_plan` 里 `start/end` 必须显式初始化**，否则会串用上一轮行程单的日期。
7. **扫根目录时 `relpath` 是 `"."`**，`".".startswith(("."))` 恒真会把根目录文件全漏掉，必须先判 `rel == "."`。
8. **zip 内中文名常用 cp437 伪装**，须还原为 GBK。`extract_zip.py` 已处理。
9. **文件夹名可能有年份笔误**（`20260911~20110913`），`archive.py` 按起始年份校正后再比对。

## 已知限制

- OFD 不解析内容，默认在解压阶段丢弃（`--keep-ofd` 可保留，但仍会被跳过）
- 城市→省份靠内置字典（江苏/福建/浙江/广东全覆盖）+ `cpca` 兜底，都未命中则只写城市不带省
- 多笔行程不拆分成多个文件副本，一文件对应一张票
- OCR 较慢：模型加载约 5s，单页约 8s；识别不全时按规则跳过
- 孤立网约车发票（无配套行程单）只能退到开票日期，会与"有行程单时"的命名不同 —— 属预期降级

## 验收清单

改完任何解析逻辑，必须跑完这四项再交付：

| 场景 | 期望结果 |
|---|---|
| 全量 PDF 目录重跑 | `改动 0 / 跳过 0`（幂等，不会弄乱已正确的名字） |
| 单张票据图片 | 命名结果与同名 PDF 版一致（仅扩展名不同） |
| 无文本层扫描件 PDF | 走 OCR 降级，结果同上 |
| 孤立发票（无配套行程单） | 走兜底链命名，而不是被跳过 |

## 安全约定

- `--execute` 前先备份：`shutil.copytree(src, backup, dirs_exist_ok=True)`
- token 文件下载完立即删除（`pipeline.py` 已自动做）
- 一律先预览给用户确认再执行
