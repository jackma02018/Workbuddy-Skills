---
name: ima-knowledge-archive
description: 把本地文件（Markdown/PPT/Word/PDF 等）归档进 ima 个人知识库，并按主题归类到指定文件夹、统一命名与标签。ima MCP 没有"直接传文本"的入口，必须走 create_media → COS PUT 上传 → add_knowledge 三步。当用户说"存到 ima 知识库""归档到 ima""把这次的资料入库""下次能检索到"时使用。
agent_created: true
---

# ima 知识库归档

## 核心结论（先看）

ima MCP **没有文本入库接口**：`add_knowledge` 只接受 `media_id`，必须先拿到文件实体。固定三步：

```
create_media(...)  →  返回 media_id + cos_credential（临时密钥）
        ↓
PUT 文件到 COS（自写签名上传，纯 stdlib）
        ↓
add_knowledge(knowledge_base_id, media_id, folder_id)  → 正式登记
```

`search_knowledge` 只能读、不能写；`fetch_media_content` 只能取解析后的文本。

## 一、准备：定位知识库与文件夹

1. `get_knowledge_base_list` → 取个人知识库 id（本机常用：`001a44a0c3805fba`，Jack 的知识库）
2. `get_knowledge_list(knowledge_base_id, limit=40, cursor="")` → 列出根层，**文件夹在返回项里带 `folder_info.folder_id`**
   - 坑：`filters` 用 `MEDIA_TYPE_FILTER_TYPE: ["FOLDER"]` 返回空，别用；直接不带 filter 列。
   - 坑：`cursor` 首次传空字符串，翻页用返回的 `next_cursor`；`is_end=true` 到底。
3. 已知目录（智慧工厂 `folder_7393820341719354` 下）：

| 目录 | folder_id | 用途 |
|---|---|---|
| 07客户_组装行业 | `folder_7393992274617939` | 客户方案/调研/交付件 |
| 00方案素材 | `folder_7490704263620515` | 79 份通用素材（全文档，无位图） |
| 摩尔方案工具 | `folder_7393822891858379` | 方法论、脚本、素材地图、索引 |
| 摩尔效率提升 | `folder_7393825815289697` | 环境手册、效率类 |

## 二、Step 1 — create_media

参数：`knowledge_base_id`、`file_name`（含扩展名）、`file_ext`、`file_size`（**必须与实际字节数一致**）、`content_type`。

content_type 对照（错填会被拒，禁止 `application/octet-stream` 兜底）：
`md/markdown→text/markdown`；`txt→text/plain`；`pdf→application/pdf`；`docx→application/vnd.openxmlformats-officedocument.wordprocessingml.document`；`pptx→application/vnd.openxmlformats-officedocument.presentationml.presentation`；`xlsx→application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`；`png→image/png`。

小技巧：大脚本（.py）若想入库，复制成 `.md` 后按 markdown 上传，标题与内容仍可全文检索。

## 三、Step 2 — 上传到 COS

把 create_media 返回的 `{file, media_id, cos_credential}` 逐条写入一个 JSON 数组，然后：

```bash
python scripts/ima_upload.py <凭据JSON路径>
# 输出： 200  filename.md  media_id    （非 200 会打印响应体前 400 字符）
```

脚本要点（已验证可用，纯 stdlib）：
- host = `<bucket_name>.cos.<region>.myqcloud.com`，URL = `https://<host>/<cos_key>`
- 签名：`sha1\n<keyTime>\n<sha1(httpString)>\n`，`httpString = put\n/<cos_key>\n\n<httpHeaders>\n`
- `q-header-list` 必须小写字典序；`Authorization` 用 `q-sign-algorithm=sha1&q-ak=...&q-sign-time=...&q-key-time=...&q-header-list=...&q-url-param-list=&q-signature=...`
- 必须带 `x-cos-security-token`（凭证里的 `token`）
- 凭证有效期约 12 小时，超时重新 create_media

## 四、Step 3 — add_knowledge

`add_knowledge(knowledge_base_id, media_id, folder_id, duplicate_name_strategy="DUPLICATE_NAME_STRATEGY_SAVE")`

返回 `{"media_id": ...}` 即成功。校验：`get_knowledge_list(knowledge_base_id, folder_id, filters=[MEDIA_TYPE_FILTER_TYPE])`，看 `parse_progress`（100 = 解析完成）、`media_state`（1=解析中，2=完成）。解析中的条目暂时检索不到，属正常。

## 五、命名与标签规则

- 命名：`主题-对象-日期-版本`，例 `MES方案-绿控传动生产管理专题PPT-20260917-v3`
- 版本：v1 初版 / v2 结构改动 / v3 内容或视觉增强；同主题只留最新与上一版
- 标签维度：文档类型（方案/调研/方法论/工具/素材地图/环境）+ 客户名 + 业务域 + 可复用性
- 归类：客户相关 → 客户文件夹；可复用方法论/工具 → 摩尔方案工具；环境效率 → 摩尔效率提升

## 六、归档前先做去重与剔除

建议剔除（避免污染检索）：中间版本文件、被新版取代的脚本、含密钥的临时文件（如 `_creds.json`）、与客户原始件重复的内容。剔除要在交付清单里写明理由。

## 七、收尾

产出一份 `索引-ima入库保存结果清单-YYYYMMDD-vN.md`（标题/分类/标签/一句话摘要 + 未入库理由 + 检索模板），并把它自己也入库到「摩尔方案工具」，下次按索引直取。

检索提示：`search_knowledge` 单次约 12 万字符会被截断并落盘到沙箱外路径（读不到），大范围查找一律改用 `get_knowledge_list` 翻页。
