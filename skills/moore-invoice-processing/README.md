# 摩尔发票处理 / Moore Invoice Processing

从腾讯（QQ）邮箱筛选并下载发票邮件附件，解析票面字段，按规范批量重命名并归档；
同时提供一个纯标准库的 HTTP 服务，可发布成在线链接供他人调用。

**在线服务（已发布）：https://moore-invoice.app.workbuddy.host/**
浏览器打开即可上传票据、查看建议文件名；程序调用走 `POST /api/inspect`。

## 目录结构

```
moore-invoice-processing/
├── SKILL.md                    技能主文档（工作流 / 参数 / 异常处理 / 发布方式）
├── requirements.txt            线上服务最小依赖（pymupdf）
├── requirements-full.txt       本地全量依赖（含邮箱链路 cryptography、OCR）
├── deploy/Dockerfile           自建容器部署
└── scripts/
    ├── pipeline.py             一键流水线（编排，推荐入口）
    ├── fetch_mail.py           邮件筛选 → manifest（REST 直连）
    ├── decrypt_token.py        解密本地连接器凭证取 token
    ├── fetch_attachments.py    按 manifest 批量下载附件落盘
    ├── extract_zip.py          解压 zip，只保留票证并摊平
    ├── rename_invoices.py      票面解析 + 重命名（核心）
    ├── archive.py              按业务日期归入差旅文件夹
    └── serve.py                HTTP 服务（发布上线）
```

## 快速开始

```bash
PY='C:\Users\majunpeng\.workbuddy\binaries\python\envs\default\Scripts\python.exe'
S=/c/Users/majunpeng/.workbuddy/skills/moore-invoice-processing/scripts

# 全流程预览
"$PY" "$S/pipeline.py" --dest "D:\00费用报销\未报销" --days 7 --template railway

# 确认后执行
"$PY" "$S/pipeline.py" --dest "D:\00费用报销\未报销" --days 7 --template railway --execute

# 启动在线服务
"$PY" "$S/serve.py" --token <调用密钥>
```

## 命名格式

```
20260911_火车发票_苏州南站-福州站_407.50.pdf      # railway（默认）
20260911_火车票_苏州南站-福州站_407.50.pdf        # standard
苏州南站-福州站-20260911-407.50.pdf              # legacy12306
```

字段一律取自票面内容，识别不出就跳过并在清单中写明原因。详见 `SKILL.md`。
