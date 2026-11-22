#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按时间/关键词筛选 QQ 邮箱发票邮件，列出附件并生成下载清单 manifest。

纯 REST 直连，只产出元数据 JSON，**附件内容绝不进上下文**。

用法:
    python fetch_mail.py --after 2026-09-15 --subject 电子发票 --manifest manifest.json
    python fetch_mail.py --after 2026-09-01T00:00+08:00 --sender didifapiao --ext zip,pdf

输出 manifest 为 JSON 数组，每项:
    {"message_id","attachment_id","filename","size","subject","from","created_at"}

------------------------------------------------------------------
接口事实（2026-09 实测于 api.mail.qq.com，写脚本前务必知道这四条）:

1. GET /v1/aliases/{alias}/messages?dir=inbox&limit=N
   -> {"data":[{message_id,dir,from,to,subject,snippet,is_read,has_attachments,created_at}],
       "pagination":{"next_cursor","has_more"}}
   **limit 上限是 50**，传 100 直接 HTTP 400。

2. `after=<RFC3339 UTC>` 有效，由服务端过滤（实测总数 18 -> after 后 9）。
   `before=` 同理。时间必须是 UTC 且以 Z 结尾。

3. **`q=` 关键词参数无效**，会被静默忽略（传 q=12306 返回的结果与不传完全一致）。
   所以关键词筛选必须在**客户端**做，本脚本用 --sender/--subject/--keyword 过滤。

4. GET /v1/aliases/{alias}/messages/{mid}/attachments
   -> {"data":[{attachment_id,filename,content_type,size}]}

限流 10 次/分钟，脚本内置 6.5s 最小间隔 + 429 指数退避。
"""
import argparse
import datetime
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://api.mail.qq.com"
MIN_INTERVAL = 6.5
MAX_LIMIT = 50

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


# ------------------------------------------------------------------ HTTP
def call(path, token, timeout=120):
    req = urllib.request.Request(
        BASE + path,
        headers={"Authorization": "Bearer " + token,
                 "Accept": "application/json, */*",
                 "User-Agent": "moore-invoice-processing/1.0"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
    return json.loads(raw) if raw else {}


def call_safe(path, token, _last=[0.0]):
    """带限流与重试的请求。返回 (data, error)。"""
    delay = 60
    for attempt in range(4):
        wait = MIN_INTERVAL - (time.time() - _last[0])
        if wait > 0:
            time.sleep(wait)
        try:
            _last[0] = time.time()
            return call(path, token), None
        except urllib.error.HTTPError as exc:
            body = ""
            try:
                body = exc.read()[:200].decode("utf-8", "replace")
            except Exception:
                pass
            if exc.code == 401:
                return None, ("HTTP 401 token 已过期：先调用一次 MCP 工具（如 GetMe）触发连接器续期，"
                              "再重跑 decrypt_token.py。 response=%s" % body)
            if exc.code == 429:
                time.sleep(delay)
                delay *= 2
                continue
            return None, "HTTP %s %s" % (exc.code, body)
        except Exception as exc:
            return None, str(exc)
    return None, "重试多次仍失败（持续限流）"


def resolve_alias(token):
    data, err = call_safe("/v1/me", token)
    if err:
        sys.exit("获取 alias 失败: " + err)
    for alias in data["data"]["aliases"]:
        if alias.get("is_primary"):
            return alias["alias_id"], data["data"]["aliases"][0]["email"]
    return data["data"]["aliases"][0]["alias_id"], data["data"]["aliases"][0]["email"]


# ------------------------------------------------------------------ 工具
def to_utc(s, assume_tz="+08:00"):
    """'2026-09-15' / '2026-09-15T14:00' / '2026-09-15T00:00+08:00' -> '2026-09-15T00:00:00Z'"""
    if not s:
        return None
    t = s.strip()
    if "T" not in t:
        t += "T00:00:00"
    elif len(t) == 16:
        t += ":00"
    if not re.search(r"(Z|[+-]\d{2}:?\d{2})$", t):
        t += assume_tz
    dt = datetime.datetime.fromisoformat(t.replace("Z", "+00:00"))
    return dt.astimezone(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def matches(text, pattern):
    if not pattern:
        return True
    try:
        return re.search(pattern, text or "", re.I) is not None
    except re.error:
        return pattern.lower() in (text or "").lower()


def msg_matches(m, args):
    sender = "%s %s" % (m.get("from", {}).get("email", ""), m.get("from", {}).get("name", ""))
    haystack = " ".join([m.get("subject", ""), m.get("snippet", ""), sender])
    return (matches(sender, args.sender)
            and matches(m.get("subject", ""), args.subject)
            and matches(haystack, args.keyword))


def ext_ok(name, allowed):
    if not allowed:
        return True
    ext = os.path.splitext(name)[1].lower().lstrip(".")
    return ext in {e.lower().lstrip(".") for e in allowed}


# ------------------------------------------------------------------ 主流程
def main():
    ap = argparse.ArgumentParser(description="筛选发票邮件并生成附件下载清单")
    ap.add_argument("--token", default=os.path.join(os.path.expanduser("~"), ".qqmail_token"))
    ap.add_argument("--alias", help="alias_id，省略则自动从 /v1/me 获取")
    ap.add_argument("--dir", default="inbox")
    ap.add_argument("--after", help="起始时间，如 2026-09-15 或 2026-09-15T00:00+08:00")
    ap.add_argument("--before", help="截止时间，格式同 --after")
    ap.add_argument("--sender", help="发件人邮箱/名称，支持正则")
    ap.add_argument("--subject", help="主题关键词，支持正则")
    ap.add_argument("--keyword", help="主题+摘要+发件人全文关键词，支持正则")
    ap.add_argument("--ext", default="", help="只看这些扩展名的附件，逗号分隔，如 zip,pdf")
    ap.add_argument("--skip-ofd", action="store_true", help="排除 .ofd 附件（默认只保留 PDF 时用）")
    ap.add_argument("--max-pages", type=int, default=10, help="最多翻几页，默认 10")
    ap.add_argument("--manifest", required=True, help="输出清单 JSON 路径")
    ap.add_argument("--list-only", action="store_true", help="只打印邮件摘要，不查附件")
    args = ap.parse_args()

    token = open(args.token).read().strip()
    alias, email = (args.alias, "") if args.alias else resolve_alias(token)
    print("账号: %s  alias: %s" % (email or "(未提供)", alias))

    allowed = [e.strip() for e in args.ext.split(",") if e.strip()]

    # ---- 1. 拉邮件列表（分页）
    messages, cursor, errors = [], None, []
    for page in range(args.max_pages):
        q = {"dir": args.dir, "limit": MAX_LIMIT}
        a, b = to_utc(args.after), to_utc(args.before)
        if a:
            q["after"] = a
        if b:
            q["before"] = b
        if cursor:
            q["cursor"] = cursor
        path = "/v1/aliases/%s/messages?%s" % (alias, urllib.parse.urlencode(q))
        data, err = call_safe(path, token)
        if err:
            errors.append(err)
            break
        batch = data.get("data", [])
        messages += batch
        pg = data.get("pagination", {})
        print("  第%d页 收到 %d 封 (累计 %d)" % (page + 1, len(batch), len(messages)))
        if not pg.get("has_more") or not pg.get("next_cursor"):
            break
        cursor = pg["next_cursor"]
    if errors:
        sys.exit("拉取邮件失败: " + errors[0])

    # ---- 2. 客户端关键词过滤（服务端 q= 无效）
    hits = [m for m in messages if msg_matches(m, args)]
    print("\n命中 %d / %d 封邮件" % (len(hits), len(messages)))
    for m in hits:
        print("  - [%s] %-28s %s" % (m.get("created_at", "")[:10],
                                     (m.get("from", {}) or {}).get("email", "")[:28],
                                     m.get("subject", "")[:40]))

    if args.list_only:
        return

    # ---- 3. 逐封查附件
    items = []
    for idx, m in enumerate(hits, 1):
        if not m.get("has_attachments"):
            continue
        path = "/v1/aliases/%s/messages/%s/attachments" % (alias, m["message_id"])
        data, err = call_safe(path, token)
        if err:
            errors.append("%s: %s" % (m["message_id"], err))
            if "401" in err:
                break
            continue
        group = []
        for att in data.get("data", []):
            fn = att.get("filename", "")
            if args.skip_ofd and fn.lower().endswith(".ofd"):
                continue
            if not ext_ok(fn, allowed):
                continue
            group.append({
                "message_id": m["message_id"],
                "attachment_id": att.get("attachment_id"),
                "filename": fn,
                "size": att.get("size"),
                "subject": m.get("subject", ""),
                "from": (m.get("from", {}) or {}).get("email", ""),
                "created_at": m.get("created_at", ""),
            })
        print("  [%d/%d] %s -> %d 个附件" % (idx, len(hits), m.get("subject", "")[:30], len(group)))
        items += group

    os.makedirs(os.path.dirname(os.path.abspath(args.manifest)) or ".", exist_ok=True)
    with open(args.manifest, "w", encoding="utf-8") as fh:
        json.dump(items, fh, ensure_ascii=False, indent=2)
    print("\n清单已写出: %s  共 %d 个附件" % (args.manifest, len(items)))
    if args.skip_ofd:
        print("已按 --skip-ofd 排除 ofd 文件")
    if errors:
        print("\n警告 %d 条:" % len(errors))
        for e in errors[:10]:
            print("   -", e)


if __name__ == "__main__":
    main()
