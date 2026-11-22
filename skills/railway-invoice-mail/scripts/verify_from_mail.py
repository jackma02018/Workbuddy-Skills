#!/usr/bin/env python3
"""从 12306 发票通知邮件正文中提取清单，用于与 PDF 解析结果交叉校验。

邮件正文是 HTML，必须先剥标签再匹配。字段: 发票号码 乘车日期 车次 发到站 项目 金额

用法:
    python verify_from_mail.py msg_xxx [msg_yyy ...] [--token /tmp/.qqmail_token]

输出: 逐行清单 + 张数 + 合计金额。
"""
import argparse
import json
import re
import sys
import time
import urllib.request

BASE = "https://api.mail.qq.com"

ROW = re.compile(
    r"(\d{20})\s+(\d{4}年\d{1,2}月\d{1,2}日)\s+([GDCKZ]\d+)\s+"
    r"([\u4e00-\u9fa5]+)-([\u4e00-\u9fa5]+)\s+(\S+?)\s+([\d.]+)"
)


def get_message(mid: str, alias: str, token: str) -> str:
    req = urllib.request.Request(
        f"{BASE}/v1/aliases/{alias}/messages/{mid}",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.loads(resp.read())
    return (data.get("data") or {}).get("body", "") or ""


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", html)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", text)


def resolve_alias(token: str) -> str:
    req = urllib.request.Request(
        BASE + "/v1/me",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        me = json.loads(resp.read())
    for alias in me["data"]["aliases"]:
        if alias.get("is_primary"):
            return alias["alias_id"]
    return me["data"]["aliases"][0]["alias_id"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("message_ids", nargs="+")
    ap.add_argument("--token", default="/tmp/.qqmail_token")
    ap.add_argument("--alias")
    args = ap.parse_args()

    token = open(args.token).read().strip()
    alias = args.alias or resolve_alias(token)

    rows = []
    for i, mid in enumerate(args.message_ids):
        if i:
            time.sleep(7)  # 限流 10 次/分钟
        rows.extend(ROW.finditer(strip_html(get_message(mid, alias, token))))

    rows.sort(key=lambda m: m.group(2))
    print(f"共 {len(rows)} 条")
    for m in rows:
        print(f"{m.group(1)}  {m.group(2)}  {m.group(3):<6} "
              f"{m.group(4)}-{m.group(5)}  {m.group(6)}  {m.group(7)}")
    print(f"合计金额: {sum(float(m.group(7)) for m in rows):.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
