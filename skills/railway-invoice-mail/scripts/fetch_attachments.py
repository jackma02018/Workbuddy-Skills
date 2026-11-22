#!/usr/bin/env python3
"""直接调用 QQ 邮箱 REST API 批量下载附件到本地（避免 base64 经过上下文）。

为什么不用 MCP 的 DownloadAttachment: 它把整个附件以 base64 塞进工具返回值，
单个 42KB 附件约 1.4 万 token，20 个附件就能打爆上下文。本脚本直接落盘。

用法:
    python decrypt_token.py --out /tmp/.qqmail_token
    python fetch_attachments.py --manifest attachments.json --out ./原始附件

manifest 为 JSON 数组，每项:
    {"message_id": "msg_xxx", "attachment_id": "att_xxx", "filename": "xxx.zip", "size": 42384}

限流: QQ 邮箱 API 10 次/分钟。脚本内置 6.5s 最小间隔 + 429 指数退避重试。
已存在且大小正常的文件会跳过，可安全重复运行（断点续传）。
"""
import argparse
import base64
import json
import os
import sys
import time
import urllib.error
import urllib.request

BASE = "https://api.mail.qq.com"
MIN_INTERVAL = 6.5


def call(path: str, token: str, timeout: int = 180) -> bytes:
    req = urllib.request.Request(
        BASE + path,
        headers={
            "Authorization": f"Bearer {token}",
            "User-Agent": "workbuddy-local/1.0",
            "Accept": "application/json, */*",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def resolve_alias(token: str) -> str:
    me = json.loads(call("/v1/me", token))
    for alias in me["data"]["aliases"]:
        if alias.get("is_primary"):
            return alias["alias_id"]
    return me["data"]["aliases"][0]["alias_id"]


def decode_attachment(raw: bytes) -> bytes:
    """响应可能是裸字节，也可能是 {content: base64} / {data: {content: base64}}。"""
    try:
        payload = json.loads(raw)
    except Exception:
        return raw
    for node in (payload, payload.get("data") if isinstance(payload.get("data"), dict) else None):
        if not isinstance(node, dict):
            continue
        for key in ("content", "data", "base64"):
            if isinstance(node.get(key), str):
                return base64.b64decode(node[key])
    raise RuntimeError(f"无法解析响应: {raw[:120]}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--manifest", required=True, help="附件清单 JSON 文件")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--token", default="/tmp/.qqmail_token")
    ap.add_argument("--alias", help="alias_id，默认从 /v1/me 自动获取")
    args = ap.parse_args()

    token = open(args.token).read().strip()
    alias = args.alias or resolve_alias(token)
    items = json.load(open(args.manifest))
    os.makedirs(args.out, exist_ok=True)

    pending = [
        it for it in items
        if not os.path.exists(os.path.join(args.out, it["filename"]))
        or os.path.getsize(os.path.join(args.out, it["filename"])) < 1000
    ]
    print(f"共 {len(items)} 个，待下载 {len(pending)} 个（已存在 {len(items) - len(pending)} 个）")

    ok, bad, last = len(items) - len(pending), [], 0.0
    for idx, it in enumerate(pending, 1):
        blob = None
        for attempt in range(5):
            wait = MIN_INTERVAL - (time.time() - last)
            if wait > 0:
                time.sleep(wait)
            try:
                last = time.time()
                blob = decode_attachment(call(
                    f"/v1/aliases/{alias}/messages/{it['message_id']}"
                    f"/attachments/{it['attachment_id']}", token))
                break
            except urllib.error.HTTPError as exc:
                if exc.code == 429:
                    time.sleep(15 * (attempt + 1))
                    continue
                bad.append(f"{it['filename']}: HTTP {exc.code}")
                break
            except Exception as exc:
                bad.append(f"{it['filename']}: {exc}")
                break
        if blob is None:
            print(f"[{idx}/{len(pending)}] 跳过 {it['filename']}")
            continue
        with open(os.path.join(args.out, it["filename"]), "wb") as fh:
            fh.write(blob)
        print(f"[{idx}/{len(pending)}] {'OK ' if len(blob) > 1000 else '?? '}"
              f"{it['filename']}  {len(blob)} B")
        ok += 1

    print(f"\n完成: {ok}/{len(items)}")
    for b in bad:
        print("  失败:", b)
    return 0 if ok == len(items) else 1


if __name__ == "__main__":
    sys.exit(main())
