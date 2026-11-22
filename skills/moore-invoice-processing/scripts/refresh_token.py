#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用本地凭证里的 refresh_token 向官方 OAuth 端点续期 QQ 邮箱 access token。

只在本次会话使用：新 token 写到 --out 指定的临时文件，不回写凭证库。
"""
import argparse
import base64
import hashlib
import hmac as _hmac
import json
import os
import sys
import urllib.parse
import urllib.request

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HKDF_INFO = b"workbuddy-oauth-credentials-v1"
CONN_ROOT = os.path.expanduser("~/.workbuddy/connectors")
TOKEN_URL = "https://wx.mail.qq.com/oauth/token"
SERVER_KEY = "qq-mail|8e55080c6c967571"


def hmac_digest(key, msg):
    return _hmac.new(key, msg, hashlib.sha256).digest()


def hkdf_sha256(ikm, salt, info, length):
    prk = hmac_digest(salt, ikm)
    okm, block, counter = b"", b"", 1
    while len(okm) < length:
        block = hmac_digest(prk, block + info + bytes([counter]))
        okm += block
        counter += 1
    return okm[:length]


def compute_check(payload, salt):
    return base64.b64encode(hashlib.sha256(payload + salt).digest()[:16]).decode()


def find_user_dir():
    cands = [d for d in sorted(os.listdir(CONN_ROOT))
             if os.path.isdir(os.path.join(CONN_ROOT, d)) and d != "default"
             and os.path.exists(os.path.join(CONN_ROOT, d, ".credentials.v3.json"))]
    if not cands:
        sys.exit("未找到已登录账号的连接器凭证目录")
    return os.path.join(CONN_ROOT, cands[0])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    conn_dir = find_user_dir()
    user_id = os.path.basename(conn_dir.rstrip("/"))
    cred = json.load(open(os.path.join(conn_dir, ".credentials.v3.json"), encoding="utf-8"))
    header = cred["encryption"]
    salt = base64.b64decode(header["salt"])
    master_key = open(os.path.join(conn_dir, ".master.key"), "rb").read()
    aes_key = hkdf_sha256(master_key + user_id.encode(), salt, HKDF_INFO, 32)

    entry = cred["mcpOAuth"][SERVER_KEY]
    client_id = cred["mcpClientInfo"][SERVER_KEY]["client_id"]

    def dec(field):
        env = entry[field]
        aad = ("%s|%s|%s" % (user_id, SERVER_KEY, field)).encode()
        return AESGCM(aes_key).decrypt(base64.b64decode(env["iv"]),
                                       base64.b64decode(env["ct"]) + base64.b64decode(env["tag"]),
                                       aad).decode()

    refresh_token = dec("refreshToken")
    print("client_id=%s refresh_token_len=%d" % (client_id, len(refresh_token)))

    body = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "scope": entry.get("scope", "alias:read mail:read"),
    }).encode()
    req = urllib.request.Request(TOKEN_URL, data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded",
                                          "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read() or b"{}")
    except Exception as exc:
        detail = ""
        if hasattr(exc, "read"):
            try:
                detail = exc.read()[:300].decode("utf-8", "replace")
            except Exception:
                pass
        sys.exit("续期失败: %s %s" % (exc, detail))

    token = data.get("access_token")
    if not token:
        sys.exit("响应中没有 access_token: %s" % json.dumps(data, ensure_ascii=False)[:300])
    with open(args.out, "w") as fh:
        fh.write(token)
    os.chmod(args.out, 0o600)
    print("OK 新 access_token len=%d expires_in=%s -> %s"
          % (len(token), data.get("expires_in"), args.out))


if __name__ == "__main__":
    main()
