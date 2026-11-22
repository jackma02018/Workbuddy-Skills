#!/usr/bin/env python3
"""从 WorkBuddy 本地连接器凭证中解密 QQ 邮箱 access token。

仅在本人读取自己邮箱的场景下使用。产物写入 /tmp 下 0600 文件，**不打印 token**。

用法:
    python decrypt_token.py [--out /tmp/token] [--server qq-mail]

输出: 一行 "OK accessToken (serverKey=...) -> <path> (len=N)"
"""
import argparse
import base64
import hashlib
import hmac as _hmac
import json
import os
import sys

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
except ModuleNotFoundError:
    sys.exit("缺少 cryptography，请先: pip install cryptography")

HKDF_INFO = b"workbuddy-oauth-credentials-v1"
CONN_ROOT = os.path.expanduser("~/.workbuddy/connectors")


def hmac_digest(key: bytes, msg: bytes) -> bytes:
    return _hmac.new(key, msg, hashlib.sha256).digest()


def hkdf_sha256(ikm: bytes, salt: bytes, info: bytes, length: int) -> bytes:
    # RFC 5869: PRK = HMAC(key=salt, msg=ikm)
    prk = hmac_digest(salt, ikm)
    okm, block, counter = b"", b"", 1
    while len(okm) < length:
        block = hmac_digest(prk, block + info + bytes([counter]))
        okm += block
        counter += 1
    return okm[:length]


def compute_check(payload: bytes, salt: bytes) -> str:
    return base64.b64encode(hashlib.sha256(payload + salt).digest()[:16]).decode()


def find_user_dir() -> str:
    """连接器目录: <uuid>/ 为已登录账号, default/ 为未登录占位。"""
    cands = [
        d for d in sorted(os.listdir(CONN_ROOT))
        if os.path.isdir(os.path.join(CONN_ROOT, d))
        and d != "default"
        and os.path.exists(os.path.join(CONN_ROOT, d, ".credentials.v3.json"))
    ]
    if not cands:
        sys.exit("未找到已登录账号的连接器凭证目录")
    return os.path.join(CONN_ROOT, cands[0])


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="/tmp/.qqmail_token")
    ap.add_argument("--server", default="qq-mail")
    ap.add_argument("--conn-dir", help="指定连接器目录（默认自动探测）")
    args = ap.parse_args()

    conn_dir = args.conn_dir or find_user_dir()
    user_id = os.path.basename(conn_dir.rstrip("/"))
    cred = json.load(open(os.path.join(conn_dir, ".credentials.v3.json")))
    header = cred["encryption"]
    salt = base64.b64decode(header["salt"])

    with open(os.path.join(conn_dir, ".master.key"), "rb") as fh:
        master_key = fh.read()
    if len(master_key) != 32:
        sys.exit("master.key 长度异常（应为 32 字节）")

    if compute_check(user_id.encode(), salt) != header["userIdCheck"]:
        sys.exit("userIdCheck 不匹配")
    if compute_check(master_key, salt) != header["keyCheck"]:
        sys.exit("keyCheck 不匹配")

    aes_key = hkdf_sha256(master_key + user_id.encode(), salt, HKDF_INFO, 32)

    tried = []
    for server_key, entry in cred.get("mcpOAuth", {}).items():
        if not server_key.split("|")[0] == args.server:
            continue
        for field, env in entry.items():
            if not isinstance(env, dict) or "ct" not in env:
                continue
            aad = f"{user_id}|{server_key}|{field}".encode()
            try:
                plain = AESGCM(aes_key).decrypt(
                    base64.b64decode(env["iv"]),
                    base64.b64decode(env["ct"]) + base64.b64decode(env["tag"]),
                    aad,
                )
            except Exception as exc:
                tried.append(f"{server_key}/{field}: {type(exc).__name__}")
                continue
            out = args.out if field == "accessToken" else f"{args.out}.{field}"
            with open(out, "w") as fh:
                fh.write(plain.decode())
            os.chmod(out, 0o600)
            print(f"OK {field} (serverKey={server_key}) -> {out} (len={len(plain)})")
            return 0

    sys.exit("解密失败，尝试过: " + "; ".join(tried))


if __name__ == "__main__":
    sys.exit(main())
