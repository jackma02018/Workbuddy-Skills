#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""摩尔发票处理 —— HTTP 服务（纯标准库，无第三方 Web 依赖）。

启动:
    python serve.py --host 0.0.0.0 --port 8080
    python serve.py --token <调用密钥>            # 开启 Bearer 鉴权
    python serve.py --allow-fs "D:\\00费用报销"    # 放行目录级批量改名接口

接口:
    GET  /healthz                 健康检查
    GET  /                        单页上传页（浏览器可直接用）
    POST /api/inspect             上传票据 -> 返回解析字段与建议文件名
                                  body: {"files":[{"filename":"a.pdf","content_base64":"..."}],
                                         "template":"railway"}
    POST /api/rename              目录级批量改名（需 --allow-fs 放行且 root 在其内）
                                  body: {"root":"D:\\...","template":"railway","execute":false}

说明: /api/inspect 把上传文件写入临时目录后复用 rename_invoices.build_plan，
      因此单票据与多票据（含行程单配对）走的是完全相同的一套解析逻辑。
"""
import argparse
import base64
import importlib.util
import json
import os
import re
import shutil
import sys
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
VERSION = "1.0.0"

_spec = importlib.util.spec_from_file_location("rename_invoices",
                                               os.path.join(HERE, "rename_invoices.py"))
ri = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ri)

AUTH_TOKEN = None
ALLOWED_ROOTS = []

INDEX_HTML = """<!doctype html><meta charset="utf-8"><title>摩尔发票处理</title>
<style>
body{font-family:-apple-system,"Microsoft YaHei",sans-serif;max-width:820px;margin:40px auto;padding:0 20px;color:#1a1a1a;background:#fff}
h1{font-size:22px} .row{margin:14px 0}
select,input,button{font-size:14px;padding:8px 10px}
button{background:#0a6;color:#fff;border:0;border-radius:6px;cursor:pointer}
pre{background:#f6f8fa;padding:14px;border-radius:8px;overflow:auto;font-size:13px}
</style>
<h1>摩尔发票处理 · 票据识别与命名</h1>
<div class="row"><input type="file" id="f" multiple accept=".pdf,.ofd,.jpg,.jpeg,.png,.webp"></div>
<div class="row">命名口径
<select id="t"><option value="railway">年月日_火车发票_出发站_到达站_金额</option>
<option value="standard">日期_类型_补充_金额</option>
<option value="legacy12306">出发站-到达站-日期-金额</option></select>
<button onclick="go()">解析</button></div>
<pre id="o">选择票据后点击解析。</pre>
<script>
const b64 = f => new Promise(r=>{const x=new FileReader();x.onload=()=>r(x.result.split(',')[1]);x.readAsDataURL(f)});
async function go(){
  const fs=document.getElementById('f').files; if(!fs.length){o.textContent='请先选择文件';return}
  o.textContent='解析中…';
  const files=[]; for(const f of fs) files.push({filename:f.name, content_base64:await b64(f)});
  const r=await fetch('/api/inspect',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({files, template:document.getElementById('t').value})});
  o.textContent=JSON.stringify(await r.json(),null,2);
}
</script>"""


def _safe_name(name):
    name = os.path.basename(name or "")
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fa5._-]", "_", name)[:120] or "upload.pdf"


def same_branch(child, parent):
    """child 是否位于 parent 之内。跨盘符（C: vs D:）时 commonpath 会抛 ValueError，必须兜住。"""
    try:
        return os.path.commonpath([os.path.normcase(child),
                                   os.path.normcase(parent)]) == os.path.normcase(parent)
    except ValueError:                      # 不同盘符 -> 必然不在其内
        return False


def do_inspect(payload):
    files = payload.get("files") or []
    if not files:
        raise ValueError("files 为空")
    template = payload.get("template", "railway")
    sep = payload.get("station_sep", "-")
    tmp = tempfile.mkdtemp(prefix="moore_invoice_")
    try:
        for item in files:
            name = _safe_name(item.get("filename"))
            blob = base64.b64decode(item.get("content_base64", ""))
            if not blob:
                continue
            with open(os.path.join(tmp, name), "wb") as fh:
                fh.write(blob)
        results = ri.build_plan(tmp, template, sep)
        ri.resolve_conflicts(results)
        out = []
        for r in results:
            out.append({
                "old": r["old"],
                "new": r.get("new"),
                "type": r.get("type"),
                "ok": bool(r.get("new")),
                "skip_reason": r.get("skip_reason") or None,
            })
        return {"ok": True, "template": template, "version": VERSION, "results": out}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def do_rename(payload):
    if not ALLOWED_ROOTS:
        raise PermissionError("服务端未开启目录操作（启动时需传 --allow-fs）")
    root = os.path.abspath(payload.get("root", ""))
    if not root or not os.path.isdir(root):
        raise ValueError("root 不是有效目录: %s" % root)
    allowed = [os.path.abspath(p) for p in ALLOWED_ROOTS]
    if not any(same_branch(root, a) for a in allowed):
        raise PermissionError("root 不在放行目录内，当前允许: %s" % allowed)
    template = payload.get("template", "railway")
    results = ri.build_plan(root, template, payload.get("station_sep", "-"))
    ri.resolve_conflicts(results)
    execute = bool(payload.get("execute"))
    changed = failed = 0
    if execute:
        for r in results:
            if not r.get("new"):
                continue
            src, dst = r["path"], os.path.join(os.path.dirname(r["path"]), r["new"])
            try:
                if os.path.abspath(src) != os.path.abspath(dst):
                    if os.path.exists(dst):
                        raise FileExistsError(dst)
                    os.rename(src, dst)
                changed += 1
            except Exception as exc:
                failed += 1
                r["skip_reason"] = "改名失败: %s" % exc
    return {
        "ok": True, "executed": execute, "changed": changed, "failed": failed,
        "results": [{"old": r["old"], "new": r.get("new"),
                     "type": r.get("type"),
                     "skip_reason": r.get("skip_reason") or None} for r in results],
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "MooreInvoice/" + VERSION
    protocol_version = "HTTP/1.1"

    # ---- helpers
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _authorized(self):
        if not AUTH_TOKEN:
            return True
        header = self.headers.get("Authorization", "")
        return header.strip() == "Bearer " + AUTH_TOKEN

    def log_message(self, fmt, *args):
        sys.stderr.write("[%s] %s\n" % (time_str(), fmt % args))

    # ---- routes
    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            return self._send(200, INDEX_HTML, "text/html; charset=utf-8")
        if path == "/healthz":
            return self._send(200, json.dumps(
                {"ok": True, "service": "摩尔发票处理", "version": VERSION,
                 "auth_required": bool(AUTH_TOKEN),
                 "fs_enabled": bool(ALLOWED_ROOTS)}, ensure_ascii=False))
        self._send(404, json.dumps({"ok": False, "error": "not found"}, ensure_ascii=False))

    def do_POST(self):
        path = self.path.split("?")[0]
        if not self._authorized():
            return self._send(401, json.dumps(
                {"ok": False, "error": "unauthorized，需要 Authorization: Bearer <token>"},
                ensure_ascii=False))
        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            payload = json.loads(raw.decode("utf-8")) if raw else {}
        except Exception as exc:
            return self._send(400, json.dumps({"ok": False, "error": "JSON 解析失败: %s" % exc},
                                              ensure_ascii=False))
        try:
            if path == "/api/inspect":
                return self._send(200, json.dumps(do_inspect(payload), ensure_ascii=False))
            if path == "/api/rename":
                return self._send(200, json.dumps(do_rename(payload), ensure_ascii=False))
            return self._send(404, json.dumps({"ok": False, "error": "not found"},
                                              ensure_ascii=False))
        except PermissionError as exc:
            return self._send(403, json.dumps({"ok": False, "error": str(exc)},
                                              ensure_ascii=False))
        except Exception as exc:
            import traceback
            traceback.print_exc()
            return self._send(500, json.dumps({"ok": False, "error": "%s: %s" % (type(exc).__name__, exc)},
                                              ensure_ascii=False))


def time_str():
    import time as _t
    return _t.strftime("%Y-%m-%d %H:%M:%S")


def main():
    global AUTH_TOKEN, ALLOWED_ROOTS
    ap = argparse.ArgumentParser(description="摩尔发票处理 HTTP 服务")
    ap.add_argument("--host", default=os.environ.get("HOST", "0.0.0.0"))
    # 发布到线上时必须监听 PORT 环境变量并绑定 0.0.0.0（WorkBuddy 发布为应用 / Docker 通用约定）
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT", 8080)))
    ap.add_argument("--token", default=os.environ.get("MOORE_TOKEN"),
                    help="调用鉴权令牌；也可由环境变量 MOORE_TOKEN 提供。"
                         "留空则无需鉴权（适合内网或短期分享）")
    ap.add_argument("--allow-fs", action="append", default=[],
                    help="放行 /api/rename 的目录，可重复传多次")
    args = ap.parse_args()

    AUTH_TOKEN = args.token
    ALLOWED_ROOTS = [os.path.abspath(p) for p in args.allow_fs]
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    print("摩尔发票处理服务已启动  http://%s:%d" % (args.host, args.port))
    print("  鉴权: %s" % ("Bearer 令牌已启用" if AUTH_TOKEN else "未启用（仅限内网）"))
    print("  目录操作: %s" % (ALLOWED_ROOTS or "未放行"))
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
        srv.server_close()


if __name__ == "__main__":
    main()
