# -*- coding: utf-8 -*-
"""ima 知识库入库上传器：按 create_media 返回的 COS 临时凭证上传文件（纯 stdlib，无新依赖）

用法:
    python ima_upload.py <凭据JSON路径>
凭据JSON 格式（由 ima create_media 返回）:
    [{"file": "绝对路径", "media_id": "...", "cos_credential": {...}}, ...]
"""
import sys, os, json, hmac, hashlib, time
import urllib.request
import urllib.parse


def hmac_sha1(key: bytes, msg: str) -> str:
    return hmac.new(key, msg.encode('utf-8'), hashlib.sha1).hexdigest()


def sha1_hex(msg: str) -> str:
    return hashlib.sha1(msg.encode('utf-8')).hexdigest()


def qencode(v: str) -> str:
    return urllib.parse.quote(v, safe='')


def cos_put(local_path, cred):
    c = cred['cos_credential']
    secret_id = c['secret_id']
    secret_key = c['secret_key'].encode('utf-8')
    token = c.get('token', '')
    start, end = str(c['start_time']), str(c['expired_time'])
    key_time = '%s;%s' % (start, end)
    bucket = c['bucket_name']
    region = c['region']
    cos_key = c['cos_key']

    host = '%s.cos.%s.myqcloud.com' % (bucket, region)
    url = 'https://%s/%s' % (host, cos_key)

    data = open(local_path, 'rb').read()
    ctype = 'text/markdown'
    if local_path.lower().endswith('.pptx'):
        ctype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

    # 参与签名的 header（小写，按字典序）
    headers_sign = {'content-type': ctype, 'host': host}
    hl = ';'.join(sorted(headers_sign.keys()))
    http_headers = '&'.join('%s=%s' % (k, qencode(headers_sign[k])) for k in sorted(headers_sign))

    key_name = '/' + cos_key
    http_string = 'put\n%s\n%s\n%s\n' % (key_name, '', http_headers)
    string_to_sign = 'sha1\n%s\n%s\n' % (key_time, sha1_hex(http_string))
    sign_key = hmac_sha1(secret_key, key_time)
    signature = hmac_sha1(sign_key.encode('utf-8'), string_to_sign)

    auth = ('q-sign-algorithm=sha1&q-ak=%s&q-sign-time=%s&q-key-time=%s'
            '&q-header-list=%s&q-url-param-list=&q-signature=%s'
            % (secret_id, key_time, key_time, hl, signature))

    req = urllib.request.Request(url, data=data, method='PUT')
    req.add_header('Host', host)
    req.add_header('Content-Type', ctype)
    req.add_header('Authorization', auth)
    if token:
        req.add_header('x-cos-security-token', token)
    req.add_header('Content-Length', str(len(data)))

    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.status, resp.read().decode('utf-8', 'ignore')


def main():
    path = sys.argv[1]
    items = json.load(open(path, encoding='utf-8'))
    for it in items:
        try:
            st, body = cos_put(it['file'], it)
            print('%-6s %s  %s' % (st, os.path.basename(it['file']), it['media_id']))
            if st >= 400:
                print('   body:', body[:400])
        except Exception as e:
            print('FAIL  %s  %s' % (os.path.basename(it['file']), e))
            if hasattr(e, 'read'):
                try:
                    print('   body:', e.read().decode('utf-8', 'ignore')[:600])
                except Exception:
                    pass


if __name__ == '__main__':
    main()
