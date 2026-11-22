# 自建部署（任意云主机 / K8s）用。WorkBuddy「发布为应用」不需要此文件。
FROM python:3.11-slim

WORKDIR /app

# 只装核心依赖：PDF 解析。OCR 体积过大，按需解开下面两行。
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt \
    -i https://mirrors.cloud.tencent.com/pypi/simple

COPY scripts/ ./scripts/
COPY README.md ./

ENV PORT=8080
# 单端口 HTTP 服务，必须绑定 0.0.0.0
CMD ["python", "-u", "scripts/serve.py", "--host", "0.0.0.0", "--port", "8080"]
