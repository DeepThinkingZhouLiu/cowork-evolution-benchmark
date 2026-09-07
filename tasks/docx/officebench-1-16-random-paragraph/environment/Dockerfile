FROM python:3.11-slim-bookworm

ARG PIP_INDEX_URL
ARG PIP_DEFAULT_TIMEOUT=120
ENV DEBIAN_FRONTEND=noninteractive PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
# 原始 Debian HTTP 镜像在正式 Harbor 冷构建中连续 30 分钟无法完成包下载；
# 换用同步的 HTTPS 镜像并保留重试，避免把基础设施超时误认为任务失败。
RUN sed -ri 's|http://deb.debian.org|https://mirrors.aliyun.com|g' /etc/apt/sources.list.d/debian.sources \
    && apt-get -o Acquire::Retries=5 -o Acquire::https::Timeout=120 update \
    && apt-get -o Acquire::Retries=5 -o Acquire::https::Timeout=120 install -y --no-install-recommends libreoffice-writer poppler-utils fonts-liberation2 \
    && rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir python-docx==1.1.2
WORKDIR /app
COPY input_files/ /app/input_files/
RUN mkdir -p /app/output /logs/artifacts /logs/verifier /data && ln -s /app/input_files /data/input_files
