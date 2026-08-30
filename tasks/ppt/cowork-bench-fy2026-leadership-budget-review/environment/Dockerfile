FROM python:3.11-slim-bookworm

ARG DEBIAN_MIRROR=http://mirrors.tuna.tsinghua.edu.cn/debian
ARG DEBIAN_SECURITY_MIRROR=http://mirrors.tuna.tsinghua.edu.cn/debian-security

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN sed -i \
       -e "s|http://deb.debian.org/debian-security|${DEBIAN_SECURITY_MIRROR}|g" \
       -e "s|http://deb.debian.org/debian|${DEBIAN_MIRROR}|g" \
       /etc/apt/sources.list.d/debian.sources \
    && apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-calc \
       libreoffice-impress \
       poppler-utils \
       fonts-dejavu-core \
       fonts-liberation \
       fonts-crosextra-carlito \
       fonts-crosextra-caladea \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
    --index-url https://pypi.tuna.tsinghua.edu.cn/simple \
    lxml==6.0.1 \
    openpyxl==3.1.5 \
    Pillow==11.3.0 \
    python-pptx==1.0.2

WORKDIR /app
COPY input_files /app/input_files
