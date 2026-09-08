FROM python:3.11-slim-bookworm

ARG PIP_INDEX_URL
ARG PIP_DEFAULT_TIMEOUT=120
ENV DEBIAN_FRONTEND=noninteractive PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
RUN sed -i 's|http://deb.debian.org/debian|https://mirrors.aliyun.com/debian|g; s|http://security.debian.org/debian-security|https://mirrors.aliyun.com/debian-security|g' /etc/apt/sources.list.d/debian.sources \
    && apt-get -o Acquire::Retries=3 -o Acquire::https::Timeout=60 update \
    && apt-get install -y --no-install-recommends libreoffice-writer poppler-utils fonts-liberation2 \
    && rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir openpyxl==3.1.5 python-docx==1.2.0 lxml==5.3.2
WORKDIR /app
COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output /data && ln -s /app/data/input_files /data/input_files
