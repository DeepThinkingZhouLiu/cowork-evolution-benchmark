FROM python:3.11-slim-bookworm

ARG PIP_INDEX_URL
ARG PIP_DEFAULT_TIMEOUT=120
ENV DEBIAN_FRONTEND=noninteractive PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
RUN sed -ri 's|http://deb.debian.org|https://mirrors.aliyun.com|g' /etc/apt/sources.list.d/debian.sources \
    && apt-get -o Acquire::Retries=5 -o Acquire::https::Timeout=120 update \
    && apt-get -o Acquire::Retries=5 -o Acquire::https::Timeout=120 install -y --no-install-recommends libreoffice-writer fonts-noto-cjk fonts-liberation2 \
    && rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir python-docx==1.1.2 lxml==5.4.0 openpyxl==3.1.5 Pillow==11.2.1
WORKDIR /app
COPY input_files/ /app/input_files/
RUN mkdir -p /app/output /logs/artifacts /logs/verifier /data && ln -s /app/input_files /data/input_files
