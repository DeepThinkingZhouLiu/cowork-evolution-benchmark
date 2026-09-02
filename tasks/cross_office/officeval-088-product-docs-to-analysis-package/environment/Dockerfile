FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN sed -i \
    -e 's|http://deb.debian.org/debian-security|http://mirrors.aliyun.com/debian-security|g' \
    -e 's|http://deb.debian.org/debian|http://mirrors.aliyun.com/debian|g' \
    /etc/apt/sources.list.d/debian.sources \
    && apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-writer \
       libreoffice-calc \
       poppler-utils \
       tesseract-ocr \
       tesseract-ocr-chi-sim \
       fonts-liberation2 \
       fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir --index-url https://mirrors.aliyun.com/pypi/simple \
    openpyxl==3.1.5 \
    python-docx==1.1.2 \
    Pillow==11.1.0

WORKDIR /app
COPY input_files /app/input_files
RUN mkdir -p /app/output /logs/artifacts /logs/verifier
